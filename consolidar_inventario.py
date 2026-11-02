#!/usr/bin/env python3
"""Consolida inventarios CSV/ZIP en un XLSX usando SQLite como almacenamiento temporal."""

from __future__ import annotations

import csv
import io
import itertools
import logging
import unicodedata
import json
import os
import re
import shutil
import sqlite3
import sys
import time
import zipfile
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Iterator, TextIO

import tempfile
from xml.sax.saxutils import escape, quoteattr


SCRIPT_VERSION = "2026.11.01.31-sla-en-plazo"
CONFIG_FILENAME = "inventario_config.json"
PARTNERS_CONFIG_FILENAME = "partners_config.json"
DATABASE_FILENAME = "inventario_consolidacion.sqlite"
BATCH_SIZE = 5_000
PROGRESS_EVERY = 25_000
EXCEL_MAX_ROWS = 1_048_576


def configure_csv_field_limit() -> int:
    """Amplía el máximo de caracteres por campo CSV hasta el límite admitido."""
    limit = sys.maxsize
    while True:
        try:
            csv.field_size_limit(limit)
            return limit
        except OverflowError:
            limit //= 10


CSV_FIELD_SIZE_LIMIT = configure_csv_field_limit()

SOURCE_FIELDS = [
    "title", "last_detection", "state", "found_by", "custom_risk", "id",
    "finding_uuid", "date", "planned_resolution_date", "cwe", "cve",
    "last_reviewed", "created", "modified", "cappm", "date_deadline",
    "date_paused", "date_closed", "date_active", "ticketing_tool_key", "tags",
    "endpoint__product__name", "endpoint__product__plataforma",
    "endpoint__product__grupo", "endpoint__product__servicio",
    "endpoint__product__main_provider", "data_from",
]
SLA_FIELDS = ["Cumple_SLA", "Mes_inicio_incumplimiento", "Mes_fin_incumplimiento", "Mes_cumplimiento"]
OUTPUT_FIELDS = SOURCE_FIELDS + ["framework", "perimetro_agente", "Ejecución_agente", "resultado_agente", "fecha_ult_ejecucion_agente", "partner", "APLICACION_BAJA"] + SLA_FIELDS
DATE_FIELDS = {
    "last_detection", "date", "planned_resolution_date", "last_reviewed", "created",
    "modified", "date_deadline", "date_paused", "date_closed", "date_active", "data_from",
    "fecha_ult_ejecucion_agente",
    "Mes_inicio_incumplimiento", "Mes_fin_incumplimiento", "Mes_cumplimiento",
}
TEXT_FIELDS = {"found_by", "cwe", "finding_uuid"}
UUID_RE = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$", re.IGNORECASE)
FRAMEWORK_RE = re.compile(r"^Application\s*-\s*([^.]+)\.")


class ConsolidationError(Exception):
    """Error controlado con mensaje apto para consola."""


def script_directory() -> Path:
    return Path(__file__).resolve().parent


def load_config(config_path: Path) -> dict[str, Any]:
    if not config_path.is_file():
        raise ConsolidationError(f"No existe el archivo de configuración: {config_path}")
    try:
        with config_path.open("r", encoding="utf-8-sig") as handle:
            config = json.load(handle)
    except json.JSONDecodeError as exc:
        raise ConsolidationError(
            f"JSON inválido en {config_path} (línea {exc.lineno}, columna {exc.colno}): {exc.msg}"
        ) from exc
    except OSError as exc:
        raise ConsolidationError(f"No se puede leer la configuración {config_path}: {exc}") from exc

    required = {"input_folder", "output_file", "prefixes", "sheet_name", "table_name", "keep_database"}
    missing = sorted(required - set(config))
    if missing:
        raise ConsolidationError(f"Faltan claves obligatorias en la configuración: {', '.join(missing)}")
    if not isinstance(config["prefixes"], list) or not config["prefixes"] or not all(
        isinstance(item, str) and item.strip() for item in config["prefixes"]
    ):
        raise ConsolidationError("'prefixes' debe ser una lista no vacía de textos.")
    if not isinstance(config["keep_database"], bool):
        raise ConsolidationError("'keep_database' debe ser true o false.")
    if "fw_agents" not in config or not isinstance(config["fw_agents"], dict):
        raise ConsolidationError("'fw_agents' debe existir y ser un objeto con framework y fecha.")
    normalized_fw_agents: dict[str, str] = {}
    for framework, cutoff in config["fw_agents"].items():
        if not isinstance(framework, str) or not framework.strip():
            raise ConsolidationError("Las claves de 'fw_agents' deben ser textos no vacíos.")
        if not isinstance(cutoff, str) or not cutoff.strip():
            raise ConsolidationError(f"La fecha de 'fw_agents.{framework}' debe ser un texto no vacío.")
        parsed_cutoff = parse_datetime(cutoff)
        if parsed_cutoff is None:
            raise ConsolidationError(
                f"Fecha inválida en 'fw_agents.{framework}': {cutoff}. Use DD/MM/YYYY o una fecha ISO válida."
            )
        normalized_fw_agents[framework.strip().casefold()] = datetime_to_storage(parsed_cutoff)
    config["_fw_agents_normalized"] = normalized_fw_agents
    for key in ("input_folder", "output_file", "sheet_name", "table_name"):
        if not isinstance(config[key], str) or not config[key].strip():
            raise ConsolidationError(f"'{key}' debe ser un texto no vacío.")
    return config


def load_partners_config(config_path: Path) -> dict[str, str]:
    if not config_path.is_file():
        raise ConsolidationError(f"No existe el archivo de partners: {config_path}")
    try:
        with config_path.open("r", encoding="utf-8-sig") as handle:
            config = json.load(handle)
    except json.JSONDecodeError as exc:
        raise ConsolidationError(
            f"JSON inválido en {config_path} (línea {exc.lineno}, columna {exc.colno}): {exc.msg}"
        ) from exc
    except OSError as exc:
        raise ConsolidationError(f"No se puede leer {config_path}: {exc}") from exc
    mapping = config.get("mapping")
    if not isinstance(mapping, dict) or not mapping:
        raise ConsolidationError("partners_config.json debe contener un objeto 'mapping' no vacío.")
    normalized: dict[str, str] = {}
    for provider, partner in mapping.items():
        if not isinstance(provider, str) or not provider.strip():
            raise ConsolidationError("Las denominaciones de Main Provider deben ser textos no vacíos.")
        if not isinstance(partner, str) or not partner.strip():
            raise ConsolidationError(f"El partner agrupado de '{provider}' debe ser un texto no vacío.")
        key = provider.strip().casefold()
        if key in normalized and normalized[key] != partner.strip():
            raise ConsolidationError(f"Main Provider duplicado con agrupaciones distintas: {provider}")
        normalized[key] = partner.strip()
    return normalized


def select_latest_files(input_folder: Path, prefixes: list[str]) -> list[tuple[str, Path]]:
    try:
        candidates = [p for p in input_folder.iterdir() if p.is_file() and p.suffix.lower() in {".csv", ".zip"}]
    except OSError as exc:
        raise ConsolidationError(f"No se puede listar la carpeta de entrada {input_folder}: {exc}") from exc

    selected: list[tuple[str, Path]] = []
    for prefix in prefixes:
        matches = [p for p in candidates if p.name.lower().startswith(prefix.lower())]
        if not matches:
            raise ConsolidationError(
                f"No se encontró ningún archivo CSV o ZIP para el prefijo: {prefix}"
            )
        try:
            latest = max(matches, key=lambda p: (p.stat().st_mtime_ns, p.name.lower()))
        except OSError as exc:
            raise ConsolidationError(f"No se puede consultar la fecha de los archivos de '{prefix}': {exc}") from exc
        selected.append((prefix, latest))
    return selected


def largest_csv_in_zip(zip_path: Path) -> zipfile.ZipInfo:
    try:
        with zipfile.ZipFile(zip_path, "r") as archive:
            csv_members = [info for info in archive.infolist() if not info.is_dir() and info.filename.lower().endswith(".csv")]
            if not csv_members:
                raise ConsolidationError(f"El ZIP no contiene ningún CSV: {zip_path}")
            return max(csv_members, key=lambda info: (info.file_size, info.filename.lower()))
    except zipfile.BadZipFile as exc:
        raise ConsolidationError(f"El archivo ZIP está dañado o no es válido: {zip_path}") from exc
    except OSError as exc:
        raise ConsolidationError(f"No se puede abrir el ZIP {zip_path}: {exc}") from exc


@contextmanager
def open_csv_stream(path: Path) -> Iterator[tuple[TextIO, str]]:
    """Abre CSV normal o el CSV más grande de un ZIP, siempre en streaming y con newline=''."""
    if path.suffix.lower() == ".zip":
        try:
            archive = zipfile.ZipFile(path, "r")
            csv_members = [info for info in archive.infolist() if not info.is_dir() and info.filename.lower().endswith(".csv")]
            if not csv_members:
                archive.close()
                raise ConsolidationError(f"El ZIP no contiene ningún CSV: {path}")
            member = max(csv_members, key=lambda info: (info.file_size, info.filename.lower()))
            binary = archive.open(member, "r")
            text = io.TextIOWrapper(binary, encoding="utf-8-sig", errors="strict", newline="")
            try:
                yield text, f"{path}::{member.filename}"
            finally:
                text.close()
                archive.close()
        except zipfile.BadZipFile as exc:
            raise ConsolidationError(f"El archivo ZIP está dañado o no es válido: {path}") from exc
        except UnicodeDecodeError as exc:
            raise ConsolidationError(f"Codificación incorrecta en el CSV contenido en {path}; se esperaba UTF-8: {exc}") from exc
        except OSError as exc:
            raise ConsolidationError(f"No se puede leer el ZIP {path}: {exc}") from exc
    else:
        try:
            with path.open("r", encoding="utf-8-sig", errors="strict", newline="") as text:
                yield text, str(path)
        except UnicodeDecodeError as exc:
            raise ConsolidationError(f"Codificación incorrecta en {path}; se esperaba UTF-8: {exc}") from exc
        except OSError as exc:
            raise ConsolidationError(f"No se puede leer el CSV {path}: {exc}") from exc


def parse_datetime(value: Any) -> datetime | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None

    iso_candidate = text
    if iso_candidate.endswith("Z"):
        iso_candidate = iso_candidate[:-1] + "+00:00"
    try:
        parsed = datetime.fromisoformat(iso_candidate)
        return normalize_timezone(parsed)
    except ValueError:
        pass

    formats = (
        "%d/%m/%Y %H:%M:%S.%f", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M",
        "%d/%m/%Y", "%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M", "%Y-%m-%d",
    )
    for fmt in formats:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def normalize_timezone(value: datetime) -> datetime:
    if value.tzinfo is not None:
        return value.astimezone(timezone.utc).replace(tzinfo=None)
    return value


def datetime_to_storage(value: datetime) -> str:
    return value.isoformat(timespec="microseconds")


def parse_decimal(value: Any) -> float | None:
    if value is None:
        return None
    text = str(value).strip().replace(" ", "")
    if not text:
        return None
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    else:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def calculate_sla(custom_risk: Any, date_active_value: Any, date_closed_value: Any) -> tuple[str, str, str, str]:
    """Devuelve estado SLA y meses de incumplimiento/cumplimiento."""
    date_active = parse_datetime(date_active_value)
    if date_active is None:
        return "N/A", "", "", ""

    risk = parse_decimal(custom_risk)
    sla_days = 30 if risk is not None and risk in {8, 9, 10} else 90
    deadline = date_active + timedelta(days=sla_days)
    date_closed = parse_datetime(date_closed_value)
    raw_closed = "" if date_closed_value is None else str(date_closed_value).strip()
    closed = date_closed is not None and bool(raw_closed)
    if closed and date_closed <= deadline:
        return "SI", "", "", datetime(date_closed.year, date_closed.month, 1).isoformat()

    comparison_date = date_closed if closed else datetime.now()
    if comparison_date <= deadline:
        return "En plazo", "", "", ""

    breach_start = (deadline.date() + timedelta(days=1))
    start_month = datetime(breach_start.year, breach_start.month, 1).isoformat()
    end_month = (
        datetime(date_closed.year, date_closed.month, 1).isoformat()
        if closed else "9999-12-31T00:00:00"
    )
    return "NO", start_month, end_month, ""


def parse_integer_when_possible(value: Any) -> int | str:
    text = "" if value is None else str(value).strip()
    if not text:
        return ""
    if re.fullmatch(r"[+-]?\d+", text):
        try:
            return int(text)
        except ValueError:
            pass
    if re.fullmatch(r"[+-]?\d+\.0+", text):
        try:
            return int(float(text))
        except (ValueError, OverflowError):
            pass
    return text


def prepare_row(row: dict[str, Any], fw_agents: dict[str, str], partners: dict[str, str], applications: dict[str, bool]) -> tuple[list[Any], str] | None:
    raw_uuid = "" if row.get("finding_uuid") is None else str(row["finding_uuid"]).strip()
    if not UUID_RE.fullmatch(raw_uuid):
        return None
    uuid_value = raw_uuid.lower()

    values: list[Any] = []
    sort_key = ""
    for field in SOURCE_FIELDS:
        raw = row.get(field)
        text = "" if raw is None else str(raw).strip()
        if field == "finding_uuid":
            value: Any = uuid_value
        elif field == "custom_risk":
            value = parse_decimal(raw)
        elif field == "id":
            value = parse_integer_when_possible(raw)
        elif field in DATE_FIELDS:
            parsed = parse_datetime(raw)
            value = datetime_to_storage(parsed) if parsed is not None else text
            if field == "data_from":
                sort_key = datetime_to_storage(parsed) if parsed is not None else ""
        elif field in TEXT_FIELDS:
            value = text
        else:
            value = text
        values.append(value)

    product_name = "" if row.get("endpoint__product__name") is None else str(row["endpoint__product__name"]).strip()
    match = FRAMEWORK_RE.match(product_name)
    framework = match.group(1).strip() if match else ""
    values.append(framework)

    perimetro_agente = ""
    cutoff_stored = fw_agents.get(framework.casefold()) if framework else None
    if cutoff_stored is not None:
        cutoff = parse_datetime(cutoff_stored)
        date_active = parse_datetime(row.get("date_active"))
        raw_date_closed = "" if row.get("date_closed") is None else str(row.get("date_closed")).strip()
        date_closed = parse_datetime(raw_date_closed)
        in_scope = False
        if cutoff is not None and date_active is not None:
            in_scope = (
                date_active >= cutoff
                or (date_active < cutoff and date_closed is not None and date_closed > cutoff)
                or (date_active < cutoff and raw_date_closed == "")
            )
        perimetro_agente = "SI" if in_scope else "NO"
    values.append(perimetro_agente)
    values.extend(["NO", "", ""])
    main_provider = "" if row.get("endpoint__product__main_provider") is None else str(row.get("endpoint__product__main_provider")).strip()
    # Si no hay agrupación explícita, se conserva la denominación original.
    partner = partners.get(main_provider.casefold(), main_provider) if main_provider else ""
    values.append(partner)

    application_name = "" if row.get("endpoint__product__name") is None else str(row.get("endpoint__product__name")).strip()
    ide_active, _matched_key, _match_method = resolve_ide00838_application(application_name, applications)
    # Regla: no encontrado = SI; encontrado activo = NO; encontrado no activo = SI.
    aplicacion_baja = "NO" if ide_active is True else "SI"
    values.append(aplicacion_baja)
    values.extend(calculate_sla(row.get("custom_risk"), row.get("date_active"), row.get("date_closed")))
    return values, sort_key


def quote_identifier(identifier: str) -> str:
    return '"' + identifier.replace('"', '""') + '"'


def create_database(db_path: Path) -> tuple[sqlite3.Connection, str]:
    try:
        if db_path.exists():
            db_path.unlink()
        connection = sqlite3.connect(db_path)
        connection.execute("PRAGMA journal_mode=WAL")
        connection.execute("PRAGMA synchronous=NORMAL")
        definitions = []
        for field in OUTPUT_FIELDS:
            if field == "finding_uuid":
                definitions.append(f"{quote_identifier(field)} TEXT PRIMARY KEY")
            elif field == "custom_risk":
                definitions.append(f"{quote_identifier(field)} REAL")
            else:
                definitions.append(f"{quote_identifier(field)}")
        definitions.append('"_sort_key" TEXT NOT NULL')
        connection.execute(f"CREATE TABLE inventario ({', '.join(definitions)})")

        columns = OUTPUT_FIELDS + ["_sort_key"]
        quoted = ", ".join(quote_identifier(c) for c in columns)
        placeholders = ", ".join("?" for _ in columns)
        assignments = ", ".join(
            f"{quote_identifier(c)} = excluded.{quote_identifier(c)}"
            for c in columns if c != "finding_uuid"
        )
        sql = (
            f"INSERT INTO inventario ({quoted}) VALUES ({placeholders}) "
            f"ON CONFLICT(finding_uuid) DO UPDATE SET {assignments} "
            f"WHERE excluded._sort_key > inventario._sort_key"
        )
        return connection, sql
    except (OSError, sqlite3.Error) as exc:
        raise ConsolidationError(f"No se puede crear la base temporal {db_path}: {exc}") from exc


def flush_batch(connection: sqlite3.Connection, sql: str, batch: list[tuple[Any, ...]]) -> int:
    if not batch:
        return 0
    before = connection.total_changes
    try:
        connection.executemany(sql, batch)
        connection.commit()
    except sqlite3.Error as exc:
        connection.rollback()
        raise ConsolidationError(f"Error al insertar o actualizar la base temporal: {exc}") from exc
    changed = connection.total_changes - before
    batch.clear()
    return changed


def process_file(connection: sqlite3.Connection, sql: str, path: Path, fw_agents: dict[str, str], partners: dict[str, str], applications: dict[str, bool]) -> dict[str, int]:
    stats = {"read": 0, "valid": 0, "discarded": 0, "changed": 0}
    batch: list[tuple[Any, ...]] = []

    with open_csv_stream(path) as (stream, source_name):
        print(f"  CSV utilizado: {source_name}")
        reader = csv.DictReader(
            stream,
            delimiter=";",
            quotechar='"',
            doublequote=True,
        )
        if reader.fieldnames is None:
            raise ConsolidationError(f"El CSV no contiene cabecera: {source_name}")
        normalized_headers = [header.strip() if header is not None else "" for header in reader.fieldnames]
        reader.fieldnames = normalized_headers
        missing = [field for field in SOURCE_FIELDS if field not in normalized_headers]
        if missing:
            raise ConsolidationError(
                f"Columnas obligatorias ausentes en {source_name}: {', '.join(missing)}"
            )

        try:
            for row in reader:
                stats["read"] += 1
                prepared = prepare_row(row, fw_agents, partners, applications)
                if prepared is None:
                    stats["discarded"] += 1
                else:
                    values, sort_key = prepared
                    stats["valid"] += 1
                    batch.append(tuple(values + [sort_key]))
                    if len(batch) >= BATCH_SIZE:
                        stats["changed"] += flush_batch(connection, sql, batch)

                if stats["read"] % PROGRESS_EVERY == 0:
                    print(
                        f"    Progreso: {stats['read']:,} leídos | {stats['valid']:,} UUID válidos | "
                        f"{stats['discarded']:,} descartados | {stats['changed']:,} inserciones/actualizaciones"
                    )
        except csv.Error as exc:
            physical_line = getattr(reader, "line_num", "desconocida")
            raise ConsolidationError(
                f"CSV mal formado en {source_name}, cerca del registro lógico "
                f"{stats['read'] + 1} y línea física {physical_line}: {exc}. "
                f"Límite de tamaño por campo configurado: {CSV_FIELD_SIZE_LIMIT:,} caracteres."
            ) from exc
        except UnicodeDecodeError as exc:
            raise ConsolidationError(f"Codificación incorrecta en {source_name}; se esperaba UTF-8: {exc}") from exc

    stats["changed"] += flush_batch(connection, sql, batch)
    return stats


def select_ide00838_file(input_folder: Path) -> Path:
    candidates = [
        path for path in input_folder.iterdir()
        if path.is_file()
        and path.suffix.casefold() in {".csv", ".zip"}
        and path.name.casefold().startswith("ide00838")
    ]
    if not candidates:
        raise ConsolidationError(
            "No se encontró ningún fichero IDE00838*.CSV o .zip en la carpeta de datos."
        )
    return max(candidates, key=lambda path: (path.stat().st_mtime_ns, path.name.casefold()))


def normalize_application_name(value: Any) -> str:
    """Normaliza nombres para el cruce sin alterar su significado."""
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value))
    text = text.replace("\ufeff", "").replace("\u00a0", " ")
    return " ".join(text.strip().split()).casefold()


def resolve_ide00838_application(application_name: Any, applications: dict[str, bool]) -> tuple[bool | None, str, str]:
    """Busca únicamente por la clave completa normalizada."""
    normalized = normalize_application_name(application_name)
    if not normalized:
        return None, "", "nombre vacío"
    if normalized in applications:
        return applications[normalized], normalized, "clave única BP.D.B"
    return None, normalized, "sin coincidencia"


def load_ide00838_applications(path: Path) -> tuple[dict[str, bool], dict[str, int]]:
    """Carga IDE00838 y compone una única clave: Application - <BP>.<D>.<B>."""
    stats = {"read": 0, "valid_names": 0, "duplicates": 0, "active": 0, "inactive": 0}
    applications: dict[str, bool] = {}

    if path.suffix.casefold() == ".zip":
        with zipfile.ZipFile(path, "r") as archive:
            members = [name for name in archive.namelist() if name.casefold().endswith(".csv")]
            if not members:
                raise ConsolidationError(f"El ZIP {path} no contiene ningún CSV.")
            member = max(members, key=lambda name: archive.getinfo(name).file_size)
            raw = archive.read(member)
            source_name = f"{path}!{member}"
    else:
        raw = path.read_bytes()
        source_name = str(path)

    text = None
    used_encoding = ""
    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = raw.decode(encoding)
            used_encoding = encoding
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise ConsolidationError(f"No se pudo determinar la codificación de {source_name}.")

    reader = csv.reader(io.StringIO(text, newline=""), delimiter=";")
    try:
        first_row = next(reader)
    except StopIteration:
        raise ConsolidationError(f"IDE00838 está vacío: {source_name}")

    normalized_first = [normalize_application_name(value) for value in first_row]
    has_header = "nombre" in normalized_first and "tipo" in normalized_first and "activa" in normalized_first
    if has_header:
        short_name_index = normalized_first.index("nombre")
        type_index = normalized_first.index("tipo")
        bp_index = normalized_first.index("plataforma") if "plataforma" in normalized_first else 67
        active_index = normalized_first.index("activa")
        rows = reader
        header_mode = "con cabecera"
    else:
        # Formato sin cabecera: clave BP.D.B y actividad en K (columna 11).
        # Python usa índices desde cero: B=1, D=3, K=10 y BP=67.
        short_name_index = 1
        type_index = 3
        active_index = 10
        bp_index = 67
        rows = itertools.chain([first_row], reader)
        header_mode = "sin cabecera"

    required_length = max(short_name_index, type_index, active_index, bp_index) + 1
    for row in rows:
        stats["read"] += 1
        if len(row) < required_length:
            continue

        bp_value = row[bp_index].strip()
        short_name = row[short_name_index].strip()
        application_type = row[type_index].strip()
        if not (bp_value and short_name and application_type):
            continue

        key = normalize_application_name(
            f"Application - {bp_value}.{application_type}.{short_name}"
        )
        if not key:
            continue

        stats["valid_names"] += 1
        is_active = row[active_index].strip().upper() == "S"
        if key in applications:
            stats["duplicates"] += 1
            applications[key] = applications[key] or is_active
        else:
            applications[key] = is_active

    stats["active"] = sum(1 for value in applications.values() if value)
    stats["inactive"] = sum(1 for value in applications.values() if not value)
    print(
        f"  IDE00838 utilizado: {source_name} ({header_mode}, codificación {used_encoding}, "
        f"clave única Application - <BP>.<D>.<B>; Activa=columna {active_index + 1})"
    )
    return applications, stats

def select_agent_execution_file(input_folder: Path) -> Path:
    candidates = [p for p in input_folder.iterdir() if p.is_file() and p.suffix.lower() in {".csv", ".zip"} and p.name.casefold().startswith("remediacion_agentica")]
    if not candidates:
        raise ConsolidationError("No se encontró ningún fichero remediacion_agentica*.csv o .zip en la carpeta de datos.")
    return max(candidates, key=lambda p: (p.stat().st_mtime_ns, p.name.casefold()))


def create_agent_table(connection: sqlite3.Connection) -> None:
    connection.execute(
        "CREATE TABLE agent_executions (finding_id TEXT PRIMARY KEY, resolution TEXT, "
        "execution_timestamp TEXT NOT NULL, detail_json TEXT NOT NULL)"
    )
    connection.commit()


def process_agent_execution_file(connection: sqlite3.Connection, path: Path) -> dict[str, int]:
    stats = {"read": 0, "completed": 0, "json_valid": 0, "json_invalid": 0, "findings": 0, "upserts": 0}
    sql = ("INSERT INTO agent_executions (finding_id, resolution, execution_timestamp, detail_json) VALUES (?, ?, ?, ?) "
           "ON CONFLICT(finding_id) DO UPDATE SET resolution=excluded.resolution, "
           "execution_timestamp=excluded.execution_timestamp, detail_json=excluded.detail_json "
           "WHERE excluded.execution_timestamp > agent_executions.execution_timestamp")
    batch: list[tuple[str, str, str, str]] = []
    with open_csv_stream(path) as (stream, source_name):
        print(f"  CSV de ejecuciones utilizado: {source_name}")
        reader = csv.DictReader(stream, delimiter=",", quotechar='"', doublequote=True)
        if reader.fieldnames is None:
            raise ConsolidationError(f"El fichero de ejecuciones no contiene cabecera: {source_name}")
        reader.fieldnames = [name.strip() if name else "" for name in reader.fieldnames]
        missing = sorted({"timestamp", "status", "details"} - set(reader.fieldnames))
        if missing:
            raise ConsolidationError(f"Columnas ausentes en {source_name}: {', '.join(missing)}")
        try:
            for row in reader:
                stats["read"] += 1
                if (row.get("status") or "").strip().casefold() != "completed":
                    continue
                stats["completed"] += 1
                timestamp = parse_datetime(row.get("timestamp"))
                details_text = (row.get("details") or "").strip()
                if timestamp is None or not details_text:
                    continue
                try:
                    details = json.loads(details_text)
                except (json.JSONDecodeError, TypeError):
                    stats["json_invalid"] += 1
                    continue
                stats["json_valid"] += 1
                findings = details.get("findings")
                if not isinstance(findings, list):
                    continue
                timestamp_stored = datetime_to_storage(timestamp)
                for finding in findings:
                    if not isinstance(finding, dict):
                        continue
                    finding_id = str(finding.get("finding_id", "")).strip()
                    if not finding_id:
                        continue
                    resolution = str(finding.get("resolution", "") or "").strip()
                    detail = {
                        "execution": {key: row.get(key, "") for key in (
                            "id", "agent_name", "user_id", "action", "timestamp", "status",
                            "duration_ms", "repository",
                        )},
                        "framework": details.get("framework", "") if isinstance(details, dict) else "",
                        "project": details.get("project", "") if isinstance(details, dict) else "",
                        "finding": finding,
                    }
                    batch.append((finding_id, resolution, timestamp_stored,
                                  json.dumps(detail, ensure_ascii=False, separators=(",", ":"))))
                    stats["findings"] += 1
                    if len(batch) >= BATCH_SIZE:
                        before = connection.total_changes
                        connection.executemany(sql, batch)
                        connection.commit()
                        stats["upserts"] += connection.total_changes - before
                        batch.clear()
                if stats["read"] % PROGRESS_EVERY == 0:
                    print(f"    Ejecuciones: {stats['read']:,} registros leídos | {stats['findings']:,} findings extraídos")
        except csv.Error as exc:
            raise ConsolidationError(f"CSV de ejecuciones mal formado en {source_name}, línea física {reader.line_num}: {exc}") from exc
    if batch:
        before = connection.total_changes
        connection.executemany(sql, batch)
        connection.commit()
        stats["upserts"] += connection.total_changes - before
    return stats


def enrich_inventory_with_agent_data(connection: sqlite3.Connection) -> tuple[int, int]:
    total_agents = connection.execute("SELECT COUNT(*) FROM agent_executions").fetchone()[0]
    sql = """UPDATE inventario
             SET "Ejecución_agente"='SI',
                 "resultado_agente"=(SELECT a.resolution FROM agent_executions a WHERE a.finding_id=CAST(inventario."id" AS TEXT)),
                 "fecha_ult_ejecucion_agente"=(SELECT a.execution_timestamp FROM agent_executions a WHERE a.finding_id=CAST(inventario."id" AS TEXT))
             WHERE EXISTS (SELECT 1 FROM agent_executions a WHERE a.finding_id=CAST(inventario."id" AS TEXT))"""
    connection.execute(sql)
    connection.commit()
    matched = connection.execute("SELECT COUNT(*) FROM inventario WHERE \"Ejecución_agente\"='SI'").fetchone()[0]
    return total_agents, matched


def write_unmatched_agent_findings_csv(
    connection: sqlite3.Connection, output_path: Path
) -> tuple[Path, int]:
    """Exporta findings de agente que no cruzan con id del inventario."""
    # La expresión coincide con la condición de NOT EXISTS de abajo. Sin este
    # índice SQLite puede recorrer todo el inventario por cada finding agente.
    connection.execute(
        'CREATE INDEX IF NOT EXISTS idx_inventario_id_text '
        'ON inventario(CAST("id" AS TEXT))'
    )
    logs_dir = output_path.parent / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    detail_path = logs_dir / f"ejecuciones_agente_sin_inventario_{timestamp}.csv"
    latest_path = logs_dir / "ejecuciones_agente_sin_inventario_ULTIMO.csv"
    query = """SELECT a.finding_id, a.resolution, a.execution_timestamp, a.detail_json
               FROM agent_executions a
               WHERE NOT EXISTS (
                   SELECT 1 FROM inventario i
                   WHERE a.finding_id=CAST(i."id" AS TEXT)
               )
               ORDER BY a.finding_id"""
    count = 0
    with detail_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle, delimiter=";")
        writer.writerow([
            "finding_id", "resolution", "timestamp", "execution_id", "agent_name",
            "action", "status", "duration_ms", "repository", "framework", "project",
            "severity", "product", "summary", "detalle_finding_json",
        ])
        for finding_id, resolution, execution_timestamp, detail_text in connection.execute(query):
            detail = json.loads(detail_text)
            execution = detail.get("execution", {})
            finding = detail.get("finding", {})
            writer.writerow([
                finding_id, resolution, execution_timestamp,
                execution.get("id", ""), execution.get("agent_name", ""),
                execution.get("action", ""), execution.get("status", ""),
                execution.get("duration_ms", ""), execution.get("repository", ""),
                detail.get("framework", ""),
                finding.get("project") or detail.get("project", ""),
                finding.get("severity", ""), finding.get("product", ""),
                finding.get("summary", ""),
                json.dumps(finding, ensure_ascii=False, separators=(",", ":")),
            ])
            count += 1
    shutil.copyfile(detail_path, latest_path)
    return detail_path, count


def estimate_uncompressed_input_size(selected: list[tuple[str, Path]]) -> int:
    total = 0
    for _, path in selected:
        if path.suffix.lower() == ".zip":
            total += largest_csv_in_zip(path).file_size
        else:
            total += path.stat().st_size
    return total


def check_disk_space(folder: Path, estimated_input_size: int) -> None:
    try:
        free = shutil.disk_usage(folder).free
    except OSError as exc:
        raise ConsolidationError(f"No se puede comprobar el espacio libre en {folder}: {exc}") from exc
    required = max(200 * 1024 * 1024, int(estimated_input_size * 1.5))
    if free < required:
        raise ConsolidationError(
            f"Espacio en disco insuficiente en {folder}. Libre: {free / 1024**3:.2f} GB; "
            f"estimación mínima requerida: {required / 1024**3:.2f} GB."
        )


def temp_output_path(output_path: Path) -> Path:
    return output_path.with_name(f"{output_path.stem}.tmp{output_path.suffix}")


def excel_column_letter(index: int) -> str:
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def excel_date_serial(value: datetime) -> float:
    epoch = datetime(1899, 12, 30)
    delta = value - epoch
    return delta.days + (delta.seconds + delta.microseconds / 1_000_000) / 86_400


def xml_text(value: Any) -> str:
    text = "" if value is None else str(value)
    text = "".join(ch for ch in text if ch in "\t\n\r" or ord(ch) >= 32)
    return escape(text)


def numeric_xml(value: int | float) -> str:
    return format(value, ".15g") if isinstance(value, float) else str(value)


def cell_xml(reference: str, field: str, value: Any, header: bool = False) -> str:
    if value is None or value == "":
        return f'<c r="{reference}" s="{1 if header else 0}"/>'
    if header:
        return f'<c r="{reference}" s="1" t="inlineStr"><is><t>{xml_text(value)}</t></is></c>'
    if field in DATE_FIELDS:
        parsed = parse_datetime(value)
        if parsed is not None:
            return f'<c r="{reference}" s="2"><v>{numeric_xml(excel_date_serial(parsed))}</v></c>'
    if field in TEXT_FIELDS:
        return f'<c r="{reference}" s="3" t="inlineStr"><is><t>{xml_text(value)}</t></is></c>'
    if field == "custom_risk" and isinstance(value, (int, float)):
        # Se asigna explícitamente el estilo 0 (General). El valor se guarda como
        # número real; Excel mostrará 5 como 5 y 5.5 como 5,5 en configuración ES.
        # La asignación explícita evita conservar/aplicar un formato decimal previo.
        numeric_value = int(value) if isinstance(value, float) and value.is_integer() else value
        return f'<c r="{reference}" s="0"><v>{numeric_xml(numeric_value)}</v></c>'
    if field == "id" and isinstance(value, int):
        return f'<c r="{reference}"><v>{value}</v></c>'
    if isinstance(value, (int, float)):
        return f'<c r="{reference}"><v>{numeric_xml(value)}</v></c>'
    return f'<c r="{reference}" t="inlineStr"><is><t>{xml_text(value)}</t></is></c>'


def write_worksheet_xml(connection: sqlite3.Connection, xml_path: Path, unique_count: int) -> int:
    last_row = unique_count + 1
    widths = {1: 55, 7: 38, 22: 45, 23: 28, 24: 28, 25: 28, 26: 30, 27: 24, 28: 18, 29: 18, 30: 18, 31: 24, 32: 26, 33: 24, 34: 20, 35: 16, 36: 26, 37: 26, 38: 22}
    quoted_columns = ", ".join(quote_identifier(field) for field in OUTPUT_FIELDS)
    cursor = connection.execute(f"SELECT {quoted_columns} FROM inventario ORDER BY rowid")
    with xml_path.open("w", encoding="utf-8", newline="") as handle:
        handle.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        handle.write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
        last_column = excel_column_letter(len(OUTPUT_FIELDS))
        handle.write(f'<dimension ref="A1:{last_column}{last_row}"/>')
        handle.write('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>')
        handle.write('<sheetFormatPr defaultRowHeight="15"/><cols>')
        for index in range(1, len(OUTPUT_FIELDS) + 1):
            handle.write(f'<col min="{index}" max="{index}" width="{widths.get(index, 18)}" customWidth="1"/>')
        handle.write('</cols><sheetData>')
        cells = "".join(cell_xml(f"{excel_column_letter(i)}1", field, field, True) for i, field in enumerate(OUTPUT_FIELDS, 1))
        handle.write(f'<row r="1">{cells}</row>')
        written = 0
        for db_row in cursor:
            row_number = written + 2
            cells = "".join(cell_xml(f"{excel_column_letter(i)}{row_number}", field, value) for i, (field, value) in enumerate(zip(OUTPUT_FIELDS, db_row), 1))
            handle.write(f'<row r="{row_number}">{cells}</row>')
            written += 1
            if written % PROGRESS_EVERY == 0:
                print(f"    Excel: {written:,} registros escritos")
        handle.write(f'</sheetData><autoFilter ref="A1:{last_column}{last_row}"/><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>')
    return written


def styles_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="DD/MM/YYYY"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF17365D"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'''


def table_xml(table_name: str, last_row: int) -> str:
    columns = "".join(f'<tableColumn id="{i}" name={quoteattr(field)}/>' for i, field in enumerate(OUTPUT_FIELDS, 1))
    last_column = excel_column_letter(len(OUTPUT_FIELDS))
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name={quoteattr(table_name)} displayName={quoteattr(table_name)} ref="A1:{last_column}{last_row}" totalsRowShown="0"><autoFilter ref="A1:{last_column}{last_row}"/><tableColumns count="{len(OUTPUT_FIELDS)}">{columns}</tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>'''


def validate_excel_names(sheet_name: str, table_name: str) -> None:
    if not sheet_name or len(sheet_name) > 31 or any(ch in sheet_name for ch in '[]:*?/\\'):
        raise ConsolidationError("'sheet_name' no es válido para Excel: máximo 31 caracteres y sin []:*?/\\.")
    if not re.fullmatch(r"[A-Za-z_\\][A-Za-z0-9_.\\]*", table_name) or re.fullmatch(r"[A-Za-z]{1,3}[1-9][0-9]*", table_name):
        raise ConsolidationError("'table_name' no es un nombre de tabla válido para Excel.")


def write_excel(connection: sqlite3.Connection, output_path: Path, sheet_name: str, table_name: str, unique_count: int) -> int:
    if unique_count + 1 > EXCEL_MAX_ROWS:
        raise ConsolidationError(f"El resultado tiene {unique_count:,} registros y supera el máximo de Excel ({EXCEL_MAX_ROWS - 1:,} filas de datos).")
    validate_excel_names(sheet_name, table_name)
    temporary = temp_output_path(output_path)
    worksheet_temp: Path | None = None
    try:
        if temporary.exists():
            temporary.unlink()
        fd, tmp_name = tempfile.mkstemp(prefix="inventario_sheet_", suffix=".xml", dir=output_path.parent)
        os.close(fd)
        worksheet_temp = Path(tmp_name)
        written = write_worksheet_xml(connection, worksheet_temp, unique_count)
        now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
        parts = {
            "[Content_Types].xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>''',
            "_rels/.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>''',
            "xl/workbook.xml": f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets><sheet name={quoteattr(sheet_name)} sheetId="1" r:id="rId1"/></sheets></workbook>''',
            "xl/_rels/workbook.xml.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>''',
            "xl/worksheets/_rels/sheet1.xml.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>''',
            "xl/styles.xml": styles_xml(),
            "xl/tables/table1.xml": table_xml(table_name, written + 1),
            "docProps/core.xml": f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>consolidar_inventario.py</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified></cp:coreProperties>''',
            "docProps/app.xml": f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Python Standard Library</Application><AppVersion>3.11</AppVersion></Properties>''',
        }
        with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED, compresslevel=6, allowZip64=True) as xlsx:
            for name, content in parts.items():
                xlsx.writestr(name, content)
            xlsx.write(worksheet_temp, "xl/worksheets/sheet1.xml")
        with zipfile.ZipFile(temporary, "r") as check:
            bad = check.testzip()
            if bad:
                raise ConsolidationError(f"El XLSX temporal está dañado: {bad}")
            worksheet_xml = check.read("xl/worksheets/sheet1.xml").decode("utf-8")
            table_definition = check.read("xl/tables/table1.xml").decode("utf-8")
            required_headers = (
                "perimetro_agente",
                "Ejecución_agente",
                "resultado_agente",
                "fecha_ult_ejecucion_agente",
                "partner",
                "APLICACION_BAJA",
                *SLA_FIELDS,
            )
            missing_headers = [header for header in required_headers if header not in worksheet_xml]
            if missing_headers:
                raise ConsolidationError(
                    "Validación interna fallida: faltan cabeceras en el XLSX temporal: "
                    + ", ".join(missing_headers)
                )
            last_column = excel_column_letter(len(OUTPUT_FIELDS))
            if f'ref="A1:{last_column}{written + 1}"' not in table_definition:
                raise ConsolidationError(
                    "Validación interna fallida: el rango de la tabla Excel no coincide con sus columnas."
                )
        os.replace(temporary, output_path)
        return written
    except PermissionError as exc:
        raise ConsolidationError(f"No se puede escribir o reemplazar {output_path}. Cierre Excel si está abierto: {exc}") from exc
    except OSError as exc:
        if getattr(exc, "errno", None) == 28:
            raise ConsolidationError(f"Falta de espacio en disco al escribir el XLSX: {temporary}") from exc
        raise ConsolidationError(f"Error durante la escritura del XLSX {temporary}: {exc}") from exc
    except (zipfile.BadZipFile, RuntimeError) as exc:
        raise ConsolidationError(f"Error durante la generación del XLSX: {exc}") from exc
    finally:
        if worksheet_temp is not None:
            try:
                worksheet_temp.unlink(missing_ok=True)
            except OSError:
                pass

def write_ide00838_diagnostic_logs(
    connection: sqlite3.Connection,
    applications: dict[str, bool],
    ide00838_file: Path,
    output_path: Path,
) -> tuple[Path, Path]:
    """Genera un resumen TXT y un CSV detallado del cruce con IDE00838."""
    logs_dir = output_path.parent / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    summary_path = logs_dir / f"cruce_IDE00838_{timestamp}.log"
    detail_path = logs_dir / f"cruce_IDE00838_detalle_{timestamp}.csv"
    query = 'SELECT "endpoint__product__name", COUNT(*) FROM inventario GROUP BY "endpoint__product__name" ORDER BY "endpoint__product__name"'
    total_distinct = matched = active = inactive = unmatched = empty = 0
    result_si = result_no = 0

    with detail_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle, delimiter=";")
        writer.writerow([
            "endpoint__product__name_inventario", "nombre_normalizado_para_cruce",
            "encontrada_en_IDE00838", "activa_IDE00838", "APLICACION_BAJA",
            "registros_inventario", "motivo",
        ])
        for application_name, records in connection.execute(query):
            total_distinct += 1
            normalized = normalize_application_name(application_name)
            active_value, matched_key, match_method = resolve_ide00838_application(application_name, applications)
            if not normalized:
                found, result = False, "SI"
                reason = "Nombre de aplicación vacío en inventario"
                empty += 1
                unmatched += 1
            elif active_value is not None:
                found = True
                matched += 1
                if active_value:
                    active += 1
                    result = "NO"
                    reason = f"Encontrada por {match_method}: {matched_key}; Activa=S"
                else:
                    inactive += 1
                    result = "SI"
                    reason = f"Encontrada por {match_method}: {matched_key}; Activa distinta de S"
            else:
                found, result = False, "SI"
                unmatched += 1
                reason = "No encontrada por clave única Application - <BP>.<D>.<B>"
            if result == "SI":
                result_si += records
            else:
                result_no += records
            writer.writerow([
                application_name or "", normalized, "SI" if found else "NO",
                "SI" if active_value is True else ("NO" if active_value is False else ""),
                result, records, reason,
            ])

    summary_lines = [
        f"SCRIPT_VERSION={SCRIPT_VERSION}", f"FECHA={datetime.now().isoformat(timespec='seconds')}",
        f"IDE00838_FILE={ide00838_file}", f"OUTPUT_FILE={output_path}",
        "COMPOSICION_IDE00838=clave única: Application - <BP>.<D>.<B>",
        "REGLA=No encontrada: SI | Encontrada activa: NO | Encontrada no activa: SI",
        f"CATALOGO_IDE00838={len(applications)}",
        f"APLICACIONES_DISTINTAS_INVENTARIO={total_distinct}", f"ENCONTRADAS={matched}",
        f"ENCONTRADAS_ACTIVAS={active}", f"ENCONTRADAS_NO_ACTIVAS={inactive}",
        f"NO_ENCONTRADAS={unmatched}", f"NOMBRES_VACIOS={empty}",
        f"REGISTROS_RESULTADO_SI={result_si}", f"REGISTROS_RESULTADO_NO={result_no}",
        f"DETALLE_CSV={detail_path}",
    ]
    summary_text = "\n".join(summary_lines) + "\n"
    summary_path.write_text(summary_text, encoding="utf-8")
    (logs_dir / "cruce_IDE00838_ULTIMO.log").write_text(summary_text, encoding="utf-8")
    shutil.copyfile(detail_path, logs_dir / "cruce_IDE00838_detalle_ULTIMO.csv")
    return summary_path, detail_path


def verify_final_excel(output_path: Path) -> None:
    """Verifica el XLSX definitivo después del reemplazo, no solo el temporal."""
    try:
        with zipfile.ZipFile(output_path, "r") as workbook:
            worksheet_xml = workbook.read("xl/worksheets/sheet1.xml").decode("utf-8")
            table_xml_content = workbook.read("xl/tables/table1.xml").decode("utf-8")
    except (OSError, KeyError, zipfile.BadZipFile) as exc:
        raise ConsolidationError(f"No se puede validar el Excel definitivo {output_path}: {exc}") from exc

    required_headers = (
        "framework",
        "perimetro_agente",
        "Ejecución_agente",
        "resultado_agente",
        "fecha_ult_ejecucion_agente",
        "partner",
        "APLICACION_BAJA",
        *SLA_FIELDS,
    )
    missing = [header for header in required_headers if header not in worksheet_xml]
    if missing:
        raise ConsolidationError(
            "El Excel definitivo no contiene estas cabeceras: " + ", ".join(missing)
        )
    last_column = excel_column_letter(len(OUTPUT_FIELDS))
    if f'ref="A1:{last_column}' not in table_xml_content or f'<tableColumns count="{len(OUTPUT_FIELDS)}">' not in table_xml_content:
        raise ConsolidationError(
            "El Excel definitivo no tiene el rango esperado o el número esperado de columnas."
        )


def remove_database_files(db_path: Path) -> None:
    for suffix in ("", "-wal", "-shm"):
        candidate = Path(str(db_path) + suffix)
        try:
            if candidate.exists():
                candidate.unlink()
        except OSError as exc:
            print(f"AVISO: no se pudo borrar la base temporal {candidate}: {exc}")


def main() -> int:
    started = time.perf_counter()
    connection: sqlite3.Connection | None = None
    config: dict[str, Any] | None = None
    db_path: Path | None = None
    success = False

    try:
        config_path = script_directory() / CONFIG_FILENAME
        partners_config_path = script_directory() / PARTNERS_CONFIG_FILENAME
        config = load_config(config_path)
        partners = load_partners_config(partners_config_path)
        input_folder = Path(os.path.expandvars(os.path.expanduser(config["input_folder"]))).resolve()
        output_path = Path(os.path.expandvars(os.path.expanduser(config["output_file"]))).resolve()

        if not input_folder.is_dir():
            raise ConsolidationError(f"La carpeta de entrada no existe: {input_folder}")
        if output_path.suffix.lower() != ".xlsx":
            raise ConsolidationError("'output_file' debe tener extensión .xlsx.")
        try:
            output_path.parent.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            raise ConsolidationError(f"No se puede crear la carpeta de salida {output_path.parent}: {exc}") from exc

        logs_dir = output_path.parent / "logs"
        logs_dir.mkdir(parents=True, exist_ok=True)
        startup_log_path = logs_dir / "ejecucion_inventario_ULTIMA.log"
        startup_log_path.write_text(
            "\n".join([
                f"INICIO={datetime.now().isoformat(timespec='seconds')}",
                f"SCRIPT_VERSION={SCRIPT_VERSION}",
                f"SCRIPT_FILE={Path(__file__).resolve()}",
                f"WORKING_DIRECTORY={Path.cwd()}",
                f"INPUT_FOLDER={input_folder}",
                f"OUTPUT_FILE={output_path}",
                f"LOGS_FOLDER={logs_dir}",
                "ESTADO=INICIADO",
            ]) + "\n",
            encoding="utf-8",
        )
        logger = logging.getLogger("vulgest.consolidacion")
        logger.setLevel(logging.DEBUG)
        logger.handlers.clear()
        handler = logging.FileHandler(startup_log_path, mode="a", encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s", "%Y-%m-%d %H:%M:%S"))
        logger.addHandler(handler)
        logger.propagate = False
        logger.info("Ejecución iniciada; versión=%s", SCRIPT_VERSION)

        print(f"VERSIÓN DEL SCRIPT: {SCRIPT_VERSION}")
        print(f"LOG DE EJECUCIÓN CREADO: {startup_log_path}")
        print(f"Límite máximo por campo CSV: {CSV_FIELD_SIZE_LIMIT:,} caracteres")
        print("CONFIGURACIÓN UTILIZADA")
        print(f"  Configuración: {config_path}")
        print(f"  Configuración de partners: {partners_config_path}")
        print(f"  Agrupaciones de partners: {len(partners):,}")
        print(f"  Carpeta de entrada: {input_folder}")
        print(f"  Archivo de salida: {output_path}")
        print(f"  Prefijos: {', '.join(config['prefixes'])}")
        print(f"  Hoja: {config['sheet_name']}")
        print(f"  Tabla: {config['table_name']}")
        print(f"  Conservar base temporal: {config['keep_database']}")
        print(f"  Frameworks con agente: {', '.join(config['fw_agents'].keys())}")
        print(f"  Columnas de salida: {len(OUTPUT_FIELDS)}")
        print(f"  Últimas columnas: {', '.join(OUTPUT_FIELDS[-7:])}")

        selected = select_latest_files(input_folder, config["prefixes"])
        agent_file = select_agent_execution_file(input_folder)
        ide00838_file = select_ide00838_file(input_folder)
        logger.info("Fuentes seleccionadas: %s; agente=%s; IDE00838=%s", selected, agent_file, ide00838_file)
        print("\nARCHIVOS SELECCIONADOS")
        for prefix, path in selected:
            modified = datetime.fromtimestamp(path.stat().st_mtime).strftime("%d/%m/%Y %H:%M:%S")
            print(f"  {prefix}: {path} (modificado {modified})")
        agent_modified = datetime.fromtimestamp(agent_file.stat().st_mtime).strftime("%d/%m/%Y %H:%M:%S")
        print(f"  remediacion_agentica: {agent_file} (modificado {agent_modified})")
        ide_modified = datetime.fromtimestamp(ide00838_file.stat().st_mtime).strftime("%d/%m/%Y %H:%M:%S")
        print(f"  IDE00838: {ide00838_file} (modificado {ide_modified})")

        print("\nPROCESAMIENTO DE IDE00838")
        stage_started = time.perf_counter()
        applications, ide_stats = load_ide00838_applications(ide00838_file)
        logger.info("IDE00838 cargado: %s", {"estadisticas": ide_stats, "aplicaciones_unicas": len(applications), "duracion_segundos": round(time.perf_counter() - stage_started, 3)})
        print(
            f"  Finalizado: {ide_stats['read']:,} registros | {len(applications):,} aplicaciones únicas | "
            f"{ide_stats['active']:,} activas | {ide_stats['inactive']:,} no activas"
        )

        estimated_size = estimate_uncompressed_input_size(selected)
        check_disk_space(output_path.parent, estimated_size)

        db_path = output_path.parent / DATABASE_FILENAME
        connection, insert_sql = create_database(db_path)
        create_agent_table(connection)
        logger.info("Base SQLite temporal creada: %s", db_path)

        print("\nPROCESAMIENTO DE EJECUCIONES DE AGENTE")
        stage_started = time.perf_counter()
        agent_stats = process_agent_execution_file(connection, agent_file)
        logger.info("Ejecuciones de agente procesadas: %s; duracion_segundos=%.3f", agent_stats, time.perf_counter() - stage_started)
        print(f"  Finalizado: {agent_stats['read']:,} registros | {agent_stats['completed']:,} completed | {agent_stats['findings']:,} findings extraídos | {agent_stats['json_invalid']:,} JSON inválidos")

        totals = {"read": 0, "valid": 0, "discarded": 0, "changed": 0}
        print("\nPROCESAMIENTO")
        for prefix, path in selected:
            print(f"Procesando [{prefix}]: {path}")
            stage_started = time.perf_counter()
            logger.info("Inicio de inventario prefijo=%s archivo=%s", prefix, path)
            file_stats = process_file(connection, insert_sql, path, config["_fw_agents_normalized"], partners, applications)
            for key in totals:
                totals[key] += file_stats[key]
            print(
                f"  Finalizado: {file_stats['read']:,} leídos | {file_stats['valid']:,} UUID válidos | "
                f"{file_stats['discarded']:,} descartados | {file_stats['changed']:,} inserciones/actualizaciones"
            )
            logger.info("Inventario finalizado prefijo=%s estadisticas=%s duracion_segundos=%.3f", prefix, file_stats, time.perf_counter() - stage_started)

        unique_count = connection.execute("SELECT COUNT(*) FROM inventario").fetchone()[0]
        baja_si = connection.execute("SELECT COUNT(*) FROM inventario WHERE \"APLICACION_BAJA\" = 'SI'").fetchone()[0]
        baja_no = connection.execute("SELECT COUNT(*) FROM inventario WHERE \"APLICACION_BAJA\" = 'NO'").fetchone()[0]
        ide_names = {
            normalize_application_name(row[0])
            for row in connection.execute('SELECT DISTINCT "endpoint__product__name" FROM inventario')
            if normalize_application_name(row[0])
        }
        ide_matched = sum(1 for name in ide_names if resolve_ide00838_application(name, applications)[0] is not None)
        ide_matched_inactive = sum(1 for name in ide_names if resolve_ide00838_application(name, applications)[0] is False)
        ide_unmatched = sum(1 for name in ide_names if resolve_ide00838_application(name, applications)[0] is None)
        print("\nCRUCE CON IDE00838")
        print(f"  Aplicaciones distintas del inventario: {len(ide_names):,}")
        print(f"  Encontradas en IDE00838: {ide_matched:,}")
        print(f"  Encontradas como no activas: {ide_matched_inactive:,}")
        print(f"  No encontradas en IDE00838: {ide_unmatched:,}")
        print(f"  Resultado APLICACION_BAJA: SI={baja_si:,} | NO={baja_no:,}")
        ide_log_path, ide_detail_path = write_ide00838_diagnostic_logs(
            connection, applications, ide00838_file, output_path
        )
        logger.info("Diagnósticos IDE00838 escritos: resumen=%s detalle=%s", ide_log_path, ide_detail_path)
        print(f"  Log resumen del cruce: {ide_log_path}")
        print(f"  Log detallado del cruce: {ide_detail_path}")
        unique_agent_findings, matched_agent_findings = enrich_inventory_with_agent_data(connection)
        unmatched_agent_csv, unmatched_agent_count = write_unmatched_agent_findings_csv(
            connection, output_path
        )
        logger.info("Cruce de agentes finalizado: findings_unicos=%s coincidentes=%s no_coincidentes=%s csv=%s", unique_agent_findings, matched_agent_findings, unmatched_agent_count, unmatched_agent_csv)
        print("\nCRUCE CON EJECUCIONES DE AGENTE")
        print(f"  Findings únicos en ejecuciones completed: {unique_agent_findings:,}")
        print(f"  Findings del inventario con ejecución: {matched_agent_findings:,}")
        print(f"  Findings de agente sin coincidencia: {unmatched_agent_count:,}")
        print(f"  CSV detallado de no coincidentes: {unmatched_agent_csv}")
        print("\nGENERACIÓN DE EXCEL")
        print(f"  UUID únicos finales: {unique_count:,}")
        stage_started = time.perf_counter()
        logger.info("Inicio de generación XLSX: destino=%s filas=%s", output_path, unique_count)
        written = write_excel(
            connection,
            output_path,
            config["sheet_name"],
            config["table_name"],
            unique_count,
        )
        verify_final_excel(output_path)
        logger.info("XLSX generado y validado: filas=%s bytes=%s duracion_segundos=%.3f", written, output_path.stat().st_size, time.perf_counter() - stage_started)
        final_stat = output_path.stat()
        validation_path = output_path.with_name(output_path.stem + "_validacion.txt")
        validation_path.write_text(
            "\n".join([
                f"SCRIPT_VERSION={SCRIPT_VERSION}",
                f"OUTPUT_FILE={output_path}",
                f"OUTPUT_SIZE={final_stat.st_size}",
                f"COLUMN_COUNT={len(OUTPUT_FIELDS)}",
                f"LAST_COLUMN={excel_column_letter(len(OUTPUT_FIELDS))}",
                f"LAST_HEADER={OUTPUT_FIELDS[-1]}",
                f"TABLE_RANGE=A:{excel_column_letter(len(OUTPUT_FIELDS))}",
                "APLICACION_BAJA=OK",
            ]) + "\n",
            encoding="utf-8",
        )
        with startup_log_path.open("a", encoding="utf-8") as startup_log:
            startup_log.write(f"FIN={datetime.now().isoformat(timespec='seconds')}\n")
            startup_log.write("ESTADO=FINALIZADO_CORRECTAMENTE\n")
            startup_log.write(f"LOG_RESUMEN_IDE00838={ide_log_path}\n")
            startup_log.write(f"LOG_DETALLE_IDE00838={ide_detail_path}\n")
            startup_log.write(f"CSV_AGENTES_SIN_INVENTARIO={unmatched_agent_csv}\n")
        success = True
        logger.info("Ejecución completada correctamente; duración_total_segundos=%.3f", time.perf_counter() - started)

        elapsed = time.perf_counter() - started
        print("\nRESUMEN FINAL")
        print(f"  Registros lógicos leídos: {totals['read']:,}")
        print(f"  UUID válidos: {totals['valid']:,}")
        print(f"  Registros descartados: {totals['discarded']:,}")
        print(f"  Inserciones o actualizaciones: {totals['changed']:,}")
        print(f"  UUID únicos finales: {unique_count:,}")
        print(f"  Registros escritos en Excel: {written:,}")
        print("  Columnas de agente generadas: Ejecución_agente, resultado_agente, fecha_ult_ejecucion_agente")
        print("  Columna de partner generada: partner (columna AG)")
        print("  Columna de baja generada: APLICACION_BAJA (columna AH)")
        print(f"  Verificación final: {len(OUTPUT_FIELDS)} columnas; tabla A:{excel_column_letter(len(OUTPUT_FIELDS))}; cabeceras SLA presentes")
        print(f"  Tamaño del fichero final: {final_stat.st_size:,} bytes")
        print(f"  Ruta final: {output_path}")
        print(f"  Informe de validación: {validation_path}")
        print(f"  Log resumen IDE00838: {ide_log_path}")
        print(f"  Log detalle IDE00838: {ide_detail_path}")
        print(f"  Tiempo total: {elapsed:.1f} segundos")
        return 0

    except ConsolidationError as exc:
        if 'logger' in locals():
            logger.exception("Error de consolidación")
        print(f"\nERROR: {exc}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        if 'logger' in locals():
            logger.warning("Ejecución cancelada por el usuario")
        print("\nERROR: proceso cancelado por el usuario.", file=sys.stderr)
        return 130
    except Exception as exc:
        if 'logger' in locals():
            logger.exception("Error inesperado")
        print(f"\nERROR INESPERADO: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 1
    finally:
        if connection is not None:
            try:
                connection.close()
            except sqlite3.Error:
                pass
        if 'logger' in locals():
            logger.info("Limpieza finalizada; éxito=%s", success)
            for handler in logger.handlers[:]:
                handler.close()
                logger.removeHandler(handler)
        if db_path is not None and config is not None and not config.get("keep_database", False):
            remove_database_files(db_path)
        if not success and config is not None:
            try:
                failed_output = Path(os.path.expandvars(os.path.expanduser(config["output_file"]))).resolve()
                temporary = temp_output_path(failed_output)
                if temporary.exists():
                    temporary.unlink()
            except OSError:
                pass


if __name__ == "__main__":
    sys.exit(main())
