"""Pruebas del cruce IDE00838, sin escribir datos ni resultados en disco."""

import contextlib
import csv
import importlib.util
import io
from pathlib import Path
import sqlite3
import unittest
from unittest.mock import mock_open, patch
import zipfile


SCRIPT_PATH = Path(__file__).resolve().parents[1] / "consolidar_inventario.py"
SPEC = importlib.util.spec_from_file_location("consolidar_inventario", SCRIPT_PATH)
inventory = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(inventory)


class Ide00838Tests(unittest.TestCase):
    key = "Application - CLOUD.SERVICIO.001.MODULO"
    old_key = "Application - CLOUD.MODULO.SERVICIO.001"

    def record(self, active="S", platform="CLOUD"):
        row = [""] * 68
        row[1], row[3], row[10], row[67] = "MODULO", "SERVICIO.001", active, platform
        return row

    def csv_bytes(self, rows):
        stream = io.StringIO(newline="")
        csv.writer(stream, delimiter=";").writerows(rows)
        return stream.getvalue().encode("utf-8-sig")

    def load(self, rows):
        with patch.object(Path, "read_bytes", return_value=self.csv_bytes(rows)):
            with contextlib.redirect_stdout(io.StringIO()) as console:
                applications, stats = inventory.load_ide00838_applications(Path("IDE00838.csv"))
        return applications, stats, console.getvalue()

    def test_without_header_uses_bp_d_b(self):
        applications, stats, console = self.load([self.record()])
        self.assertEqual(applications, {inventory.normalize_application_name(self.key): True})
        self.assertEqual(stats["valid_names"], 1)
        self.assertIn("Application - <BP>.<D>.<B>", console)
        self.assertIn("Activa=columna 11", console)

    def test_column_11_controls_activity_not_column_10(self):
        for column_10, column_11, expected in (("N", "S", True), ("S", "N", False)):
            with self.subTest(column_10=column_10, column_11=column_11):
                row = self.record(column_11)
                row[9] = column_10
                applications, _, _ = self.load([row])
                self.assertIs(inventory.resolve_ide00838_application(self.key, applications)[0], expected)

    def test_old_order_is_not_an_alias(self):
        applications, _, _ = self.load([self.record()])
        self.assertIsNone(inventory.resolve_ide00838_application(self.old_key, applications)[0])

    def test_header_uses_platform_type_name(self):
        applications, _, _ = self.load([
            ["Activa", "Nombre", "Plataforma", "Tipo"],
            ["S", "MODULO", "CLOUD", "SERVICIO.001"],
        ])
        self.assertIs(inventory.resolve_ide00838_application(self.key, applications)[0], True)

    def test_header_without_platform_falls_back_to_bp(self):
        header = [""] * 68
        header[:3] = ["Nombre", "Tipo", "Activa"]
        row = [""] * 68
        row[:3] = ["MODULO", "SERVICIO.001", "S"]
        row[67] = "CLOUD"
        applications, _, _ = self.load([header, row])
        self.assertIs(inventory.resolve_ide00838_application(self.key, applications)[0], True)

    def test_normalization_and_match_description(self):
        applications, _, _ = self.load([self.record()])
        active, matched_key, method = inventory.resolve_ide00838_application(
            "\ufeffAPPLICATION\u00a0-  cloud.servicio.001.modulo", applications
        )
        self.assertIs(active, True)
        self.assertEqual(matched_key, inventory.normalize_application_name(self.key))
        self.assertEqual(method, "clave única BP.D.B")

    def test_duplicate_active_wins_in_either_order(self):
        for states in (("N", "S"), ("S", "N")):
            with self.subTest(states=states):
                applications, stats, _ = self.load([self.record(state) for state in states])
                self.assertIs(inventory.resolve_ide00838_application(self.key, applications)[0], True)
                self.assertEqual(stats["duplicates"], 1)

    def test_only_s_means_active(self):
        for value in ("S", "s", " S ", "SI", "SÍ", "Y", "YES", "TRUE", "1", "N", "", "UNKNOWN"):
            with self.subTest(value=value):
                applications, _, _ = self.load([self.record(value)])
                expected = value.strip().upper() == "S"
                self.assertIs(inventory.resolve_ide00838_application(self.key, applications)[0], expected)

    def test_header_activity_also_requires_s(self):
        for value in ("S", "N", "SI", "TRUE", "1", ""):
            with self.subTest(value=value):
                applications, _, _ = self.load([
                    ["Activa", "Nombre", "Plataforma", "Tipo"],
                    [value, "MODULO", "CLOUD", "SERVICIO.001"],
                ])
                self.assertIs(inventory.resolve_ide00838_application(self.key, applications)[0], value == "S")

    def test_short_rows_and_missing_components_are_skipped(self):
        rows = [["short"]]
        for index in (1, 3, 67):
            row = self.record()
            row[index] = ""
            rows.append(row)
        applications, stats, _ = self.load(rows)
        self.assertEqual(applications, {})
        self.assertEqual(stats["read"], 4)

    def test_zip_uses_bp_d_b(self):
        archive_bytes = io.BytesIO()
        with zipfile.ZipFile(archive_bytes, "w") as archive:
            archive.writestr("catalog.csv", self.csv_bytes([self.record()]))
        archive_bytes.seek(0)
        with zipfile.ZipFile(archive_bytes, "r") as archive:
            with patch.object(inventory.zipfile, "ZipFile", return_value=archive):
                with contextlib.redirect_stdout(io.StringIO()):
                    applications, _ = inventory.load_ide00838_applications(Path("IDE00838.zip"))
        self.assertIs(inventory.resolve_ide00838_application(self.key, applications)[0], True)

    def test_baja_rule_is_unchanged(self):
        for state, expected in ((True, "NO"), (False, "SI"), (None, "SI")):
            with self.subTest(state=state):
                applications = {} if state is None else {inventory.normalize_application_name(self.key): state}
                row = {
                    "finding_uuid": "12345678-1234-1234-1234-123456789abc",
                    "endpoint__product__name": self.key,
                }
                values, _ = inventory.prepare_row(row, {}, {}, applications)
                self.assertEqual(values[-1], expected)

    def test_loaded_column_11_drives_baja(self):
        for value, name, expected in (
            ("S", self.key, "NO"), ("N", self.key, "SI"),
            ("SI", self.key, "SI"), ("", self.key, "SI"),
            ("S", self.old_key, "SI"), ("S", "", "SI"),
        ):
            with self.subTest(value=value, name=name):
                source = self.record(value)
                source[9] = "N" if value == "S" else "S"
                applications, _, _ = self.load([source])
                row = {
                    "finding_uuid": "12345678-1234-1234-1234-123456789abc",
                    "endpoint__product__name": name,
                }
                values, _ = inventory.prepare_row(row, {}, {}, applications)
                self.assertEqual(values[-1], expected)

    def test_diagnostics_describe_new_composition(self):
        applications, _, _ = self.load([self.record()])
        with sqlite3.connect(":memory:") as connection:
            connection.execute('CREATE TABLE inventario ("endpoint__product__name" TEXT)')
            connection.executemany("INSERT INTO inventario VALUES (?)", [(self.key,), (self.old_key,)])
            output = mock_open()
            with patch.object(Path, "mkdir"), patch.object(Path, "open", output):
                with patch.object(Path, "write_text") as write_text:
                    with patch.object(inventory.shutil, "copyfile"):
                        inventory.write_ide00838_diagnostic_logs(
                            connection, applications, Path("IDE00838.csv"), Path("Inventario.xlsx")
                        )
            summary = write_text.call_args_list[0].args[0]
            self.assertIn("COMPOSICION_IDE00838=clave única: Application - <BP>.<D>.<B>", summary)
            self.assertIn("ENCONTRADAS=1\n", summary)
            self.assertIn("NO_ENCONTRADAS=1\n", summary)
            detail = "".join(call.args[0] for call in output().write.call_args_list)
            self.assertIn("clave única BP.D.B", detail)
            self.assertIn("No encontrada por clave única Application - <BP>.<D>.<B>", detail)


if __name__ == "__main__":
    unittest.main()
