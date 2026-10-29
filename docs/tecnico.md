# VulGest — Documento técnico

- Versión documental: 1.0
- Iteración: 001
- Fecha: 2026-10-29
- Código de referencia: `consolidar_inventario.py`, `SCRIPT_VERSION=2026.10.28.24-clave-unica-bp-b-d`.

## 1. Arquitectura y componentes

Aplicación de consola en un único script Python. Usa CSV/ZIP como entradas, SQLite como almacenamiento de consolidación y genera directamente un paquete XLSX mediante XML y ZIP. No necesita paquetes Python de terceros.

| Ruta | Función |
| --- | --- |
| `consolidar_inventario.py` | Selección de fuentes, transformación, deduplicación, enriquecimiento, exportación y validación. |
| `inventario_config.json` | Configuración principal junto al script. |
| `partners_config.json` | Mapeo de proveedores a partners, junto al script. |
| `datos/` | Ubicación de fuentes en la configuración actual; también contiene Excel auxiliares que el script no consume. |
| `Inventario.xlsx` | Salida configurada, preexistente a esta iteración. |
| `Inventario_validacion.txt` | Informe de una ejecución previa; declara una versión anterior del script. |
| `logs/` | Diagnósticos de ejecución y del cruce IDE00838. |
| `AGENTS.md` | Instrucciones persistentes de OpenCode para todo el repositorio. |
| `docs/funcional.md`, `docs/tecnico.md` | Documentos vivos obligatorios por iteración. |

En la inspección inicial también estaban `aplicar_parche_v22.py` y `consolidar_inventario.py.bak_v21`. Ambos dejaron de estar presentes durante esta iteración, sin intervención de OpenCode. No se han restaurado. El parche inspeccionado escribía en el consolidador y en su copia histórica; no era parte del flujo normal.

## 2. Entorno y dependencias

- La sintaxis utilizada requiere Python 3.10 o posterior; no se ha certificado aquí una matriz de versiones.
- El intérprete debe disponer de `sqlite3` y del soporte de compresión requerido por `zipfile.ZIP_DEFLATED`.
- Importaciones del consolidador: `csv`, `io`, `itertools`, `unicodedata`, `json`, `os`, `re`, `shutil`, `sqlite3`, `sys`, `time`, `zipfile`, `contextlib`, `datetime`, `pathlib`, `typing`, `tempfile` y `xml.sax.saxutils`.
- El parche histórico inspeccionado importaba `pathlib` y `re`; ambas pertenecen también a la biblioteca estándar. Ya no está presente en la comprobación final.
- No instalar ni usar paquetes externos para desarrollo, pruebas, validación o ejecución; tampoco paquetes ya instalados. Para pruebas nuevas se puede utilizar `unittest`.
- Las herramientas del sistema no Python, como Git y PowerShell, no quedan prohibidas por esta restricción.

## 3. Configuración y ejecución

`main()` carga los JSON desde el directorio del script, no desde el directorio de trabajo. No hay una interfaz de argumentos implementada para seleccionar otra configuración.

| Clave | Valor actual / contrato |
| --- | --- |
| `input_folder` | `C:\dev\VulGest\datos` |
| `output_file` | `C:\dev\VulGest\Inventario.xlsx`; extensión `.xlsx` obligatoria. |
| `prefixes` | Lista no vacía: `open_allrisks_cr5`, `open_allrisks_cr6_10`, `closed_allrisks`. |
| `sheet_name` | `inventario`; máximo 31 caracteres y sin `[]:*?/\`. |
| `table_name` | `tbl_inventario`; validado mediante expresiones regulares del script. |
| `keep_database` | `false`; exige un booleano JSON. |
| `fw_agents` | Objeto framework/fecha; claves recortadas con `casefold()`, fechas convertidas a representación ISO interna. |

Los valores de ruta expanden variables de entorno y `~`. Si se emplean rutas relativas, `Path.resolve()` las interpreta respecto al directorio de trabajo. El mapeo de partners exige claves y valores no vacíos y rechaza agrupaciones contradictorias tras normalizar las claves.

Ejecución habitual desde PowerShell, una vez revisada la configuración y cerrado el Excel de salida:

```powershell
python .\consolidar_inventario.py
```

Este comando procesa las fuentes reales y reemplaza resultados; no se ha ejecutado en esta iteración documental. No necesita una instalación previa mediante gestores de paquetes.

## 4. Flujo de procesamiento

1. `load_config()` y `load_partners_config()`: lectura UTF-8 con BOM opcional y validaciones.
2. Resolución de rutas y creación de la carpeta de salida y `logs/`; escritura inicial de `ejecucion_inventario_ULTIMA.log`.
3. `select_latest_files()`, `select_agent_execution_file()` y `select_ide00838_file()`: selección de archivos por modificación y nombre.
4. `load_ide00838_applications()`: carga completa del catálogo en memoria, detección de codificación y construcción de un diccionario normalizado de aplicaciones.
5. Estimación del tamaño descomprimido de los inventarios y control de espacio libre: mínimo de `max(200 MiB, 1,5 × tamaño estimado)`.
6. `create_database()` y `create_agent_table()`: base temporal y tablas de trabajo.
7. `process_agent_execution_file()`: extracción de findings de ejecuciones completadas y consolidación por timestamp.
8. `process_file()` y `prepare_row()`: lectura de inventarios en streaming, transformación y UPSERT por UUID.
9. Cálculo de estadísticas y `write_ide00838_diagnostic_logs()`.
10. `enrich_inventory_with_agent_data()`: cruce SQL por `CAST(inventario.id AS TEXT)`.
11. `write_excel()`: XML temporal, empaquetado XLSX, controles y reemplazo de la salida con `os.replace()`.
12. `verify_final_excel()`, informe de validación y marcado de finalización correcta en el log.
13. Cierre de SQLite y, si `keep_database=false`, eliminación de la base y archivos WAL/SHM, tanto en éxito como en error. En fallo se intenta eliminar el XLSX temporal.

## 5. Datos y persistencia temporal

### Columnas de origen y salida

Los inventarios requieren estas 27 columnas, que se exportan en este orden:

```text
title, last_detection, state, found_by, custom_risk, id,
finding_uuid, date, planned_resolution_date, cwe, cve,
last_reviewed, created, modified, cappm, date_deadline,
date_paused, date_closed, date_active, ticketing_tool_key, tags,
endpoint__product__name, endpoint__product__plataforma,
endpoint__product__grupo, endpoint__product__servicio,
endpoint__product__main_provider, data_from
```

Se añaden siete columnas:

| Columna Excel | Campo |
| --- | --- |
| AB | `framework` |
| AC | `perimetro_agente` |
| AD | `Ejecución_agente` |
| AE | `resultado_agente` |
| AF | `fecha_ult_ejecucion_agente` |
| AG | `partner` |
| AH | `APLICACION_BAJA` |

Las reglas de estos campos están descritas en `docs/funcional.md`.

### SQLite

- Archivo: `inventario_consolidacion.sqlite` junto al Excel de salida. Si existe, `create_database()` lo elimina antes de crear la nueva base; no se reutiliza como histórico.
- PRAGMA: `journal_mode=WAL`, `synchronous=NORMAL`.
- Tabla `inventario`: 34 campos de salida más `_sort_key TEXT NOT NULL`; `finding_uuid TEXT PRIMARY KEY` y `custom_risk REAL`; los demás campos no tienen tipo declarado.
- UPSERT: actualiza todos los campos salvo el UUID únicamente si `excluded._sort_key > inventario._sort_key`.
- Tabla `agent_executions`: `finding_id TEXT PRIMARY KEY`, `resolution TEXT`, `execution_timestamp TEXT NOT NULL`; actualiza solo con timestamp estrictamente posterior.
- Las fechas válidas se almacenan con `isoformat(timespec="microseconds")`; las comparaciones se realizan sobre estas cadenas normalizadas.
- Lotes de 5.000 registros con commits; progreso cada 25.000 registros. Las estadísticas de cambios incluyen tanto inserciones como actualizaciones, no equivalen a UUID únicos.

### XLSX

- Paquete ZIP con partes Open XML de libro, hoja, estilos, tabla, relaciones y propiedades; sin dependencia de una biblioteca externa de Excel.
- Se escribe la hoja mediante cursor SQLite `ORDER BY rowid` en un XML temporal junto a la salida.
- Se utiliza texto `inlineStr`, escape XML y limpieza de caracteres de control; las fechas se convierten a serie Excel tomando como base 1899-12-30.
- El rango de tabla y de hoja llega a AH; algunas definiciones incluyen literalmente `34` y `AH`. Un cambio de columnas exige revisar estos puntos y sus validaciones.
- El XLSX temporal se denomina `<nombre>.tmp.xlsx`; se verifica la integridad ZIP, las cabeceras derivadas exigidas y el rango de tabla antes de reemplazar el destino.
- La validación final comprueba cabeceras y definición de tabla. No es una certificación completa del esquema Open XML ni una apertura real en Excel.

## 6. Diagnósticos, errores y limitaciones

- Salida normal: código `0`; error controlado o inesperado: `1`; interrupción de teclado: `130`.
- Los errores se presentan en consola; el log de inicio solo se marca como finalizado correctamente tras el éxito. No hay escritura explícita de un estado final de fallo.
- El resumen IDE00838 y su detalle incluyen nombres de aplicaciones: tratar los logs como datos del proyecto, no como documentación pública.
- El informe de validación registra versión, ruta, tamaño, 34 columnas, AH y cabecera final. No contiene una auditoría completa de los valores de todas las celdas.
- La lectura de inventarios y ejecuciones es en streaming, pero IDE00838 se carga completo y ciertas validaciones leen el XML de la hoja completo en memoria.
- La estimación de disco no incluye todos los auxiliares, logs y necesidades reales de SQLite; es una comprobación preventiva, no una garantía de capacidad.
- Se espera que `details` sea un objeto JSON; un JSON válido de otro tipo puede provocar un error inesperado al acceder a `.get()`. No se ha modificado este comportamiento.
- El reemplazo del Excel no hace transaccional todo el proceso: una validación o escritura posterior puede fallar después de reemplazarlo.
- No ejecutar consolidaciones simultáneas sobre la misma salida: los nombres de la base y del XLSX temporal son compartidos.
- El parche v22 inspeccionado era anterior al código actual y escribía archivos al ejecutarse/importarse; no se ha ejecutado.
- `.gitignore` excluye `*.csv` y `*.log`, pero no todos los datos, ZIP, Excel o informes TXT. Revisar el contenido antes de publicar o versionar archivos. No se ha cambiado esta política.

## 7. Política de OpenCode y mantenimiento documental

`AGENTS.md` fija la restricción a biblioteca estándar y obliga a crear o actualizar ambos documentos por cada tarea entregada. Las entradas de historial comparten identificador de iteración y fecha; cuando no existe impacto se debe indicar explícitamente en el ámbito correspondiente.

Estas son instrucciones persistentes de comportamiento del agente, no un bloqueo del sistema operativo ni un mecanismo que impida a otros procesos instalar paquetes. Esta iteración no añade plugins, dependencias ni configuración global de OpenCode.

## 8. Verificación de esta iteración

- Revisión estática del consolidador completo, del parche histórico, de los JSON de configuración y del informe de validación existente.
- Comprobación estática satisfactoria de sintaxis e importaciones del consolidador mediante `ast` y `compile`, usando Python 3.11.6 aislado, sin paquetes de sitio y sin ejecutar el script; comprobación satisfactoria de validez de los dos JSON de configuración.
- Un primer intento de comprobación se interrumpió al no encontrar el parche histórico, que había desaparecido tras la revisión inicial. La comprobación se repitió sobre el consolidador presente, sin restaurar archivos ajenos.
- Comprobación satisfactoria de existencia, estructura e identificador de iteración de ambos documentos, y de las instrucciones persistentes.
- No se han ejecutado pruebas funcionales, el parche ni una consolidación completa. No se ha abierto o regenerado `Inventario.xlsx`, ni modificado los datos o los scripts.
- La evidencia preexistente `Inventario_validacion.txt` declara `2026.10.28.21-ide00838-aliases`; no se atribuyen sus resultados al código actual.

## 9. Historial

| Iteración | Fecha | Versión | Cambios e impacto |
| --- | --- | --- | --- |
| 001 | 2026-10-29 | 1.0 | Primera versión técnica y creación de `AGENTS.md` con restricción a biblioteca estándar y mantenimiento de dos documentos por iteración. Comprobaciones estáticas y documentales, sin cambios en scripts, configuración de negocio ni resultados. |
