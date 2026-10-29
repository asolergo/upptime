# VulGest — Documento técnico

- Versión documental: 1.2
- Iteración: 003
- Fecha: 2026-10-29
- Código de referencia: `consolidar_inventario.py`, `SCRIPT_VERSION=2026.10.29.26-activa-columna-11-s`.

## 1. Arquitectura y componentes

Aplicación de consola en un único script Python. Usa CSV/ZIP como entradas, SQLite como almacenamiento de consolidación y genera directamente un paquete XLSX mediante XML y ZIP. No necesita paquetes Python de terceros.

| Ruta | Función |
| --- | --- |
| `consolidar_inventario.py` | Selección de fuentes, transformación, deduplicación, enriquecimiento, exportación y validación. |
| `tests/test_ide00838.py` | Pruebas con `unittest` del cruce y de sus diagnósticos, con CSV/ZIP y SQLite en memoria y escrituras simuladas. |
| `inventario_config.json` | Configuración principal junto al script. |
| `partners_config.json` | Mapeo de proveedores a partners, junto al script. |
| `datos/` | Ubicación de fuentes en la configuración actual; también contiene Excel auxiliares que el script no consume. |
| `Inventario.xlsx` | Salida configurada, preexistente a esta iteración. |
| `Inventario_validacion.txt` | Informe de una ejecución previa; declara una versión anterior del script. |
| `logs/` | Diagnósticos de ejecución y del cruce IDE00838. |
| `AGENTS.md` | Instrucciones persistentes de OpenCode para todo el repositorio. |
| `docs/funcional.md`, `docs/tecnico.md` | Documentos vivos obligatorios por iteración. |

En la inspección inicial de la iteración 001 también estaban `aplicar_parche_v22.py` y `consolidar_inventario.py.bak_v21`. Ambos dejaron de estar presentes durante aquella iteración, sin intervención de OpenCode. No se han restaurado. El parche inspeccionado escribía en el consolidador y en su copia histórica; no era parte del flujo normal.

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

Este comando procesa las fuentes reales y reemplaza resultados; OpenCode no lo ha ejecutado en las iteraciones 001, 002 y 003. Existen resultados de ejecuciones ajenas a estas comprobaciones. No necesita una instalación previa mediante gestores de paquetes.

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

### Cruce IDE00838

- `load_ide00838_applications()` compone `Application - <BP>.<D>.<B>` y normaliza la clave completa. Los índices sin cabecera son BP=67, D=3, B=1 y K=10 (actividad, columna 11), contando desde cero. J=9 ya no se utiliza para determinar la actividad.
- Con cabecera se conservan los índices obtenidos de `Plataforma`, `Tipo`, `Nombre` y `Activa`; la composición pasa a `Plataforma.Tipo.Nombre`. Se mantiene el recurso a BP=67 cuando falta `Plataforma`.
- `resolve_ide00838_application()` no admite la composición anterior como alias y devuelve el método `clave única BP.D.B` para las coincidencias.
- La consola y los logs nuevos de `write_ide00838_diagnostic_logs()` identifican la composición BP.D.B; los logs preexistentes se conservan sin cambios.
- La actividad se determina con `row[active_index].strip().upper() == "S"`, tanto con cabecera como sin ella. Se eliminó la aceptación de otros valores afirmativos. La consola informa la posición utilizada: 11 en el formato sin cabecera, o la posición detectada de `Activa` con cabecera reconocida.
- No cambian la normalización, las reglas de duplicados ni `prepare_row()` para `APLICACION_BAJA`. El usuario confirmó K como indicador de actividad en el formato sin cabecera; las denominaciones oficiales de B/D siguen pendientes de confirmación externa.

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

### Evidencia histórica: iteración 001

- Revisión estática del consolidador completo, del parche histórico, de los JSON de configuración y del informe de validación existente.
- Comprobación estática satisfactoria de sintaxis e importaciones del consolidador mediante `ast` y `compile`, usando Python 3.11.6 aislado, sin paquetes de sitio y sin ejecutar el script; comprobación satisfactoria de validez de los dos JSON de configuración.
- Un primer intento de comprobación se interrumpió al no encontrar el parche histórico, que había desaparecido tras la revisión inicial. La comprobación se repitió sobre el consolidador presente, sin restaurar archivos ajenos.
- Comprobación satisfactoria de existencia, estructura e identificador de iteración de ambos documentos, y de las instrucciones persistentes.
- No se han ejecutado pruebas funcionales, el parche ni una consolidación completa. No se ha abierto o regenerado `Inventario.xlsx`, ni modificado los datos o los scripts.
- La evidencia preexistente `Inventario_validacion.txt` declara `2026.10.28.21-ide00838-aliases`; no se atribuyen sus resultados al código actual.

### Iteración 002: composición BP.D.B

- Se ejecutó satisfactoriamente `python -I -S -B -m unittest discover -s tests -v`: 11 pruebas, solo biblioteca estándar, sin paquetes de sitio ni escritura de bytecode.
- Cobertura: composición sin cabecera, cabeceras reordenadas, ausencia de Plataforma con alternativa BP, rechazo del orden anterior, normalización, descripción de coincidencia, duplicados en ambos órdenes, estados de actividad, filas incompletas, ZIP, APLICACION_BAJA y logs de resumen/detalle.
- Las pruebas no invocan `main()`: usan CSV/ZIP y SQLite en memoria, lecturas simuladas para los CSV sintéticos y escrituras de diagnósticos simuladas.
- Se ejecutó un contraste adicional de solo lectura del catálogo real con los nombres del último detalle histórico. Confirma coincidencias con la composición implementada; no es una consolidación ni una validación del Excel o de los inventarios actuales. No se incorporan datos de inventario a esta documentación.
- Se revisaron ambos documentos y su historial común. No se ejecutaron parches históricos ni se regeneraron o sobrescribieron Excel, datos, logs o informes existentes.

### Iteración 003: actividad en columna 11 y solo S

- Se ejecutó satisfactoriamente `python -I -S -B -m unittest discover -s tests -v`: 14 pruebas con biblioteca estándar, sin paquetes de sitio ni escritura de bytecode.
- Se actualizaron los datos sintéticos para situar actividad en K=10 y se añadieron pruebas con valores opuestos en J/K, valores no S con cabecera y la cadena completa carga del catálogo → `prepare_row()` → `APLICACION_BAJA`. Se mantienen las pruebas de BP.D.B, ZIP, duplicados y diagnósticos.
- El contraste adicional de solo lectura usa `load_ide00838_applications()` y `resolve_ide00838_application()` corregidos sobre el catálogo actual y el detalle histórico indicado por el usuario. No invoca `main()`, no genera salidas y no certifica una nueva consolidación.
- El informe preexistente inspeccionado declara `2026.10.29.25-clave-unica-bp-d-b`; las referencias anteriores a `2026.10.28.21-ide00838-aliases` son evidencia histórica de la iteración 001, no el estado actual del informe.
- Se preservan los cambios locales preexistentes del usuario en Excel e informe. Las pruebas usan escrituras simuladas y SQLite en memoria; no se regeneran Excel, datos, logs ni informes.
- Comprobaciones satisfactorias de sintaxis, existencia e historial común de los dos documentos y `git diff --check`. Los SHA-256 anteriores y posteriores coinciden para Excel, informe, catálogo y los dos archivos de diagnóstico indicados por el usuario.

## 9. Historial

| Iteración | Fecha | Versión | Cambios e impacto |
| --- | --- | --- | --- |
| 001 | 2026-10-29 | 1.0 | Primera versión técnica y creación de `AGENTS.md` con restricción a biblioteca estándar y mantenimiento de dos documentos por iteración. Comprobaciones estáticas y documentales, sin cambios en scripts, configuración de negocio ni resultados. |
| 002 | 2026-10-29 | 1.1 | Incorpora el análisis previo, actualiza composición, versión y diagnósticos a BP.D.B y añade 11 pruebas con unittest, todas satisfactorias. Contraste adicional de solo lectura. Sin nuevas dependencias, cambios de configuración ni regeneración de resultados. |
| 003 | 2026-10-29 | 1.2 | Actualiza versión, índice de actividad a K=10 y comparación exclusiva con S. Amplía la suite a 14 pruebas satisfactorias y realiza contraste de solo lectura. Mantiene BP.D.B, formatos y dependencias; sin regenerar ni sobrescribir resultados existentes. |
