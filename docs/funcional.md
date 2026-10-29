# VulGest — Documento funcional

- Versión documental: 1.6
- Iteración: 007
- Fecha: 2026-10-29
- Código de referencia: `consolidar_inventario.py`, versión `2026.10.29.29-logging-detallado`.
- Base de esta versión: registro detallado de etapas, estadísticas, duración y errores; no se ha ejecutado la consolidación completa.

## 1. Objetivo y alcance

Consolidar hallazgos de vulnerabilidades procedentes de CSV o ZIP en un único Excel, eliminando duplicados por UUID y enriqueciendo los registros con información de frameworks, perímetro de agente, ejecuciones de remediación, partners y estado de las aplicaciones.

La ejecución actual no utiliza los Excel auxiliares presentes en `datos/`. Durante la revisión inicial se observó un script de parche histórico, ajeno al flujo normal, que ya no estaba presente al realizar las comprobaciones finales.

## 2. Entradas y selección

| Entrada | Comportamiento |
| --- | --- |
| `inventario_config.json` | Define carpeta de entrada, Excel de salida, prefijos, hoja, tabla, conservación de la base temporal y fechas de incorporación de agentes por framework. |
| `partners_config.json` | Define el objeto `mapping` de proveedores a partners agrupados. Es obligatorio y no puede estar vacío. |
| Inventarios CSV/ZIP | Se selecciona un archivo por cada prefijo configurado. Actualmente: `open_allrisks_cr5`, `open_allrisks_cr6_10` y `closed_allrisks`. |
| `remediacion_agentica*.csv` o `.zip` | Fuente obligatoria de ejecuciones de remediación. |
| `IDE00838*.csv` o `.zip` | Catálogo obligatorio de aplicaciones y estado activo. |

- La búsqueda se limita a los archivos directamente contenidos en la carpeta de entrada; no es recursiva.
- Los prefijos y extensiones se comparan sin distinguir mayúsculas y minúsculas.
- Se elige el archivo de modificación más reciente, no la fecha escrita en su nombre. En empate se usa el nombre normalizado como segundo criterio y se toma el máximo.
- En un ZIP se utiliza el CSV de mayor tamaño descomprimido. Para inventarios y ejecuciones, el empate se resuelve por nombre; en IDE00838, por el primer miembro encontrado entre los de tamaño máximo.
- Los inventarios usan `;` y las ejecuciones de agente usan `,`, ambos con cabecera y codificación UTF-8 con BOM opcional. IDE00838 usa `;` y admite UTF-8, CP1252 o Latin-1.

## 3. Reglas de consolidación

### RF-01. Validación y normalización del hallazgo

- Cada inventario debe contener las 27 columnas de origen enumeradas en el documento técnico; se permiten columnas adicionales, que no se exportan.
- Un registro solo se admite si `finding_uuid` tiene formato hexadecimal `8-4-4-4-12`. Se eliminan espacios exteriores y se convierte a minúsculas. No se validan la versión ni la variante del UUID.
- Los UUID ausentes o mal formados se descartan y contabilizan.
- Se eliminan espacios exteriores de los valores de origen.
- `custom_risk` se convierte en número, admitiendo coma o punto decimal y separadores de millares combinados. Un valor no convertible queda vacío.
- `id` se convierte a entero cuando es posible; de lo contrario se conserva como texto.
- Se reconocen fechas ISO y formatos día/mes/año, con horas opcionales. Las fechas con zona horaria se convierten a UTC; las que no tienen zona se conservan sin asignarles una.
- Las fechas no interpretables se conservan como texto. Para comparar `data_from`, una fecha no interpretable recibe una clave vacía.

### RF-02. Duplicados

- La clave única es `finding_uuid`, no `id`.
- Para un UUID repetido se conserva el registro con `data_from` más reciente.
- Solo una fecha estrictamente mayor sustituye al registro previo. En empate se conserva el primero procesado.
- Una fecha válida prevalece sobre una clave de fecha vacía; si ambas son vacías, se conserva el primer registro.
- Los archivos se procesan en el orden de los prefijos configurados y sus registros en el orden de lectura.
- El resultado se exporta por orden interno de inserción en SQLite, no ordenado por riesgo o fecha.

### RF-03. Framework y perímetro de agente

`framework` se obtiene del segmento entre `Application - ` y el primer punto de `endpoint__product__name`, usando el patrón del código, sensible a mayúsculas. Si no hay coincidencia, queda vacío.

Si el framework está configurado en `fw_agents` (comparación sin distinguir mayúsculas), `perimetro_agente` vale `SI` cuando `date_active` es válida y se cumple alguna condición respecto a la fecha de corte:

1. `date_active` es igual o posterior al corte.
2. `date_active` es anterior y `date_closed` es válida y estrictamente posterior al corte.
3. `date_active` es anterior y `date_closed` está vacía.

En los demás casos del framework configurado vale `NO`. Si el framework no está configurado, queda vacío. Un cierre exactamente en la fecha de corte no cumple la segunda condición; un cierre no vacío pero inválido tampoco.

Fechas de corte de la configuración inspeccionada:

| Framework | Fecha |
| --- | --- |
| `absis3-srv` | 14/09/2026 |
| `absis3-ca` | 26/10/2026 |
| `cloud` | 01/10/2026 |
| `batch` | 01/10/2026 |

### RF-04. Ejecuciones de remediación

- Solo se consideran filas con `status=completed`, sin distinguir mayúsculas y tras eliminar espacios exteriores.
- Se requiere un `timestamp` válido y un `details` no vacío con JSON válido que contenga una lista `findings`.
- De cada elemento de tipo objeto se extraen `finding_id` no vacío y `resolution`.
- Para cada `finding_id` se conserva la ejecución con timestamp más reciente; en empate se mantiene la primera.
- Se cruza `finding_id` con el `id` del inventario convertido a texto, no con `finding_uuid`.
- Si existe coincidencia: `Ejecución_agente=SI`, `resultado_agente=resolution` y `fecha_ult_ejecucion_agente=timestamp`.
- Sin coincidencia: `Ejecución_agente=NO` y los otros dos campos vacíos.
- Este cruce no depende de que `perimetro_agente` sea `SI`.

#### Extracción puntual de ejecuciones sin inventario (iteración 004)

- Se entrega `ejecuciones_agente_sin_inventario_20261029.csv`, con una fila por `finding_id` de ejecuciones admitidas que no aparece en el campo `id` del Excel existente. No se compara con `finding_uuid` ni se reconstruye el inventario desde sus CSV.
- Se conserva la última ejecución admisible por finding, manteniendo la primera en caso de empate, igual que el cruce habitual. El detalle incluye agente, repositorio, acción, resolución, fecha, contexto del finding y referencia al registro de origen.
- El contador `ejecuciones_completed_del_finding` indica cuántos registros de ejecución admisibles contienen ese finding; varias apariciones en un mismo registro cuentan una sola vez.
- Se verifican 3.765 findings únicos, 3.706 coincidencias únicas y 59 sin coincidencia. En este Excel también son 3.706 filas coincidentes; no debe suponerse en general que filas del inventario y findings únicos son la misma magnitud.
- Es una entrega puntual, no una salida automática del consolidador. Sin cambios funcionales en las reglas de consolidación o enriquecimiento existentes; únicamente se añade el detalle solicitado, sin modificar inventario ni fuentes.

#### Generación automática de detalle (iteración 005)

- En cada consolidación, durante el cruce de ejecuciones de agente contra el inventario SQLite ya consolidado, el script genera un CSV con los findings únicos sin coincidencia por `id`.
- Se escriben un detalle fechado y una copia `ULTIMO` en `logs/`. La salida incluye datos de la ejecución y el objeto JSON completo del finding; se conserva el registro más reciente por `finding_id`, con la misma regla de desempate que el cruce.
- El cambio previo fue una extracción puntual del Excel existente. Desde esta iteración la generación queda integrada en el flujo, pero no se ha ejecutado una consolidación real para producir los nuevos logs ni modificar el Excel.

### RF-05. Partner

Se busca `endpoint__product__main_provider` en el mapeo de partners, sin distinguir mayúsculas y eliminando espacios exteriores. Si no existe agrupación, se conserva el proveedor original. Si el proveedor está vacío, `partner` queda vacío.

### RF-06. Aplicación de baja

En IDE00838 sin cabecera se forma exclusivamente la clave `Application - <BP>.<D>.<B>`, usando los valores de esas posiciones, en ese orden. Sustituye a la composición anterior `BP.B.D`, que no se conserva como alias.

- Con cabecera reconocida se mantienen las asociaciones existentes: `Nombre` corresponde al componente B y `Tipo` al componente D. La clave es `Application - <Plataforma>.<Tipo>.<Nombre>`; si falta `Plataforma`, se usa la posición BP.
- Sin cabecera reconocida se usan B (2), D (4), K (11, activa) y BP (68). La columna J (10) no interviene en la actividad. Las filas cortas o sin alguno de los tres componentes de la clave se omiten.
- El usuario ha confirmado que el indicador de actividad está en la columna 11 del formato sin cabecera. Las denominaciones oficiales de B y D siguen pendientes de confirmación externa; no se alteran sus posiciones ni el orden BP.D.B.
- La normalización aplica Unicode NFKC, elimina BOM, normaliza espacios y no distingue mayúsculas. No se buscan sufijos ni alias.
- Se considera activa únicamente si el valor, recortado y en mayúsculas, es `S`. `N`, vacío y cualquier otro valor se interpretan como inactiva; `SI`, `SÍ`, `Y`, `YES`, `TRUE` y `1` ya no se aceptan como equivalentes de `S`. Esta interpretación también se aplica al campo `Activa` con cabecera reconocida.
- Si hay varias filas para la misma clave, basta una activa para considerar activa la aplicación.

| Resultado del cruce | `APLICACION_BAJA` |
| --- | --- |
| Encontrada y activa | `NO` |
| Encontrada e inactiva | `SI` |
| No encontrada o nombre vacío | `SI` |

`SI` es el resultado de esta regla y no demuestra por sí solo una baja administrativa: también incluye aplicaciones no encontradas.

## 4. Salidas y controles

- `Inventario.xlsx`, según la configuración actual: hoja `inventario`, tabla `tbl_inventario`, 34 columnas de A a AH, cabecera inmovilizada y filtros.
- Las fechas reconocidas son fechas Excel con presentación `DD/MM/YYYY`; el valor conserva la hora cuando existe. El riesgo es numérico con formato General. `found_by`, `cwe` y `finding_uuid` se fuerzan a texto.
- Se admiten como máximo 1.048.575 filas de datos, además de la cabecera.
- Se comprueba espacio libre antes de crear la base temporal y se genera primero un Excel temporal, que se comprueba antes de reemplazar el definitivo.
- Se comprueban también las cabeceras derivadas y el alcance de la tabla del Excel definitivo.
- Se generan un informe `<nombre_salida>_validacion.txt`, un log detallado de ejecución y logs de resumen y detalle de los cruces IDE00838 y de ejecuciones de agente en `logs/`, con copias `ULTIMO`/`ULTIMA` según el archivo. El log de ejecución incorpora eventos fechados para selección y procesamiento de fuentes, cruces, exportación, validación, duración y errores con traceback; no registra filas completas del inventario.
- La consola informa registros leídos, descartados, inserciones/actualizaciones, coincidencias y duración.
- El cruce de agente genera también el CSV de findings sin inventario, con delimitador `;` y UTF-8 con BOM para facilitar su apertura en Excel. Su formato se describe en `docs/tecnico.md`.
- El cruce de agente no coincidentes usa un índice temporal SQLite sobre el `id` del inventario convertido a texto; no cambia los resultados ni las reglas funcionales.
- Si falta una fuente obligatoria, la configuración es inválida o no se puede escribir la salida, la ejecución informa un error. No se garantiza que todo fallo deje intactas todas las salidas: algunos controles son posteriores al reemplazo y los logs se escriben durante el proceso.

## 5. Reglas de trabajo de OpenCode

- Solo se permite la biblioteca estándar de Python; está prohibido instalar o utilizar paquetes externos, incluso preinstalados.
- Cada iteración de OpenCode debe entregar actualizados este documento y `docs/tecnico.md`, con historial común. No significa generar documentos por cada fila procesada o ejecución del consolidador.
- La primera iteración establece la línea base documental; las posteriores deben reflejar cambios o declarar expresamente la ausencia de impacto.
- Las instrucciones persistentes y la lista de comprobación están en `AGENTS.md`.

## 6. Estado de validación

La iteración 005 añade una prueba sintética del flujo de ejecución: verifica la exclusión de IDs presentes, selección de la ejecución más reciente, campos de detalle y CSV `ULTIMO`. La iteración 006 añade comprobación del índice empleado por esa exportación. La iteración 004 generó puntualmente 59 filas contra el Excel existente, pero esa evidencia no certifica la generación automática sobre una consolidación completa. Se ejecutan las pruebas automatizadas descritas en el documento técnico; no se ejecuta `main()` ni se regenera el Excel.

No se ejecuta `main()`: la salida integrada y su escritura en `logs/` quedan verificadas con datos sintéticos, no con una consolidación real. El CSV puntual creado en la iteración 004 permanece como artefacto de esa iteración y no es producido por las pruebas actuales. No se modifica el Excel existente ni las fuentes. Las comprobaciones técnicas se detallan en el documento técnico.

## 7. Historial

| Iteración | Fecha | Versión | Cambios e impacto |
| --- | --- | --- | --- |
| 001 | 2026-10-29 | 1.0 | Primera versión funcional basada en código y configuración. Se establecen las restricciones de Python y la entrega documental por iteración. Sin cambios en las reglas de negocio implementadas ni ejecución del consolidador. |
| 002 | 2026-10-29 | 1.1 | Incorpora el análisis previo y cambia la clave del cruce IDE00838 a BP.D.B por petición del usuario; con cabecera, Plataforma.Tipo.Nombre. Se mantienen actividad, duplicados y APLICACION_BAJA. Pruebas satisfactorias y contraste de solo lectura; sin regenerar salidas. |
| 003 | 2026-10-29 | 1.2 | Corrige la actividad del formato sin cabecera de la columna 10 a la 11 y acepta únicamente S como activa, según la regla confirmada por el usuario. Mantiene BP.D.B y la regla de APLICACION_BAJA. 14 pruebas satisfactorias y contraste de solo lectura; sin regenerar resultados existentes. |
| 004 | 2026-10-29 | 1.3 | Entrega puntual de CSV con 59 findings de ejecuciones completed sin coincidencia en el Excel existente, conservando la última ejecución y su detalle. Sin cambios funcionales en las reglas existentes. Verificación del cruce, del CSV y 14 pruebas satisfactorias; sin modificar inventario ni fuentes. |
| 005 | 2026-10-29 | 1.4 | Integra en el consolidador la generación automática del CSV fechado y copia ULTIMO para findings de agente sin coincidencia por id, manteniendo la última ejecución y añadiendo detalles del agente/finding. Pruebas sintéticas satisfactorias; no se ejecuta la consolidación real ni se altera el Excel. |
| 006 | 2026-10-29 | 1.5 | Corrige el cuello de botella potencial del CSV de agentes sin inventario mediante índice SQLite sobre `CAST(id AS TEXT)`. Sin cambios funcionales: se mantienen el contenido y las reglas del cruce. Pruebas automatizadas ejecutadas; no se ejecuta consolidación real ni se altera el Excel. |
| 007 | 2026-10-29 | 1.6 | Añade logs detallados fechados de etapas, contadores, duraciones y errores. Sin cambios funcionales en las reglas de consolidación. Verificaciones automatizadas y estáticas; no se ejecuta consolidación real ni se altera el Excel. |
