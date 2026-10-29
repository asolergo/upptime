# VulGest — Documento funcional

- Versión documental: 1.0
- Iteración: 001
- Fecha: 2026-10-29
- Código de referencia: `consolidar_inventario.py`, versión `2026.10.28.24-clave-unica-bp-b-d`.
- Base de esta versión: revisión estática del código y de la configuración; no se ha ejecutado la consolidación.

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

### RF-05. Partner

Se busca `endpoint__product__main_provider` en el mapeo de partners, sin distinguir mayúsculas y eliminando espacios exteriores. Si no existe agrupación, se conserva el proveedor original. Si el proveedor está vacío, `partner` queda vacío.

### RF-06. Aplicación de baja

En IDE00838 se forma exclusivamente la clave `Application - <BP>.<B>.<D>`, donde BP es plataforma, B es nombre y D es tipo.

- Con cabecera reconocida se usan `Nombre`, `Tipo`, `Activa` y `Plataforma`; si falta esta última, se usa la posición BP.
- Sin cabecera reconocida se usan B (2), D (4), J (10, activa) y BP (68). Las filas cortas o sin plataforma, nombre o tipo se omiten.
- La normalización aplica Unicode NFKC, elimina BOM, normaliza espacios y no distingue mayúsculas. No se buscan sufijos ni alias.
- Se considera activa si el valor, recortado y en mayúsculas, es `S`, `SI`, `SÍ`, `Y`, `YES`, `TRUE` o `1`. Cualquier otro valor se interpreta como inactiva.
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
- Se generan un informe `<nombre_salida>_validacion.txt`, un log de ejecución y logs de resumen y detalle del cruce IDE00838 en `logs/`, con copias `ULTIMO`/`ULTIMA` según el archivo.
- La consola informa registros leídos, descartados, inserciones/actualizaciones, coincidencias y duración.
- Si falta una fuente obligatoria, la configuración es inválida o no se puede escribir la salida, la ejecución informa un error. No se garantiza que todo fallo deje intactas todas las salidas: algunos controles son posteriores al reemplazo y los logs se escriben durante el proceso.

## 5. Reglas de trabajo de OpenCode

- Solo se permite la biblioteca estándar de Python; está prohibido instalar o utilizar paquetes externos, incluso preinstalados.
- Cada iteración de OpenCode debe entregar actualizados este documento y `docs/tecnico.md`, con historial común. No significa generar documentos por cada fila procesada o ejecución del consolidador.
- La primera iteración establece la línea base documental; las posteriores deben reflejar cambios o declarar expresamente la ausencia de impacto.
- Las instrucciones persistentes y la lista de comprobación están en `AGENTS.md`.

## 6. Estado de validación

La línea base describe el comportamiento inspeccionado, no una ejecución certificada con los datos actuales. `Inventario_validacion.txt` declara la versión `2026.10.28.21-ide00838-aliases`, anterior a la del código revisado; no valida la versión actual. En esta iteración no se ha regenerado ni validado el Excel existente. Las comprobaciones documentales y estáticas se detallan en el documento técnico.

## 7. Historial

| Iteración | Fecha | Versión | Cambios e impacto |
| --- | --- | --- | --- |
| 001 | 2026-10-29 | 1.0 | Primera versión funcional basada en código y configuración. Se establecen las restricciones de Python y la entrega documental por iteración. Sin cambios en las reglas de negocio implementadas ni ejecución del consolidador. |
