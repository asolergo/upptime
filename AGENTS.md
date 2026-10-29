# Instrucciones de OpenCode para VulGest

Estas reglas se aplican a todo el proyecto y a cada iteración de trabajo de OpenCode.

## 1. Python: únicamente biblioteca estándar

- Se permite utilizar Python y los módulos incluidos en su biblioteca estándar.
- Está prohibido instalar, importar o utilizar paquetes Python de terceros, aunque ya estén instalados en el equipo.
- La prohibición afecta al código del proyecto, pruebas, scripts auxiliares, validaciones y herramientas Python que ejecute OpenCode, directa o indirectamente.
- No ejecutar instalaciones mediante `pip`, `pipx`, `uv pip`, Poetry, Conda ni mecanismos equivalentes. No descargar ni copiar paquetes para eludir esta regla.
- Resolver las necesidades con la biblioteca estándar; por ejemplo, `csv`, `json`, `sqlite3`, `zipfile`, `unittest` y `xml.etree.ElementTree`. No usar, por ejemplo, pandas, openpyxl, numpy ni pytest.
- Si una tarea no puede resolverse bajo esta restricción, explicar el impedimento y pedir una decisión al usuario; no instalar dependencias por iniciativa propia.
- No eliminar código o dependencias preexistentes sin autorización: señalar cualquier incumplimiento encontrado.

## 2. Dos documentos obligatorios en cada iteración

Los entregables documentales permanentes son:

1. `docs/funcional.md`: objetivos, entradas, salidas, comportamientos, reglas de negocio, validaciones funcionales y limitaciones.
2. `docs/tecnico.md`: arquitectura, componentes, configuración, formatos, dependencias, ejecución, comprobaciones técnicas y limitaciones.

Se entiende por iteración una tarea solicitada por el usuario que termina con una entrega de OpenCode, no cada llamada a una herramienta ni cada ejecución del consolidador.

- En la primera iteración, crear la versión inicial de ambos documentos a partir del código y la configuración realmente inspeccionados.
- En cada iteración posterior, revisar y actualizar ambos documentos antes de dar la tarea por terminada, incluso si la tarea solo afecta a documentación o configuración.
- Mantener estos dos archivos como documentos vivos; no crear copias nuevas por cada iteración salvo petición expresa del usuario.
- Registrar en ambos una entrada de historial con el mismo identificador de iteración, fecha, cambios e impacto. Si no existe impacto en un ámbito, registrar explícitamente «Sin cambios funcionales» o «Sin cambios técnicos», según corresponda, explicando el motivo.
- Mantener coherentes las reglas descritas con el comportamiento implementado. Distinguir hechos verificados, resultados de pruebas, evidencias antiguas y aspectos pendientes.
- No afirmar que una prueba se ha ejecutado cuando solo se ha revisado el código. Documentar las comprobaciones realizadas y las que no se ejecutaron.
- No incluir datos de inventario, secretos ni detalles sensibles innecesarios en la documentación.
- Si el modo de trabajo o una instrucción de mayor prioridad impide escribir archivos, respetar esa limitación e indicar que la actualización documental queda pendiente; no presentarla como completada.

## 3. Cierre y protección del trabajo existente

- Antes de cerrar una iteración, comprobar que ambos documentos existen, reflejan la tarea y contienen su entrada de historial.
- En la respuesta final, mencionar los documentos actualizados y resumir las verificaciones o impedimentos relevantes.
- No regenerar `Inventario.xlsx`, ejecutar parches históricos ni sobrescribir datos o resultados existentes como efecto secundario de tareas documentales.
- Estas instrucciones son reglas de actuación del agente; no constituyen un aislamiento técnico del sistema operativo.
