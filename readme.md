Excel Airflow es un motor de orquestación desarrollado en VBA, capaz de ejecutar flujos de trabajo complejos (DAGs), programar tareas, gestionar dependencias, colorear estados y registrar logs… todo dentro de Microsoft Excel, sin necesidad de Python, servidores externos ni permisos corporativos.

Nació como una solución para entornos con restricciones técnicas donde no se puede instalar software externo, pero donde sí existe la necesidad de automatizar procesos de datos reales.

💡 Es, en esencia, un Apache Airflow operativo dentro de Excel.

⚙️ Instalación

1. Abre Excel.

2. Pulsa ALT + F11 para abrir el editor de VBA.

3. Menú Archivo → Importar archivo…

4. Importa los módulos del directorio: AirflowCore/

5. Guarda el libro como .xlsm.

6. Dale formato a la hoja de excel siguiendo la imagen de /assets

🏗️ Cómo funciona Excel Airflow

Excel Airflow implementa un sistema completo de orquestación:

✔ DAGs

Cada DAG es un módulo .bas con tareas definidas mediante funciones o subrutinas.

✔ Scheduler

Una función interna reconstruye el grafo, valida dependencias y ejecuta las tareas en orden.

✔ Estados de ejecución

Las tareas se colorean automáticamente:

🟩 Correcto

🟥 Error

🟧 En ejecución

✔ Logs

Registra cada evento con fecha, tarea y duración.

✔ Integración con otras herramientas

Puede llamar:

Macros de Excel

Scripts externos

Módulos de Access

Macros de Word

Procesos ETL internos

💡 Motivación del proyecto

Excel Airflow se creó para dar solución a un problema muy habitual en empresas con fuerte bloqueo tecnológico:

No se permite Python

No se permite instalar librerías

No se permite conectarse a servidores externos

Pero sí se necesita automatizar procesos de datos reales

Este framework permite construir pipelines reproducibles, organizadas y profesionales, usando únicamente Excel, algo que se encuentra en prácticamente cualquier entorno corporativo.

🧪 Estado actual

✔ Motor funcional

✔ Scheduler estable

✔ Soporte para dependencias

✔ Colores y logs

✔ DAGs de ejemplo

⏳ Documentación ampliada (en desarrollo)

⏳ Ejemplos avanzados

🤝 Contribuir

Las contribuciones son bienvenidas:

Crear branches específicas

Abrir issues con mejoras

Enviar PRs con ejemplos de DAGs o mejoras en el motor

📄 Licencia

MIT License.
