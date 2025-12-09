📦 AirflowCore – Núcleo del Orquestador Excel Airflow

AirflowCore contiene los dos módulos esenciales que permiten que Excel Airflow funcione como un motor de orquestación completo, inspirado en Apache Airflow pero construido íntegramente en Excel + VBA.

Estos módulos implementan:

El ejecutor principal (dispatcher)

El scheduler tipo cron con reprogramación automática

Marcado de estados (amarillo / verde / rojo)

Logs en tiempo real (Immediate Window)

Llamadas desacopladas a tareas externas (Access, Word, Excel)

Gestión segura de errores

Se trata del “core engine”, la parte más importante del sistema.

🧩 Arquitectura del núcleo

                 ┌────────────────────────────┐
                 │        AirflowCore         │
                 ├──────────────┬─────────────┤
                 │ Módulo 1     │  Módulo 2    │
                 │ Ejecutor     │  Scheduler   │
                 └──────────────┴─────────────┘

Módulo 1 → Ejecuta tareas individuales (ID → tarea).

Módulo 2 → Programa tareas, las relanza con OnTime y activa el Módulo 1.

Ambos módulos trabajan sobre la hoja Interfaz ETL, donde se define:

ID del proceso

Nombre

Estado

Periodicidad (minutos o “daily”)

Celdas de destino para colorear estados

🧱 Módulo 1 – Ejecutor de Tareas (Dispatcher)

Este módulo es el “corazón” de la ejecución:

✔ Traduce un ID de proceso en una tarea real

Cada fila de la interfaz tiene un ID (1–14).
El módulo asigna ese ID a una celda de estado y a una subrutina concreta:

Select Case procesoID
    Case 1:  Tarea_Acceso_Principal
    Case 2:  Tarea_Word_Conversion
    Case 3:  Tarea_Word_Analisis
    ...
Esto permite un sistema tipo DAG, donde cada tarea es independiente.

✔ Actualiza el estado visual en la interfaz

Durante la ejecución:

Amarillo = en progreso

Verde = finalizado con éxito

Rojo = error

targetCell.Interior.Color = RGB(255,255,0)   ' ejecutando
targetCell.Interior.Color = RGB(0,255,0)     ' OK
targetCell.Interior.Color = RGB(255,0,0)     ' error

✔ Ejecuta tareas desacopladas

Cada tarea puede ser:

un módulo Access

un macro Word

un Excel externo

un proceso ETL

un pipeline concreto

Ejemplo genérico:

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase rutaAccess
appAccess.Run "ProcedimientoPrincipal"

Esto permite integrar diferentes herramientas corporativas en un único motor.

✔ Gestión centralizada de errores

Si ocurre un fallo:

el estado pasa a rojo

se muestra un mensaje

se restaura el estado de Excel

With targetCell
    .Value = "Error: " & Now
    .Interior.Color = RGB(255,0,0)
End With

⏱️ Módulo 2 – Scheduler (Programación Automática)

Este módulo convierte Excel Airflow en un orquestador de verdad, capaz de ejecutar procesos programados como si fuera un cron interno.

El scheduler permite:

ejecutar un proceso cada X minutos

ejecutarlo diariamente

relanzarlo automáticamente tras cada ejecución

✔ Procesa una fila completa del panel ETL
ProcesarFilaDeProcesoETL fila


Este método:

Determina el ID.

Llama al dispatcher del Módulo 1.

Marca inicio y fin.

Registra errores si los hay.

Todo queda completamente aislado del motor principal.

✔ Programación automática mediante Application.OnTime

El scheduler calcula la siguiente hora:

horaProxima = DateAdd("n", CInt(periodicidad), Now)
Application.OnTime horaProxima, "EjecutarFilaDeProcesoETLtemp_" & fila


Esto convierte Excel en:

un scheduler recurrente

sin necesidad de que el usuario esté delante

sin complementos de terceros

sin Power Automate, sin Airflow real, sin nada externo

✔ Compatibilidad con tres tipos de periodicidad
Valor en columna “Programación”	Acción del scheduler

off	No se programa

daily	Ejecuta cada día a las 00:00

180 (u otro número)	Ejecuta cada X minutos

                    Usuario / Programación
                               │
                               ▼
                ┌──────────────────────────┐
                │ Scheduler (módulo 2)      │
                │ - Lee periodicidad        │
                │ - Calcula la siguiente    │
                │   ejecución                │
                └────────────┬─────────────┘
                             │
                             ▼
                ┌──────────────────────────┐
                │ Dispatcher (módulo 1)     │
                │ - Identifica tarea        │
                │ - Llama a Tarea_X         │
                │ - Marca estado            │
                └────────────┬─────────────┘
                             │
                             ▼
                 Tareas externas (Access, Word,
                      Excel, ETL corporativos)

La separación de responsabilidades lo hace estable, mantenible y muy fácil de ampliar.

🧠 Por qué este diseño funciona tan bien

✔ Cada tarea es desacoplada → se puede modificar sin romper otras.
✔ El scheduler no necesita saber qué hace cada tarea.
✔ El dispatcher no necesita saber cuándo debe ejecutarse.
✔ Las celdas de estado mantienen una UI clara y visual.
✔ Es un patrón muy parecido al de Airflow real:

Scheduler

Executor

Tasks

Logs visuales

🧩 Extensibilidad

Puedes añadir nuevas tareas simplemente:

Añadiendo un nuevo ID en la interfaz

Creando una nueva Tarea_Nueva

Añadiendo una línea en el Select Case

Ejemplo:
Case 15: Tarea_NuevoProceso

El sistema crece sin modificar la arquitectura.

