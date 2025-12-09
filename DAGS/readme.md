🧩 Cómo crear un DAG en Excel Airflow

Excel Airflow permite definir DAGs (Directed Acyclic Graphs) de forma muy similar a Apache Airflow, pero usando VBA + Excel.
Un DAG no es más que un conjunto de tareas que se ejecutan en un orden determinado, con dependencias entre ellas y un control visual del estado de ejecución.

Este documento explica cómo crear tu propio DAG dentro de la carpeta DAGS/.

📘 ¿Qué es un DAG en Excel Airflow?

Un DAG en este sistema es:

un módulo VBA (.bas)

que contiene una función principal (por ejemplo, Sub MiDAG())

dentro de la cual se definen tareas mediante llamadas a Task o a funciones propias

que el executor (del módulo 1 de AirflowCore) ejecutará cuando el usuario pulse EJECUTAR o cuando el scheduler lo programe

La filosofía es:

DAG = Lista de tareas + Dependencias + Lógica propia

✔️ Estructura recomendada de un DAG

Un módulo .bas con:

Sub NombreDelDAG()
    ' Definición de dependencias y tareas
End Sub

' Implementación de tareas
Sub Tarea1()
End Sub

Sub Tarea2()
End Sub

🧱 Paso 1: Crear un nuevo módulo dentro de /DAGS/

En Excel → ALT + F11

Insertar → Módulo

Guardarlo como:

/DAGS/MiPrimerDAG.bas

🧩 Paso 2: Definir la función principal del DAG

Esta función es el punto de entrada del DAG.

Ejemplo:

🧩 Paso 2: Definir la función principal del DAG

Esta función es el punto de entrada del DAG.

Ejemplo:

Sub MiPrimerDAG()

    ' Definir las tareas con sus dependencias
    Call EjecutarTarea("ExtraccionA", "N/A")
    Call EjecutarTarea("LimpiezaA", "ExtraccionA")
    Call EjecutarTarea("CargaA", "LimpiezaA")

End Sub

En Excel Airflow, el orquestador interpreta esto como:

ExtraccionA → LimpiezaA → CargaA

🧠 Paso 3: Crear las tareas

Cada tarea es simplemente una macro VBA que ejecuta algo:

Sub ExtraccionA()
    ' Ejemplo: importar un fichero
End Sub

Sub LimpiezaA()
    ' Ejemplo: eliminar duplicados
End Sub

Sub CargaA()
    ' Ejemplo: cargar datos en Access
End Sub

Las tareas son independientes, igual que en Airflow real.

🔧 Paso 4: Asociar las tareas al sistema de ejecución

Excel Airflow mantiene un dispatcher que ejecuta tareas según su nombre.

Si usas llamadas estilo:

Call EjecutarTarea("NombreTarea", "Dependencia")

el orquestador:

Reconoce la dependencia

Ordena el flujo

Marca el estado en la interfaz

Llama a la subrutina correspondiente

Registra el resultado

🔄 Paso 5: Añadir el DAG al Panel de Control

En la hoja Interfaz ETL:

Crear una nueva fila

Asignar un ID libre (por ejemplo, 15)

En la columna Proceso, escribir:
EjecutarMiPrimerDAG
En la columna de periodicidad:

off para manual

número (minutos) para ejecución recurrente

daily para ejecución diaria


🧰 Buenas prácticas para crear DAGs en Excel Airflow
✔ Mantén cada tarea pequeña y clara

Igual que en Airflow: una tarea = una función bien definida.

✔ Usa nombres neutros

Evita nombres con datos internos o procesos reales si el repositorio es público.

✔ Loggea tiempo y errores dentro de cada tarea

Súper útil para debugging.

✔ Los DAGs no deben contener lógica compleja

La lógica debe vivir dentro de las tareas.

✔ Evita que los DAGs modifiquen configuración del motor

El DAG define qué se hace, el motor define cómo se ejecuta.

🚀 Añadir un DAG al scheduler

Si quieres que tu DAG se ejecute solo:

Escribe daily en la columna de periodicidad

O un número en minutos (ej: 60 → cada hora)

El scheduler del Módulo 2 lo añadirá automáticamente con OnTime

📄 En resumen

Un DAG en Excel Airflow no es más que:

Un módulo VBA

Con una lista de tareas y dependencias

Que el motor ejecuta y marca visualmente

Con opción de programación automática

Con esto tienes un sistema de orquestación 100% funcional, 100% Excel, 100% corporativo-friendly, sin necesidad de software externo.
