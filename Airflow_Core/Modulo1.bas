Option Explicit

' ==========================================
'   ORQUESTADOR PRINCIPAL - EXCEL AIRFLOW
'   Versión anonimizada para ejemplo público
' ==========================================

Sub EjecutarProcesoETL()
    Dim fila As Integer, procesoID As Integer
    Dim targetCell As Range, ws As Worksheet

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Set ws = ThisWorkbook.Sheets("Interfaz ETL")
    On Error GoTo ErrorHandler

    fila = ActiveCell.Row
    procesoID = ws.Cells(fila, 1).Value

    Select Case procesoID
        Case 1: Set targetCell = ws.Range("D3")
        Case 2: Set targetCell = ws.Range("D4")
        Case 3: Set targetCell = ws.Range("D5")
        Case 4: Set targetCell = ws.Range("D6")
        Case 5: Set targetCell = ws.Range("D7")
        Case 6: Set targetCell = ws.Range("D8")
        Case 7: Set targetCell = ws.Range("D9")
        Case 8: Set targetCell = ws.Range("D10")
        Case 9: Set targetCell = ws.Range("D11")
        Case 10: Set targetCell = ws.Range("D12")
        Case 11: Set targetCell = ws.Range("D13")
        Case 12: Set targetCell = ws.Range("D14")
        Case 13: Set targetCell = ws.Range("D15")
        Case 14: Set targetCell = ws.Range("D16")
        Case Else
            MsgBox "Proceso desconocido o no implementado.", vbExclamation
            GoTo Limpieza
    End Select

    ' Mostrar amarillo (en ejecución)
    Application.ScreenUpdating = True
    With targetCell
        .Value = "Iniciado: " & Now
        .Interior.Color = RGB(255, 255, 0)
    End With
    Application.ScreenUpdating = False

    DoEvents

    ' Dispatcher: ejecuta la tarea correspondiente al ID
    Select Case procesoID
        Case 1:  Tarea_Access_Principal
        Case 2:  Tarea_Word_Conversion
        Case 3:  Tarea_Word_Analisis
        Case 4:  Tarea_Access_Calculos
        Case 5:  Tarea_Access_Importacion1
        Case 6:  Tarea_Word_Analisis2
        Case 7:  Tarea_Access_Importacion2
        Case 8:  Tarea_Access_Importacion3
        Case 9:  Tarea_Word_ProcesoAvanzado
        Case 10: Tarea_Access_Pipeline1
        Case 11: Tarea_Access_Notificaciones
        Case 12: Tarea_Access_ActualizacionPlazos
        Case 13: Tarea_Access_ActualizacionReferencias
        Case 14: Tarea_Excel_DAG_Externo
    End Select

    ' Mostrar verde (finalizado bien)
    Application.ScreenUpdating = True
    With targetCell
        .Value = "Última ejecución: " & Now
        .Interior.Color = RGB(0, 255, 0)
    End With
    Application.ScreenUpdating = False

Limpieza:
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Exit Sub

ErrorHandler:
    ' Mostrar rojo (error)
    Application.ScreenUpdating = True
    With targetCell
        .Value = "Error: " & Now
        .Interior.Color = RGB(255, 0, 0)
    End With
    Application.ScreenUpdating = False

    MsgBox "Se produjo un error al ejecutar el proceso: " & Err.Description, vbCritical
    Resume Limpieza

End Sub

' =====================================================
'   TAREAS (EJEMPLOS ANONIMIZADOS)
'   Cada tarea llama a Access / Word / Excel externos
' =====================================================

' Tarea 1: Ejecutar un módulo de Access protegido
Sub Tarea_Acceso_Principal()
    Dim appAccess As Object
    Dim rutaAccess As String
    Dim pwd As String

    rutaAccess = "C:\RUTA\A\TU_BASE\base_principal.accdb"
    pwd = "TU_PASSWORD_AQUI"  ' <-- reemplazar en entorno real

    Set appAccess = CreateObject("Access.Application")
    appAccess.OpenCurrentDatabase rutaAccess, False, pwd
    appAccess.Run "ProcedimientoPrincipal"  ' nombre genérico
    appAccess.Quit
    Set appAccess = Nothing
End Sub

' Tarea 2: Ejecutar macro de Word para conversión de archivos
Sub Tarea_Word_Conversion()
    Dim appWord As Object
    Dim rutaWord As String

    rutaWord = "C:\RUTA\A\TUS_DOCS\documento_con_macro.docm"

    Set appWord = CreateObject("Word.Application")
    appWord.Documents.Open rutaWord
    appWord.Run "MacroConversionDocumentos"  ' nombre genérico
    appWord.ActiveDocument.Close False
    appWord.Quit
    Set appWord = Nothing
End Sub

' Tarea 3: Ejecutar macro de Word para análisis
Sub Tarea_Word_Analisis()
    Dim appWord As Object
    Dim rutaWord As String

    rutaWord = "C:\RUTA\A\TUS_DOCS\analisis.docm"

    Set appWord = CreateObject("Word.Application")
    appWord.Documents.Open rutaWord
    appWord.Run "MacroAnalisisDocumentos"  ' nombre genérico
    appWord.ActiveDocument.Close False
    appWord.Quit
    Set appWord = Nothing
End Sub

' Tarea 4: Ejecutar módulo de Access (cálculos)
Sub Tarea_Acceso_Calculos()
    Dim appAccess As Object
    Dim rutaAccess As String

    rutaAccess = "C:\RUTA\A\TU_BASE\calculos.accdb"

    Set appAccess = CreateObject("Access.Application")
    appAccess.OpenCurrentDatabase rutaAccess
    appAccess.Run "ProcedimientoCalculos"  ' nombre genérico
    appAccess.Quit
    Set appAccess = Nothing
End Sub

' Tarea 5: Importaciones / ETL 1 en Access
Sub Tarea_Acceso_Importacion1()
    Dim appAccess As Object
    Dim rutaAccess As String

    rutaAccess = "C:\RUTA\A\TU_BASE\etl_import1.accdb"

    Set appAccess = CreateObject("Access.Application")
    appAccess.OpenCurrentDatabase rutaAccess
    appAccess.Run "ProcedimientoImportacion1"
    appAccess.Quit
    Set appAccess = Nothing
End Sub

' Tarea 6: Segundo análisis en Word
Sub Tarea_Word_Analisis2()
    Dim appWord As Object
    Dim rutaWord As String

    rutaWord = "C:\RUTA\A\TUS_DOCS\analisis_avanzado.docm"

    Set appWord = CreateObject("Word.Application")
    appWord.Documents.Open rutaWord
    appWord.Run "MacroAnalisisAvanzado"
    appWord.ActiveDocument.Close False
    appWord.Quit
    Set appWord = Nothing
End Sub

' Tarea 7: Importaciones / ETL 2 en Access
Sub Tarea_Acceso_Importacion2()
    Dim appAccess As Object
    Dim rutaAccess As String

    rutaAccess = "C:\RUTA\A\TU_BASE\etl_import2.accdb"

    Set appAccess = CreateObject("Access.Application")
    appAccess.OpenCurrentDatabase rutaAccess
    appAccess.Run "ProcedimientoImportacion2"
    appAccess.Quit
    Set appAccess = Nothing
End Sub

' Tarea 8: Importaciones / ETL 3 en Access
Sub Tarea_Acceso_Importacion3()
    Dim appAccess As Object
    Dim rutaAccess As String

    rutaAccess = "C:\RUTA\A\TU_BASE\etl_import3.accdb"

    Set appAccess = CreateObject("Access.Application")
    appAccess.OpenCurrentDatabase rutaAccess
    appAccess.Run "ProcedimientoImportacion3"
    appAccess.Quit
    Set appAccess = Nothing
End Sub

' Tarea 9: Proceso avanzado en Word
Sub Tarea_Word_ProcesoAvanzado()
    Dim appWord As Object
    Dim rutaWord As String

    rutaWord = "C:\RUTA\A\TUS_DOCS\proceso_avanzado.docm"

    Set appWord = CreateObject("Word.Application")
    appWord.Documents.Open rutaWord
    appWord.Run "MacroProcesoAvanzado"
    appWord.ActiveDocument.Close False
    appWord.Quit
    Set appWord = Nothing
End Sub

' Tarea 10: Pipeline 1 en Access
Sub Tarea_Acceso_Pipeline1()
    Dim appAccess As Object
    Dim rutaAccess As String

    rutaAccess = "C:\RUTA\A\TU_BASE\pipeline1.accdb"

    Set appAccess = CreateObject("Access.Application")
    appAccess.OpenCurrentDatabase rutaAccess
    appAccess.Run "EjecutarPipeline1"
    appAccess.Quit
    Set appAccess = Nothing
End Sub

' Tarea 11: Envío de notificaciones desde Access
Sub Tarea_Acceso_Notificaciones()
    Dim appAccess As Object
    Dim rutaAccess As String

    rutaAccess = "C:\RUTA\A\TU_BASE\notificaciones.accdb"

    Set appAccess = CreateObject("Access.Application")
    appAccess.OpenCurrentDatabase rutaAccess
    appAccess.Run "EnviarNotificaciones"
    appAccess.Quit
    Set appAccess = Nothing
End Sub

' Tarea 12: Actualización de plazos / datos en Access
Sub Tarea_Acceso_ActualizacionPlazos()
    Dim appAccess As Object
    Dim rutaAccess As String

    rutaAccess = "C:\RUTA\A\TU_BASE\actualizacion_plazos.accdb"

    Set appAccess = CreateObject("Access.Application")
    appAccess.OpenCurrentDatabase rutaAccess
    appAccess.Run "ActualizarPlazos"
    appAccess.Quit
    Set appAccess = Nothing
End Sub

' Tarea 13: Actualización de referencias en Access
Sub Tarea_Acceso_ActualizacionReferencias()
    Dim appAccess As Object
    Dim rutaAccess As String

    rutaAccess = "C:\RUTA\A\TU_BASE\actualizacion_referencias.accdb"

    Set appAccess = CreateObject("Access.Application")
    appAccess.OpenCurrentDatabase rutaAccess
    appAccess.Run "ActualizarReferencias"
    appAccess.Quit
    Set appAccess = Nothing
End Sub

' Tarea 14: Ejecutar un DAG en un Excel externo
Sub Tarea_Excel_DAG_Externo()
    Dim appExcel As Object
    Dim rutaExcel As String

    rutaExcel = "C:\RUTA\A\TUS_EXCEL\dag_ejemplo.xlsm"

    Set appExcel = CreateObject("Excel.Application")
    appExcel.Workbooks.Open rutaExcel
    appExcel.Run "EjecutarPipelineEjemplo"  ' macro del libro externo
    appExcel.Quit
    Set appExcel = Nothing
End Sub
