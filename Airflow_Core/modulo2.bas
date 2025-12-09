Option Explicit

Dim filaActual As Integer

' Helpers para programar filas concretas (ejemplos)
Public Sub EjecutarFilaDeProcesoETLtemp_3()
    EjecutarFilaDeProcesoETL 3
End Sub

Public Sub EjecutarFilaDeProcesoETLtemp_6()
    EjecutarFilaDeProcesoETL 6
End Sub

Public Sub EjecutarFilaDeProcesoETLtemp_10()
    EjecutarFilaDeProcesoETL 10
End Sub

' ===========================================================
'  PROGRAMADOR: EJECUTA Y REPROGRAMA UNA FILA DEL PANEL ETL
' ===========================================================
Public Sub EjecutarFilaDeProcesoETL(fila As Integer)
    Dim ws As Worksheet
    Dim periodicidad As Variant
    Dim horaProxima As Date

    Set ws = ThisWorkbook.Sheets("Interfaz ETL")

    If fila <= 0 Then
        MsgBox "Fila no configurada correctamente.", vbCritical
        Exit Sub
    End If

    periodicidad = ws.Cells(fila, 6).Value   ' Columna de programación

    If periodicidad <> "" And LCase(periodicidad) <> "off" Then
        ProcesarFilaDeProcesoETL fila

        ' Recalcular la próxima hora de ejecución
        If LCase(periodicidad) = "daily" Then
            horaProxima = DateAdd("d", 1, Date) + TimeValue("00:00:00")
        ElseIf IsNumeric(periodicidad) Then
            horaProxima = DateAdd("n", CInt(periodicidad), Now)
        End If
        
        ' Reprogramar usando OnTime
        Debug.Print "Reprogramando fila: " & fila & " para hora: " & horaProxima
        Application.OnTime horaProxima, "'" & ThisWorkbook.Name & "'!EjecutarFilaDeProcesoETLtemp_" & fila, , True
    End If
End Sub

' ===========================================================
'  EJECUCIÓN DE UNA FILA: CAMBIA ESTADO Y DESPACHA LA TAREA
' ===========================================================
Public Sub ProcesarFilaDeProcesoETL(fila As Integer)
    Dim ws As Worksheet
    Dim procesoID As Integer
    Dim targetCell As Range

    Set ws = ThisWorkbook.Sheets("Interfaz ETL")

    ' Id del proceso en la columna 1
    procesoID = ws.Cells(fila, 1).Value

    ' Celda de estado asociada al proceso
    Select Case procesoID
        Case 1:  Set targetCell = ws.Range("D3")
        Case 2:  Set targetCell = ws.Range("D4")
        Case 3:  Set targetCell = ws.Range("D5")
        Case 4:  Set targetCell = ws.Range("D6")
        Case 5:  Set targetCell = ws.Range("D7")
        Case 6:  Set targetCell = ws.Range("D8")
        Case 7:  Set targetCell = ws.Range("D9")
        Case 8:  Set targetCell = ws.Range("D10")
        Case 9:  Set targetCell = ws.Range("D11")
        Case 10: Set targetCell = ws.Range("D12")
        Case 11: Set targetCell = ws.Range("D13")
        Case 12: Set targetCell = ws.Range("D14")
        Case Else
            MsgBox "Proceso desconocido o no implementado.", vbExclamation
            Exit Sub
    End Select

    ' Registrar inicio (amarillo)
    Debug.Print "Iniciado proceso ID: " & procesoID & " en fila: " & fila & _
                " a las " & Format(Now, "dd/mm/yyyy hh:nn:ss")
    With targetCell
        .Value = "Iniciado: " & Format(Now, "dd/mm/yyyy hh:nn:ss")
        .Interior.Color = RGB(255, 255, 0)
    End With

    ' Despachar tarea según el ID
    On Error GoTo ErrorHandler
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
        Case 11: Tarea_Access
        Case 12: Tarea_Access
    End Select

    ' Registrar finalización correcta (verde)
    With targetCell
        .Value = "Última ejecución: " & Format(Now, "dd/mm/yyyy hh:nn:ss")
        .Interior.Color = RGB(0, 255, 0)
    End With

    Exit Sub

ErrorHandler:
    ' Registrar error (rojo)
    With targetCell
        .Value = "Error: " & Format(Now, "dd/mm/yyyy hh:nn:ss")
        .Interior.Color = RGB(255, 0, 0)
    End With
    MsgBox "Se produjo un error al ejecutar el proceso: " & Err.Description, vbCritical
End Sub
