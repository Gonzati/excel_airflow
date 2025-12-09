' Ejemplo de DAG:
' Limpia dos ficheros de correspondencias (fuentes A y B),
' elimina duplicados, renombra columnas y los carga en una base
' de datos Access para su posterior cruce.

Sub LimpiarYCargarFuentesAB()

    Dim dbPath As String
    Dim folderPath As String
    Dim accessApp As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim fso As Object
    Dim logFile As String
    Dim logNumber As Integer
    Dim fileSrcA As String, fileSrcB As String
    Dim importOK As Boolean
    Dim startTime As Double, stepTime As Double
    Dim fechaHoy As String
    
    ' ================================
    '  CONFIGURACIÓN DEL PROCESO
    ' ================================
    folderPath = "C:\ETL\staging\"                  ' Carpeta de los Excel de entrada
    dbPath = "C:\ETL\db\StagingDB.accdb"            ' Ruta de la base de datos Access
    fileSrcA = folderPath & "SOURCE_A.xlsx"
    fileSrcB = folderPath & "SOURCE_B.xlsx"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fechaHoy = Format(Date, "yyyymmdd")
    logFile = folderPath & "log_carga_" & fechaHoy & ".txt"
    logNumber = FreeFile
    
    ' ================================
    '  INICIO DEL LOG
    ' ================================
    Open logFile For Append As #logNumber
    Print #logNumber, vbCrLf & "==== Proceso iniciado: " & Now & " ===="
    startTime = Timer
    
    ' Verificar existencia de archivos
    If Not fso.FileExists(fileSrcA) Then
        Print #logNumber, "ERROR: No existe SOURCE_A.xlsx"
        Close #logNumber
        MsgBox "No existe SOURCE_A.xlsx", vbCritical
        Exit Sub
    End If
    
    If Not fso.FileExists(fileSrcB) Then
        Print #logNumber, "ERROR: No existe SOURCE_B.xlsx"
        Close #logNumber
        MsgBox "No existe SOURCE_B.xlsx", vbCritical
        Exit Sub
    End If
    
    ' ==============================================
    '  LIMPIEZA Y NORMALIZACIÓN DE SOURCE_A
    ' ==============================================
    stepTime = Timer
    Set wb = Workbooks.Open(fileSrcA)
    Set ws = wb.Sheets(1)
    
    ' Eliminamos duplicados por las dos primeras columnas
    ws.UsedRange.RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
    
    ' Renombramos las columnas clave
    ws.Cells(1, 1).Value = "Key_A"
    ws.Cells(1, 2).Value = "Key_B_SourceA"
    
    wb.Save
    wb.Close
    Print #logNumber, "Duplicados eliminados y columnas renombradas en SOURCE_A.xlsx (" & _
                      Format(Timer - stepTime, "0.00") & " seg)"
    
    ' ==============================================
    '  LIMPIEZA Y NORMALIZACIÓN DE SOURCE_B
    ' ==============================================
    stepTime = Timer
    Set wb = Workbooks.Open(fileSrcB)
    Set ws = wb.Sheets(1)
    
    ws.UsedRange.RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
    ws.Cells(1, 1).Value = "Key_A"
    ws.Cells(1, 2).Value = "Key_B_SourceB"
    
    wb.Save
    wb.Close
    Print #logNumber, "Duplicados eliminados y columnas renombradas en SOURCE_B.xlsx (" & _
                      Format(Timer - stepTime, "0.00") & " seg)"
    
    ' ==============================================
    '  APERTURA DE ACCESS
    ' ==============================================
    stepTime = Timer
    Set accessApp = CreateObject("Access.Application")
    accessApp.OpenCurrentDatabase dbPath
    Print #logNumber, "Base de datos abierta (" & Format(Timer - stepTime, "0.00") & " seg)"
    
    ' ==============================================
    '  CONSULTAS PREVIAS (TRUNCATE TABLAS STAGING)
    ' ==============================================
    On Error Resume Next
    accessApp.DoCmd.OpenQuery "TRUNCATE_Stage_SourceB"
    If Err.Number = 0 Then
        Print #logNumber, "TRUNCATE_Stage_SourceB ejecutada"
    Else
        Print #logNumber, "TRUNCATE_Stage_SourceB falló: " & Err.Description
        Err.Clear
    End If
    
    accessApp.DoCmd.OpenQuery "TRUNCATE_Stage_SourceA"
    If Err.Number = 0 Then
        Print #logNumber, "TRUNCATE_Stage_SourceA ejecutada"
    Else
        Print #logNumber, "TRUNCATE_Stage_SourceA falló: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    
    ' ==============================================
    '  IMPORTACIÓN DE SOURCE_A
    ' ==============================================
    stepTime = Timer
    importOK = True
    On Error Resume Next
    accessApp.DoCmd.TransferSpreadsheet _
        TransferType:=0, _
        SpreadsheetType:=10, _
        TableName:="Stage_SourceA", _
        FileName:=fileSrcA, _
        HasFieldNames:=True
    
    If Err.Number <> 0 Then
        Print #logNumber, "ERROR importando SOURCE_A.xlsx: " & Err.Description
        importOK = False
        Err.Clear
    Else
        Print #logNumber, "SOURCE_A.xlsx importado correctamente (" & _
                          Format(Timer - stepTime, "0.00") & " seg)"
    End If
    
    ' ==============================================
    '  IMPORTACIÓN DE SOURCE_B
    ' ==============================================
    stepTime = Timer
    accessApp.DoCmd.TransferSpreadsheet _
        TransferType:=0, _
        SpreadsheetType:=10, _
        TableName:="Stage_SourceB", _
        FileName:=fileSrcB, _
        HasFieldNames:=True
    
    If Err.Number <> 0 Then
        Print #logNumber, "ERROR importando SOURCE_B.xlsx: " & Err.Description
        importOK = False
        Err.Clear
    Else
        Print #logNumber, "SOURCE_B.xlsx importado correctamente (" & _
                          Format(Timer - stepTime, "0.00") & " seg)"
    End If
    On Error GoTo 0
    
    accessApp.Quit
    Set accessApp = Nothing
    Print #logNumber, "Base de datos cerrada"
    
    ' ==============================================
    '  LIMPIEZA DE FICHEROS STAGING
    ' ==============================================
    If importOK Then
        On Error Resume Next
        Kill fileSrcA
        Kill fileSrcB
        If Err.Number = 0 Then
            Print #logNumber, "Archivos Excel eliminados"
        Else
            Print #logNumber, "ERROR al eliminar Excel: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Else
        Print #logNumber, "Archivos Excel NO eliminados por error en importación"
    End If
    
    ' ==============================================
    '  FIN DEL LOG
    ' ==============================================
    Print #logNumber, "Tiempo total: " & Format(Timer - startTime, "0.00") & " seg"
    Print #logNumber, "==== Proceso finalizado: " & Now & " ===="
    Close #logNumber

End Sub
