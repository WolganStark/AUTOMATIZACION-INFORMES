Attribute VB_Name = "Módulo2"
Option Explicit

' ==================================
' CONFIGURACIÓN (Ajuste los nombres)
' ==================================

Private Const TABLE_NAME As String = "Pacientes"           'Nombre de la tabla
Private Const EVENTS_SHEET As String = "Eventos_Detallados" 'Nombre de la hoja en donde se normaliza para la creación del informe
Private Const MAX_EVENTS As Long = 10                      'Cantidad máxima de eventos

'===================================

' =============Ayudas===============
Private Function WorksheetExists(sName As String) As Boolean
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(sName)
    WorksheetExists = Not sh Is Nothing
    On Error GoTo 0
End Function

Private Function GetListObjectByName(ByVal tblName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set lo = ws.ListObjects(tblName)
        On Error GoTo 0
        If Not lo Is Nothing Then
            Set GetListObjectByName = lo
            Exit Function
        End If
    Next ws
End Function

'----------------- 1) Genera la hoja Eventos_Detallados -----------------
Public Sub BuildEventDetail()
    Dim lo As ListObject
    Set lo = GetListObjectByName(TABLE_NAME)
    If lo Is Nothing Then
        MsgBox "No se encontró la tabla llamada '" & TABLE_NAME & "'. Asegurate de nombrarla correctamente.", vbExclamation
        Exit Sub
    End If
        
    Dim wsE As Worksheet
    If WorksheetExists(EVENTS_SHEET) Then
        Set wsE = ThisWorkbook.Worksheets(EVENTS_SHEET)
        wsE.Cells.Clear
    Else
        Set wsE = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsE.Name = EVENTS_SHEET
    End If
        
    '--------Columnas destino--------
    Dim headers As Variant
    headers = Array("TipoDocumento", "NumeroDocumento", "Nombre", "Apellido", "FechaTransplante", "Tipo_Evento", "Fecha_Evento", "Codigo_Evento", "Fase_Evento", _
                    "Ano", "NumeroMes", "Mes", "NumeroTrimestre", "EtiquetaTrimestre", "AnoMes")
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        wsE.Cells(1, i + 1).Value = headers(i)
    Next i
    
    '---------------Se mapean encabezados de la tabla origen---------------
    Dim colMap As Object: Set colMap = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 1 To lo.ListColumns.Count
        colMap(lo.ListColumns(c).Name) = c
    Next c
    '-------Nombres de Columnas --------
    Dim colTipoDocumento As Long: colTipoDocumento = 0
    Dim colNumeroDocumento As Long: colNumeroDocumento = 0
    Dim colNombre As Long: colNombre = 0
    Dim colApellido As Long: colApellido = 0
    Dim colFechaTransplante As Long: colFechaTransplante = 0
    
    If colMap.Exists("TipoDocumento") Then colTipoDocumento = colMap("TipoDocumento")
    If colMap.Exists("NumeroDocumento") Then colNumeroDocumento = colMap("NumeroDocumento")
    If colMap.Exists("Nombre") Then colNombre = colMap("Nombre")
    If colMap.Exists("Apellido") Then colApellido = colMap("Apellido")
    If colMap.Exists("FechaTransplante") Then colFechaTransplante = colMap("FechaTransplante")
    
    Dim nextRow As Long: nextRow = 2
    Dim rw As Range, ev As Long
    For Each rw In lo.DataBodyRange.Rows
        Dim td As Variant, nd As Variant, nombreVal As Variant, apeVal As Variant, fTrans As Variant
        If colTipoDocumento > 0 Then td = rw.Cells(colTipoDocumento).Value Else td = ""
        If colNumeroDocumento > 0 Then nd = rw.Cells(colNumeroDocumento).Value Else nd = ""
        If colNombre > 0 Then nombreVal = rw.Cells(colNombre).Value Else nombreVal = ""
        If colApellido > 0 Then apeVal = rw.Cells(colApellido).Value Else apeVal = ""
        If colFechaTransplante > 0 Then fTrans = rw.Cells(colFechaTransplante).Value Else fTrans = ""
        
        For ev = 1 To MAX_EVENTS
            Dim tipoName As String, fechaName As String, codigoName As String, faseName As String
            tipoName = "Tipo_Evento" & ev
            fechaName = "Fecha_Evento" & ev
            codigoName = "Codigo_Evento" & ev
            faseName = "Fase_Evento" & ev
            
            Dim tCol As Long, fCol As Long, codCol As Long, faseCol As Long
            tCol = 0: fCol = 0: codCol = 0: faseCol = 0
            If colMap.Exists(tipoName) Then tCol = colMap(tipoName)
            If colMap.Exists(fechaName) Then fCol = colMap(fechaName)
            If colMap.Exists(codigoName) Then codCol = colMap(codigoName)
            If colMap.Exists(faseName) Then faseCol = colMap(faseName)
            
            If fCol > 0 Then
                Dim fVal As Variant: fVal = rw.Cells(fCol).Value
                If Not IsEmpty(fVal) And IsDate(fVal) Then
                    wsE.Cells(nextRow, 1).Value = td
                    wsE.Cells(nextRow, 2).Value = nd
                    wsE.Cells(nextRow, 3).Value = nombreVal
                    wsE.Cells(nextRow, 4).Value = apeVal
                    wsE.Cells(nextRow, 5).Value = fTrans
                    wsE.Cells(nextRow, 6).Value = IIf(tCol > 0, rw.Cells(tCol).Value, "")
                    wsE.Cells(nextRow, 7).Value = fVal
                    wsE.Cells(nextRow, 8).Value = IIf(codCol > 0, rw.Cells(codCol).Value, "")
                    wsE.Cells(nextRow, 9).Value = IIf(faseCol > 0, rw.Cells(faseCol).Value, "")
                    nextRow = nextRow + 1
                End If
            End If
        Next ev
    Next rw
    
    '-------------------------Rellenar columnas auxiliares-------------------------------
    Dim lastCol As Long: lastCol = wsE.Cells(1, wsE.Columns.Count).End(xlToLeft).Column
    Dim lastRow As Long: lastRow = wsE.Cells(wsE.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        Dim d As Variant: d = wsE.Cells(r, 7).Value
        If IsDate(d) Then
            wsE.Cells(r, 10).Value = Year(d)                                                     'Año
            wsE.Cells(r, 11).Value = Month(d)                                                    'NumeroMes
            wsE.Cells(r, 12).Value = Format(d, "mmmm")                                           'NombreMes
            wsE.Cells(r, 13).Value = Int((Month(d) - 1) / 3) + 1                                 'NumeroTrimestre
            wsE.Cells(r, 14).Value = "T" & wsE.Cells(r, 13).Value & " " & wsE.Cells(r, 10).Value 'NombreTrimestre
            wsE.Cells(r, 15).Value = Format(d, "yyyy-mm")                                        'AnoMes
        End If
    Next r
    ' -------------Convertir a tabla----------------------
    Dim loE As ListObject
    On Error Resume Next
    Set loE = wsE.ListObjects("Eventos_Detallados")
    On Error GoTo 0
    If loE Is Nothing Then
        Set loE = wsE.ListObjects.Add(xlSrcRange, wsE.Range(wsE.Cells(1, 1), wsE.Cells(lastRow, lastCol)), xlYes)
        loE.Name = "Eventos_Detallados"
    Else
        loE.Resize wsE.Range("A1").CurrentRegion
    End If
    
    Dim j As Long
    For j = 1 To UBound(headers) + 1
        loE.ListColumns(j).Name = headers(j - 1)
    Next j
    
    MsgBox "Eventos_Detallados creado/actualizado. Total eventos: " & (nextRow - 2), vbInformation
            
End Sub

'-------------------- 2) GENERAR informe mensual del mes anterior --------------------

Public Sub CreateMonthlyReport_PreviousMonth()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    BuildEventDetail
    Dim prevDate As Date: prevDate = DateAdd("m", -1, Date)
    Dim Y As Long: Y = Year(prevDate)
    Dim M As Long: M = Month(prevDate)
    CreateMonthlyPivot Y, M
    MsgBox "Informe mensual generado para " & Format(prevDate, "mmmm yyyy"), vbInformation
ExitHandler:
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume ExitHandler
End Sub

Private Sub CreateMonthlyPivot(ByVal Y As Long, ByVal M As Long)
    Dim wsEvents As Worksheet: Set wsEvents = ThisWorkbook.Worksheets(EVENTS_SHEET)
    Dim loE As ListObject: Set loE = wsEvents.ListObjects("Eventos_Detallados")
    Dim sheetName As String: sheetName = "Informe_" & Format(DateSerial(Y, M, 1), "yyyy-mm")
    ' Se elimina el informe antiguo si existe
    On Error Resume Next
    If WorksheetExists(sheetName) Then Application.DisplayAlerts = False: ThisWorkbook.Worksheets(sheetName).Delete: Application.DisplayAlerts = True
    On Error GoTo 0
    Dim wsR As Worksheet: Set wsR = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsR.Name = sheetName
    wsR.Range("A1").Value = "Informe Mensual - " & Format(DateSerial(Y, M, 1), "mmmm yyyy")
    
    Dim pc As PivotCache
    Set pc = ThisWorkbook.PivotCaches.Create(xlDatabase, loE.Range.Address(True, True, xlR1C1, True))
    Dim pt As PivotTable
    Set pt = pc.CreatePivotTable(TableDestination:=wsR.Range("A3"), TableName:="PT_Monthly_" & Format(DateSerial(Y, M, 1), "yyyymm"))
    
    With pt
        .ClearAllFilters
        ' Poner Año y Numero de mes como filtros para seleccionar el mes
        .PivotFields("Ano").Orientation = xlPageField
        .PivotFields("NumeroMes").Orientation = xlPageField
        ' Filas por tipo de evento
        .PivotFields("Tipo_Evento").Orientation = xlRowField
        ' Columnas por Fase
        On Error Resume Next
        .PivotFields("Fase_Evento").Orientation = xlColumnField
        On Error GoTo 0
        ' Valor: conteo de eventos por fecha
        .AddDataField .PivotFields("Fecha_Evento"), "Total de Eventos", xlCount
        .PivotFields("Tipo_Evento").Caption = "Tipo de Evento"
        .PivotFields("Fase_Evento").Caption = "Fase del Evento"
        .PivotFields("Ano").Caption = "Año"
        .PivotFields("NumeroMes").Caption = "Número del Mes"
    End With
    
    'Aplicar los filtros
    On Error Resume Next
    pt.PivotFields("Ano").CurrentPage = Y
    pt.PivotFields("NumeroMes").CurrentPage = M
    On Error GoTo 0
    pt.RefreshTable
End Sub

'----------- 3) Generar informe trimestral ------------------------
Public Sub CreateQuarterlyReport_PreviousQuarter()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    BuildEventDetail
    Dim curQ As Long: curQ = Int((Month(Date) - 1) / 3) + 1
    Dim prevQ As Long, Y As Long
    If curQ > 1 Then
        prevQ = curQ - 1
        Y = Year(Date)
    Else
        prevQ = 4
        Y = Year(Date) - 1
    End If
    CreateQuarterlyPivot Y, prevQ
    MsgBox "Informe trimestral generado para Q" & prevQ & " " & Y, vbInformation
ExitHandler:
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume ExitHandler
End Sub

Private Sub CreateQuarterlyPivot(ByVal Y As Long, ByVal QuarterNum As Long)
    Dim wsEvents As Worksheet: Set wsEvents = ThisWorkbook.Worksheets(EVENTS_SHEET)
    Dim loE As ListObject: Set loE = wsEvents.ListObjects("Eventos_Detallados")
    Dim sheetName As String: sheetName = "Informe_T" & QuarterNum & "_" & Y
    On Error Resume Next
    If WorksheetExists(sheetName) Then Application.DisplayAlerts = False: ThisWorkbook.Worksheets(sheetName).Delete: Application.DisplayAlerts = True
    On Error GoTo 0
    Dim wsR As Worksheet: Set wsR = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsR.Name = sheetName
    wsR.Range("A1").Value = "Informe Trimestral - Q" & QuarterNum & " " & Y
    
    Dim pc As PivotCache
    Set pc = ThisWorkbook.PivotCaches.Create(xlDatabase, loE.Range.Address(True, True, xlR1C1, True))
    Dim pt As PivotTable
    Set pt = pc.CreatePivotTable(TableDestination:=wsR.Range("A3"), TableName:="PT_Quarterly_" & QuarterNum & "_" & Y)
    
    With pt
        .ClearAllFilters
        .PivotFields("Ano").Orientation = xlPageField
        .PivotFields("NumeroTrimestre").Orientation = xlPageField
        .PivotFields("Tipo_Evento").Orientation = xlRowField
        ' Columnas por Fase
        On Error Resume Next
        .PivotFields("Fase_Evento").Orientation = xlColumnField
        On Error GoTo 0
        ' Valor: conteo de eventos por fecha
        .AddDataField .PivotFields("Fecha_Evento"), "Total de Eventos", xlCount
        .PivotFields("Tipo_Evento").Caption = "Tipo de Evento"
        .PivotFields("Fase_Evento").Caption = "Fase del Evento"
        .PivotFields("Ano").Caption = "Año"
        .PivotFields("NumeroTrimestre").Caption = "Número del Trimestre"
    End With
    
    On Error Resume Next
    pt.PivotFields("Ano").CurrentPage = Y
    pt.PivotFields("NumeroTrimestre").CurrentPage = QuarterNum
    On Error GoTo 0
    pt.RefreshTable
End Sub

' --------------------------------4) Generar Informe Anual del año anterior --------------------------------
Public Sub CreateAnnualReport_PreviousYear()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    BuildEventDetail
    Dim Y As Long: Y = Year(Date) - 1
    CreateAnnualPivot Y
    MsgBox "Informe anual generado para " & Y, vbInformation
ExitHandler:
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume ExitHandler
End Sub

Private Sub CreateAnnualPivot(ByVal Y As Long)
    Dim wsEvents As Worksheet: Set wsEvents = ThisWorkbook.Worksheets(EVENTS_SHEET)
    Dim loE As ListObject: Set loE = wsEvents.ListObjects("Eventos_Detallados")
    Dim sheetName As String: sheetName = "Informe_" & Y
    On Error Resume Next
    If WorksheetExists(sheetName) Then Application.DisplayAlerts = False: ThisWorkbook.Worksheets(sheetName).Delete: Application.DisplayAlerts = True
    On Error GoTo 0
    Dim wsR As Worksheet: Set wsR = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsR.Name = sheetName
    wsR.Range("A1").Value = "Informe Anual - " & Y
    
    Dim pc As PivotCache
    Set pc = ThisWorkbook.PivotCaches.Create(xlDatabase, loE.Range.Address(True, True, xlR1C1, True))
    Dim pt As PivotTable
    Set pt = pc.CreatePivotTable(TableDestination:=wsR.Range("A3"), TableName:="PT_Annual_" & Y)
    
    With pt
        .ClearAllFilters
        .PivotFields("Ano").Orientation = xlPageField
        .PivotFields("Tipo_Evento").Orientation = xlRowField
        ' Columnas por Fase
        On Error Resume Next
        .PivotFields("Fase_Evento").Orientation = xlColumnField
        On Error GoTo 0
        ' Valor: conteo de eventos por fecha
        .AddDataField .PivotFields("Fecha_Evento"), "Total de Eventos", xlCount
        .PivotFields("Tipo_Evento").Caption = "Tipo de Evento"
        .PivotFields("Fase_Evento").Caption = "Fase del Evento"
        .PivotFields("Ano").Caption = "Año"
    End With
    
    On Error Resume Next
    pt.PivotFields("Ano").CurrentPage = Y
    On Error GoTo 0
    pt.RefreshTable
End Sub

'-------------------- 5) Crear una hoja MENU con botones --------------------
Public Sub CreateMenuSheet()
    Dim ws As Worksheet
    If WorksheetExists("MENU") Then Set ws = ThisWorkbook.Worksheets("MENU") Else Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Sheets(1)): ws.Name = "MENU"
    ws.Cells.Clear
    ws.Range("A1").Value = "Panel de generación de informes"
    ws.Range("A2").Value = "Usa los botones para generar los informes (mensual, trimestral o anual)."
    
    'Borrar botones previos (por sí los hay)
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    
    'Crear botón mensual
    Dim b As Shape
    Set b = ws.Shapes.AddShape(msoShapeRoundedRectangle, 20, 70, 300, 40)
    b.TextFrame2.TextRange.Characters.Text = "Generar informe mensual (mes anterior)"
    b.OnAction = "CreateMonthlyReport_PreviousMonth"
    
    'Crear botón trimestral
    Set b = ws.Shapes.AddShape(msoShapeRoundedRectangle, 20, 130, 300, 40)
    b.TextFrame2.TextRange.Characters.Text = "Generar informe trimestral (Trimestre anterior)"
    b.OnAction = "CreateQuarterlyReport_PreviousQuarter"
    
    'Crear botón anual
    Set b = ws.Shapes.AddShape(msoShapeRoundedRectangle, 20, 190, 300, 40)
    b.TextFrame2.TextRange.Characters.Text = "Generar informe anual (año anterior)"
    b.OnAction = "CreateAnnualReport_PreviousYear"
    
    'Crear botón de todo
    Set b = ws.Shapes.AddShape(msoShapeRoundedRectangle, 20, 250, 300, 40)
    b.TextFrame2.TextRange.Characters.Text = "Generar todos los informes (mensual + trimestral + anual)"
    b.OnAction = "GenerateAllReports"
    
    MsgBox "Hoja MENU creada/Actualizada. Usa los botones para generar informes."
End Sub

Public Sub GenerateAllReports()
    CreateMonthlyReport_PreviousMonth
    CreateQuarterlyReport_PreviousQuarter
    CreateAnnualReport_PreviousYear
End Sub

