Sub PivotBuilder()
    Dim sRow1 As String, sRow2 As String, sCol As String, sData As String, sFilter As String
    Dim oTable As PivotTable, oField As PivotField, oWS As Worksheet
    
    Set oWS = ActiveSheet
    If Range("D1") = "" Then MsgBox "You are on the wrong sheet ": Exit Sub
    sRow1 = Application.InputBox("Click on label for rows", , , , , , , 2)
    sRow2 = Application.InputBox("Click on label for 2nd row (or Cancel)", , , , , , , 2)
    sCol = Application.InputBox("Click on label for cols (or Cancel)", , , , , , , 2)
    sData = Application.InputBox("Click on label for totals", , , , , , , 2)
    sFilter = Application.InputBox("Click on label for filter (or Cancel)", , , , , , , 2)
    
    Set oTable = oWS.PivotTableWizard
    Set oField = oTable.PivotFields(sRow1): oField.Orientation = xlRowField
    
    If sRow2 <> "False" Then Set oField = oTable.PivotFields(sRow2): oField.Orientation = xlRowField
    
    If sCol <> "False" Then Set oField = oTable.PivotFields(sCol): oField.Orientation = xlColumnField
    
    Set oField = oTable.PivotFields(sData): oField.Orientation = xlDataField
    oField.NumberFormat = oWS.Cells(2, WorksheetFunction.Match(sData, oWS.Rows(1), 0)).NumberFormat
    
    If sFilter <> "False" Then Set oField = oTable.PivotFields(sFilter): oField.Orientation = xlPageField
    Range("A1").CurrentRegion.EntireColumn.AutoFit
    
    
End Sub
