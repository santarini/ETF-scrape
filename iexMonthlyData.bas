Sub Macro4()

Dim PTTop, PTYrs, PTMons, cell, CellWorking As Range
Dim PTRowCount As Integer



'Insert pivot table
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Sheets.Add.Name = "BALLS"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "ADBE!R1C1:R277C9", Version:=6).CreatePivotTable TableDestination:= _
        "BALLS!R3C1", TableName:="PivotTable6", DefaultVersion:=6
    Sheets("BALLS").Select
    Cells(3, 1).Select
    
'Configure pivot table
    With ActiveSheet.PivotTables("PivotTable6").PivotFields("date")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable6").PivotFields("date").AutoGroup
    ActiveSheet.PivotTables("PivotTable6").AddDataField ActiveSheet.PivotTables( _
        "PivotTable6").PivotFields("Intraday %"), "Sum of Intraday %", xlSum
    Cells(4, 1).Select
    Selection.Group Start:=True, End:=True, Periods:=Array(False, False, False, _
        False, True, False, True)
        
'Find "Row Labels"
    Cells.Find(What:="Row Labels", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate

'set cell to Rng

    Set PTTop = Selection

'select content to bottom
    Range(Selection, Selection.End(xlDown)).Select

'RowCount
    Set PTMons = Selection

'select content to right
    Range(Selection, Selection.End(xlToRight)).Select

'copy selection

    Selection.Copy

'paste values

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'insert a column to the left push content right

    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

'Select PTMons
    PTMons.Activate
    PTMons.Offset(0, -1).Select
    Set PTYrs = Selection
    
'Add Label to top of year column
    PTTop.Offset(0, -1).Value = "Year"

'Rng activate
    PTTop.Select

For Each cell In PTMons
    If IsNumeric(cell) Then
        cell.Select
        Selection.Cut
        cell.Offset(0, -1).Select
        ActiveSheet.Paste
    End If
Next


For Each cell In PTYrs
    If IsEmpty(cell) Then
        cell.Value = cell.Offset(-1, 0).Value
    End If
Next


For Each cell In PTMons
    If IsEmpty(cell) Or InStr(cell, "Grand Total") > 0 Then
        cell.Select
        cell.EntireRow.Delete
    End If
Next

End Sub
