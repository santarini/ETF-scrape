Sub compileByPath()

Dim FolderPath As String
Dim PathCountCondition As String
Dim FileName As String
Dim Count As Integer
Dim FileNumber As Integer
Dim MainWB As Workbook
Dim WB As Workbook
Dim Rng As Range
Dim RngNoPath As String
Dim StartTime As Double
Dim SecondsElapsed As Double
Dim tickersPerSec As Double
Dim SummaryRng As Range

StartTime = Timer

'set this workbook as the main workbook

Set MainWB = ActiveWorkbook
MainWB.Sheets.Add.Name = "PathSet"
Set Rng = Range("A1")

Application.DisplayAlerts = False

'define folder path
FolderPath = "C:\Users\CommandCenter\Desktop\ETF-scrape-master\stock_dfs"

'count number of CSVs in folder

PathCountCondition = FolderPath & "\*.csv"

FileName = Dir(PathCountCondition)

Do While FileName <> ""
    Rng.Value = FileName
    Rng.Offset(1, 0).Select
    Count = Count + 1
    Set Rng = ActiveCell
    FileName = Dir()
Loop

Worksheets("PathSet").Activate
Set Rng = Range("A1")
Rng.Select
Range(Selection, Selection.End(xlDown)).Select
Count = Selection.Rows.Count

Worksheets("PathSet").Activate
Rng.Select

For FileNumber = 1 To Count 'you can change count to a constant for sample runs
    
    'open the file
    
    FileName = FolderPath & "\" & Rng
    
    Set WB = Workbooks.Open(FileName)
    
    'copy its contents
    
    WB.Activate
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

    'create new sheet, and paste it into the main workbook
    
    MainWB.Activate
    RngNoPath = Left(Rng, Len(Rng) - 4)
    MainWB.Sheets.Add.Name = RngNoPath & "(D)"
    Range("A1").Select
    ActiveSheet.Paste
    Selection.Columns.AutoFit
    Range("A1").Select
    
    'close file
    WB.Close
    
    Call orderDataForGraphingIEX
    Call manipulateDataIEX
    Call monthlyData(RngNoPath)
    
    'Worksheets("PathSet").Activate
    
    Worksheets("PathSet").Activate
    Rng.Offset(1, 0).Select
    Set Rng = ActiveCell
    
Next FileNumber

Worksheets("PathSet").Delete
                                        
'tell me how long it took
SecondsElapsed = Round(Timer - StartTime, 2)
tickersPerSec = Round(SecondsElapsed / Count, 2)
MsgBox "This code ran successfully in " & SecondsElapsed & " seconds" & vbCrLf & "Approximately " & tickersPerSec & " per second", vbInformation
                                        
End Sub
Function orderDataForGraphingIEX()
    
    'order the columns for graphing
                                                            
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("G:G").Select
    Selection.Cut
    Columns("B:B").Select
    ActiveSheet.Paste
    
    'add commas and dollar signs
    
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    Range("C2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Currency"
    
End Function
Function manipulateDataIEX()
    Dim Rng As Range
    Dim LastRow As Integer
    
    'find the last row
    
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    LastRow = Selection.Rows.Count
    
    'day average average
    Range("G1").Value = "Day Average"
    Range("G2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=AVERAGE(RC[-4]:RC[-1])"
    Range("G2").Select
    Selection.AutoFill Destination:=Range("G2:G" & LastRow)
    
    'data manipulations
    Range("H1").Value = "Intraday Open to Close"
    Application.CutCopyMode = False
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]/RC[-5]-1"
    Range("I1").Value = "Intraday %"
    Range("I2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]/RC[-6]"
    
    Range("H2:I2").Select
    Selection.AutoFill Destination:=Range("H2:I" & LastRow)
    Range("H:H").Select
    Selection.Style = "Currency"
    Range("I:I").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.000%"
    
    
    'resize one last time
    Range("A:I").Select
    Selection.Columns.AutoFit
    
    'set active cell to home
    Range("A1").Select
   
End Function

Function monthlyData(RngNoPath As String)

Dim PTTop, PTYrs, PTMons, cell, CellWorking As Range
Dim PTRowCount As Integer
Dim WBTB As Worksheet
Dim str As String
Dim useableData As Range


'Insert pivot table
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Set useableData = Selection
    str = RngNoPath & "(Mon)!R3C1"
    
    Sheets.Add.Name = RngNoPath & "(Mon)"
    
    
    Sheets(RngNoPath & "(Mon)").Activate
        ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        useableData, Version:=6).CreatePivotTable TableDestination:= _
        Sheets(RngNoPath & "(Mon)").Range("A3"), TableName:="PivotTable1", DefaultVersion:=6
    Sheets(RngNoPath & "(Mon)").Select
    Cells(3, 1).Select
    
'Configure pivot table
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("date")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields("date").AutoGroup
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Intraday %"), "Sum of Intraday %", xlSum
    Cells(4, 1).Select
    Selection.Group Start:=True, End:=True, Periods:=Array(False, False, False, _
        False, True, False, True)
        
'Find "Row Labels"
    Cells.Find(What:="Row Labels", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate

'set cell to Rng

    Set PTTop = Selection
    PTTop.Value = "Month"
    

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

'delete the top rows
    Rows("1:2").Delete

'format the data
    
    Columns("C:C").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.000%"
    

End Function
