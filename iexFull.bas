Sub compileByPath()

Dim FolderPath As String
Dim PathCountCondition As String
Dim FileName As String
Dim count As Integer
Dim FileNumber As Integer
Dim MainWB As Workbook
Dim WB As Workbook
Dim Rng As Range
Dim RngNoPath As String
Dim StartTime As Double
Dim SecondsElapsed As Double
Dim tickersPerSec As Double
Dim SummaryRng As Range
Dim CurrentSheet As Worksheet
Dim SheetName As String


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
    count = count + 1
    Set Rng = ActiveCell
    FileName = Dir()
Loop

Worksheets("PathSet").Activate
Set Rng = Range("A1")
Rng.Select
Range(Selection, Selection.End(xlDown)).Select
count = Selection.Rows.count

Worksheets("PathSet").Activate
Rng.Select

For FileNumber = 1 To count 'you can change count to a constant for sample runs
    
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
MainWB.Sheets.Add.Name = "MonthSummary"
MainWB.Sheets.Add.Name = "DailySummary"

Worksheets("MonthSummary").Activate
Range("A1").Value = "Ticker"
Range("A2").Value = "Average"
Range("A3").Value = "Variance"
Range("A4").Value = "StrdDev"

For Each CurrentSheet In Worksheets
    If InStr(1, CurrentSheet.Name, "(Mon)") > 0 Then
        CurrentSheet.Activate
        SheetName = Split(CurrentSheet.Name, "(")(0)
        Call monthlySummary(SheetName)
        CurrentSheet.Activate
    End If
    'If InStr(1, CurrentSheet.Name, "(D)") > 0 Then
    '    CurrentSheet.Activate
    '    SheetName = CurrentSheet.Name
    '    Call dailySummary(SheetName)
    '    CurrentSheet.Activate
    'End If
Next
                                        
'tell me how long it took
SecondsElapsed = Round(Timer - StartTime, 2)
tickersPerSec = Round(SecondsElapsed / count, 2)
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
    LastRow = Selection.Rows.count
    
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
    ActiveCell.FormulaR1C1 = "=RC[-2]-RC[-5]"
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
    Range("A1").Select
    
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
    
're-center A1
    Range("A1").Select
    

End Function
Function dailySummary(SheetName As String)

'find intraday %
'move one cell down
'select to bottom of content
'selection equals usableData
'get summary stats on usable data
'worksheet dailySummary select
    
End Function
Function monthlySummary(SheetName As String)

    Dim useableData As Range
    Dim MonthlyArithmeticMean, MonthlyStandardDeviation, MonthlyVariance As Double
    
    
'find "Sum of Intraday"

    Cells.Find(What:="Sum of Intraday", After:=ActiveCell, LookIn:=xlFormulas _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate

'select data beneath
    
    Selection.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    
'selection equals usableData
    Set useableData = Selection

    MonthlyArithmeticMean = Application.WorksheetFunction.Average(useableData)
    MonthlyStandardDeviation = Application.WorksheetFunction.StDev_P(useableData)
    MonthlyVariance = MonthlyStandardDeviation * MonthlyStandardDeviation

'worksheet dailySummary select
    Worksheets("MonthSummary").Activate
    Range("A1").Select
    If IsEmpty(Range("A1").Offset(0, 1)) Then
        Range("B1").Select
    Else
        Selection.End(xlToRight).Select
        Selection.Offset(0, 1).Select
    End If
    
    ActiveCell.Value = SheetName
    
    ActiveCell.Offset(1, 0).Value = MonthlyArithmeticMean
    ActiveCell.Offset(2, 0).Value = MonthlyStandardDeviation
    ActiveCell.Offset(3, 0).Value = MonthlyVariance

End Function
Function MonthCorrelate()

Dim baseData, corrData, topCell As Range
Dim countx, county As Integer
Dim CorrelationVar As Double

Sheets.Add.Name = "MonthlyCorr"
Set topCell = Range("A1")
countx = 1
county = 1

For Each Basesheet In Worksheets
    If InStr(1, Basesheet.Name, "(Mon)") > 0 Then
        Worksheets("MonthlyCorr").Select
        topCell.Offset(0, countx).Value = Split(Basesheet.Name, "(")(0)
        
        Basesheet.Activate
        
        'find "Sum of Intraday"
        
            Cells.Find(What:="Sum of Intraday", After:=ActiveCell, LookIn:=xlFormulas _
                , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False).Activate
        
        'select data beneath
            
            Selection.Offset(1, 0).Select
            Range(Selection, Selection.End(xlDown)).Select
            
        'selection equals usableData
            Set baseData = Selection
                
                For Each CurrentSheet In Worksheets
                    If InStr(1, CurrentSheet.Name, "(Mon)") > 0 Then
                            CurrentSheet.Activate
    
                            'find "Sum of Intraday"
                            
                                Cells.Find(What:="Sum of Intraday", After:=ActiveCell, LookIn:=xlFormulas _
                                    , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                    MatchCase:=False, SearchFormat:=False).Activate
                            
                            'select data beneath
                                
                                Selection.Offset(1, 0).Select
                                Range(Selection, Selection.End(xlDown)).Select
                                
                            'selection equals usableData
                                Set corrData = Selection
                                
                            'find correlation
                                CorrelationVar = Application.WorksheetFunction.Correl(baseData, corrData)
                                
                            'navigate to "MonthlyCorr"
                                Worksheets("MonthlyCorr").Select
                                
                            'paste corrData name in row
                                topCell.Offset(county, 0).Value = Split(CurrentSheet.Name, "(")(0)
                                
                            'paste correlation
                                topCell.Offset(county, countx).Value = CorrelationVar
                                county = county + 1
                    
                    End If
                Next
            county = 1
            countx = countx + 1
        Basesheet.Activate
    End If
Next
Worksheets("MonthlyCorr").Activate


End Function
