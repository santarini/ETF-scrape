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
Dim str As String

Sheets.Add.Name = "MonthlyCorr"
Set topCell = Range("A1")
countx = 1
county = 1

For Each Basesheet In Worksheets
    If InStr(1, Basesheet.Name, "(Mon)") > 0 Then
        
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

        'get average and stddev of selection
        
            BaseArithmeticMean = Application.WorksheetFunction.Average(baseData)
            BaseStandardDeviation = Application.WorksheetFunction.StDev_P(baseData)
        
        'paste stats into monthsummary
            
            Worksheets("MonthlyCorr").Select
            str = Split(Basesheet.Name, "(")(0) & vbNewLine & Chr(181) & "=" & Format(BaseArithmeticMean, "Percent") & " " & ChrW(&H3C3) & "=" & Format(BaseStandardDeviation, "Percent")
            topCell.Offset(0, countx).WrapText = True
            topCell.Offset(0, countx).Value = str
            

                
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
                                
                            'get average and stddev of selection
        
                                TrgtArithmeticMean = Application.WorksheetFunction.Average(corrData)
                                TrgtStandardDeviation = Application.WorksheetFunction.StDev_P(corrData)
                                
                            'find correlation
                                CorrelationVar = Application.WorksheetFunction.Correl(baseData, corrData)
                                
                            'navigate to "MonthlyCorr"
                                Worksheets("MonthlyCorr").Select
                                
                            'paste corrData name in row
                                str = Split(CurrentSheet.Name, "(")(0) & vbNewLine & Chr(181) & "=" & Format(TrgtArithmeticMean, "Percent") & " " & ChrW(&H3C3) & "=" & Format(TrgtStandardDeviation, "Percent")
                                topCell.Offset(county, 0).Value = str
                                
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

'center first column and row
Rows("1:1").Select
    With Selection
        .HorizontalAlignment = xlCenter
    End With
Columns("A:A").Select
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    
'format numbers and cells
Range("B2").Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.NumberFormat = "0.00"
With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With

'heat map
    Selection.FormatConditions.AddColorScale ColorScaleType:=3
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueNumber
    Selection.FormatConditions(1).ColorScaleCriteria(1).Value = -1
    With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = 7039480
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValueNumber
    Selection.FormatConditions(1).ColorScaleCriteria(2).Value = 0
    With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(3).Type = _
        xlConditionValueNumber
    Selection.FormatConditions(1).ColorScaleCriteria(3).Value = 1
    With Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor
        .Color = 8109667
        .TintAndShade = 0
    End With

'autofit rows and cols
Cells.Select
Cells.EntireColumn.AutoFit
Cells.EntireRow.AutoFit

'reset selectio
Range("A1").Select

End Function
Function createPortfolio()

Dim MainRng As Range
Dim AssetA, AssetB As Range
Dim AssetAName, AssetBName As String
Dim AssetAReturn, AssetBReturn, AssetAStD, AssetBStD As Double
Dim openPos, closePos, DataRowCount As Integer


Set MainRng = Selection

'get asset a
Selection.End(xlUp).Select
Set AssetA = Selection


'extract asset a data
AssetAName = Split(AssetA.Value, Chr(181))(0)

openPos = InStr(AssetA, Chr(181))
closePos = InStr(AssetA, "%")
AssetAReturn = Mid(AssetA, openPos + 2, closePos - openPos - 2)

openPos = InStr(1, AssetA, "=")
openPos = InStr(openPos + 1, AssetA, "=")
closePos = InStr(1, AssetA, "%")
closePos = InStr(closePos + 1, AssetA, "%")
AssetAStD = Mid(AssetA, openPos + 1, closePos - openPos - 2)


MainRng.Select

'get asset b
Selection.End(xlToLeft).Select
Set AssetB = Selection

'extract asset b data
AssetBName = Split(AssetB.Value, Chr(181))(0)

openPos = InStr(AssetB, Chr(181))
closePos = InStr(AssetB, "%")
AssetBReturn = Mid(AssetB, openPos + 2, closePos - openPos - 2)

openPos = InStr(1, AssetB, "=")
openPos = InStr(openPos + 1, AssetB, "=")
closePos = InStr(1, AssetB, "%")
closePos = InStr(closePos + 1, AssetB, "%")
AssetBStD = Mid(AssetB, openPos + 1, closePos - openPos - 2)

Sheets.Add.Name = "Portfolio"

Range("A1").Value = "Asset A Weight"
Range("B1").Value = "Asset B Weight"

j = 1
k = 0
For i = 1 To 11
    Range("A1").Offset(i, 0).Value = k
    Range("B1").Offset(i, 0).Value = j
    j = j - 0.1
    k = k + 0.1
Next

Range("A2:B2").Select
Range(Selection, Selection.End(xlDown)).Select

Selection.NumberFormat = "0%"

Range("C1").Value = "Portfolio Return"
Range("D1").Value = "Portfolio StDev"

For i = 1 To 11
    AWeight = Range("A1").Offset(i, 0).Value
    BWeight = Range("B1").Offset(i, 0).Value
    PortfolioReturn = ((AWeight * AssetAReturn) + (BWeight * AssetBReturn)) / 100
    PortfolioStdDev = Sqr(((AWeight * AssetAStD) ^ 2) + ((BWeight * AssetBStD) ^ 2) + (2 * AWeight * BWeight * MainRng.Value * AssetAStD * AssetBStD)) / 100
    Range("C1").Offset(i, 0).Value = PortfolioReturn
    Range("D1").Offset(i, 0).Value = PortfolioStdDev
Next

Range("C2:D2").Select
Range(Selection, Selection.End(xlDown)).Select

Selection.NumberFormat = "0%"

DataRowCount = Selection.Rows.count

    ActiveSheet.Shapes.AddChart2(240, xlXYScatterSmooth).Select
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(1).XValues = "=Portfolio!$D$2:$D$" & DataRowCount
    ActiveChart.FullSeriesCollection(1).Values = "=Portfolio!$C$2:$C$" & DataRowCount
    ActiveChart.HasTitle = True
    ActiveChart.ChartTitle.Text = "Efficient Frontier"
    With ActiveChart.Axes(xlValue)
     .HasTitle = True
     With .AxisTitle
     .Caption = "Portfolio Return"
     End With
    End With
    With ActiveChart.Axes(xlCategory)
     .HasTitle = True
     .AxisTitle.Caption = "Portfolio StDev"
    End With
    
End Function

