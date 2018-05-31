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
FolderPath = "C:\Users\santa\Desktop\hochuli_sample\historical-price-statistics-clone\stock_dfs"

'count number of CSVs in folder

PathCountCondition = FolderPath & "\*.csv"

FileName = Dir(PathCountCondition)

Do While FileName <> ""
    Count = Count + 1
    FileName = Dir()
    Rng.Value = FileName
    Rng.Offset(1, 0).Select
    Set Rng = ActiveCell
Loop

Worksheets("PathSet").Activate
Set Rng = Range("A1")
Rng.Select
Range(Selection, Selection.End(xlDown)).Select
Count = Selection.Rows.Count

'create summary page
MainWB.Sheets.Add.Name = "Summary"
Call createSummaryPage
Set SummaryRng = Worksheets("Summary").Range("A4")

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
    MainWB.Sheets.Add.Name = RngNoPath
    Range("A1").Select
    ActiveSheet.Paste
    Selection.Columns.AutoFit
    Range("A1").Select
    
    'close file
    WB.Close
    
    Call orderDataForGraphing
    Call manipulateData
    
    'POPULATE SUMMARY PAGE HERE
    Call populateSummary(SummaryRng)
    Worksheets("Summary").Activate
    SummaryRng.Offset(1, 0).Select
    Set SummaryRng = Selection
    
    'Worksheets("PathSet").Activate
    
    Worksheets("PathSet").Activate
    Rng.Offset(1, 0).Select
    Set Rng = ActiveCell
    
Next FileNumber
                                        
'tell me how long it took
SecondsElapsed = Round(Timer - StartTime, 2)
tickersPerSec = Round(SecondsElapsed / Count, 2)
MsgBox "This code ran successfully in " & SecondsElapsed & " seconds" & vbCrLf & "Approximately " & tickersPerSec & " seconds per ticker", vbInformation
                                        
End Sub

Function createSummaryPage()
    Dim Rng As Range

    Set Rng = Range("A1")
'top headers
    Rng.Value = "Description"
    Rng.Offset(0, 1).Value = "Description"
    Rng.Offset(0, 2).Value = "Description"
    Rng.Offset(0, 3).Value = "Volume"
    Rng.Offset(0, 4).Value = "Volume"
    Rng.Offset(0, 5).Value = "Volume"
    Rng.Offset(0, 6).Value = "Volume"
    Rng.Offset(0, 7).Value = "Volume"
    Rng.Offset(0, 8).Value = "Volume"
    Rng.Offset(0, 9).Value = "Volume"
    Rng.Offset(0, 10).Value = "Volume"
    Rng.Offset(0, 11).Value = "Volume"
    Rng.Offset(0, 12).Value = "Volume"
    Rng.Offset(0, 13).Value = "Volume"
    Rng.Offset(0, 14).Value = "Volume"
    Rng.Offset(0, 15).Value = "Volume"
    Rng.Offset(0, 16).Value = "Volume"
    Rng.Offset(0, 17).Value = "Volume"
    Rng.Offset(0, 18).Value = "Volume"
    Rng.Offset(0, 19).Value = "Volume"
    Rng.Offset(0, 20).Value = "Previous Close to Close"
    Rng.Offset(0, 21).Value = "Previous Close to Close"
    Rng.Offset(0, 22).Value = "Previous Close to Close"
    Rng.Offset(0, 23).Value = "Previous Close to Close"
    Rng.Offset(0, 24).Value = "Previous Close to Close"
    Rng.Offset(0, 25).Value = "Previous Close to Close"
    Rng.Offset(0, 26).Value = "Previous Close to Close"
    Rng.Offset(0, 27).Value = "Previous Close to Close"
    Rng.Offset(0, 28).Value = "Previous Close to Close"
    Rng.Offset(0, 29).Value = "Previous Close to Close"
    Rng.Offset(0, 30).Value = "Previous Close to Close"
    Rng.Offset(0, 31).Value = "Previous Close to Close"
    Rng.Offset(0, 32).Value = "Previous Close to Close"
    Rng.Offset(0, 33).Value = "Previous Close to Close"
    Rng.Offset(0, 34).Value = "Previous Close to Close"
    Rng.Offset(0, 35).Value = "Previous Close to Close"
    Rng.Offset(0, 36).Value = "Previous Close to Close"
    Rng.Offset(0, 37).Value = "Previous Close to Close"
    Rng.Offset(0, 38).Value = "Previous Close to Close"
    Rng.Offset(0, 39).Value = "Previous Close to Close"
    Rng.Offset(0, 40).Value = "Previous Close to Close"
    Rng.Offset(0, 41).Value = "Previous Close to Close"
    Rng.Offset(0, 42).Value = "Previous Close to Close"
    Rng.Offset(0, 43).Value = "Previous Close to Close"
    Rng.Offset(0, 44).Value = "Previous Close to Close"
    Rng.Offset(0, 45).Value = "Previous Close to Close"
    Rng.Offset(0, 46).Value = "Previous Close to Close"
    Rng.Offset(0, 47).Value = "Previous Close to Close"
    Rng.Offset(0, 48).Value = "Previous Close to Close"
    Rng.Offset(0, 49).Value = "Previous Close to Close"
    Rng.Offset(0, 50).Value = "Previous Close to Close"
    Rng.Offset(0, 51).Value = "Previous Close to Close"
    Rng.Offset(0, 52).Value = "Previous Close to Close"
    Rng.Offset(0, 53).Value = "Previous Close to Close"
    Rng.Offset(0, 54).Value = "Previous Open to Open"
    Rng.Offset(0, 55).Value = "Previous Open to Open"
    Rng.Offset(0, 56).Value = "Previous Open to Open"
    Rng.Offset(0, 57).Value = "Previous Open to Open"
    Rng.Offset(0, 58).Value = "Previous Open to Open"
    Rng.Offset(0, 59).Value = "Previous Open to Open"
    Rng.Offset(0, 60).Value = "Previous Open to Open"
    Rng.Offset(0, 61).Value = "Previous Open to Open"
    Rng.Offset(0, 62).Value = "Previous Open to Open"
    Rng.Offset(0, 63).Value = "Previous Open to Open"
    Rng.Offset(0, 64).Value = "Previous Open to Open"
    Rng.Offset(0, 65).Value = "Previous Open to Open"
    Rng.Offset(0, 66).Value = "Previous Open to Open"
    Rng.Offset(0, 67).Value = "Previous Open to Open"
    Rng.Offset(0, 68).Value = "Previous Open to Open"
    Rng.Offset(0, 69).Value = "Previous Open to Open"
    Rng.Offset(0, 70).Value = "Previous Open to Open"
    Rng.Offset(0, 71).Value = "Previous Open to Open"
    Rng.Offset(0, 72).Value = "Previous Open to Open"
    Rng.Offset(0, 73).Value = "Previous Open to Open"
    Rng.Offset(0, 74).Value = "Previous Open to Open"
    Rng.Offset(0, 75).Value = "Previous Open to Open"
    Rng.Offset(0, 76).Value = "Previous Open to Open"
    Rng.Offset(0, 77).Value = "Previous Open to Open"
    Rng.Offset(0, 78).Value = "Previous Open to Open"
    Rng.Offset(0, 79).Value = "Previous Open to Open"
    Rng.Offset(0, 80).Value = "Previous Open to Open"
    Rng.Offset(0, 81).Value = "Previous Open to Open"
    Rng.Offset(0, 82).Value = "Previous Open to Open"
    Rng.Offset(0, 83).Value = "Previous Open to Open"
    Rng.Offset(0, 84).Value = "Previous Open to Open"
    Rng.Offset(0, 85).Value = "Previous Open to Open"
    Rng.Offset(0, 86).Value = "Previous Open to Open"
    Rng.Offset(0, 87).Value = "Previous Open to Open"
    Rng.Offset(0, 88).Value = "Previous Close to Open"
    Rng.Offset(0, 89).Value = "Previous Close to Open"
    Rng.Offset(0, 90).Value = "Previous Close to Open"
    Rng.Offset(0, 91).Value = "Previous Close to Open"
    Rng.Offset(0, 92).Value = "Previous Close to Open"
    Rng.Offset(0, 93).Value = "Previous Close to Open"
    Rng.Offset(0, 94).Value = "Previous Close to Open"
    Rng.Offset(0, 95).Value = "Previous Close to Open"
    Rng.Offset(0, 96).Value = "Previous Close to Open"
    Rng.Offset(0, 97).Value = "Previous Close to Open"
    Rng.Offset(0, 98).Value = "Previous Close to Open"
    Rng.Offset(0, 99).Value = "Previous Close to Open"
    Rng.Offset(0, 100).Value = "Previous Close to Open"
    Rng.Offset(0, 101).Value = "Previous Close to Open"
    Rng.Offset(0, 102).Value = "Previous Close to Open"
    Rng.Offset(0, 103).Value = "Previous Close to Open"
    Rng.Offset(0, 104).Value = "Previous Close to Open"
    Rng.Offset(0, 105).Value = "Previous Close to Open"
    Rng.Offset(0, 106).Value = "Previous Close to Open"
    Rng.Offset(0, 107).Value = "Previous Close to Open"
    Rng.Offset(0, 108).Value = "Previous Close to Open"
    Rng.Offset(0, 109).Value = "Previous Close to Open"
    Rng.Offset(0, 110).Value = "Previous Close to Open"
    Rng.Offset(0, 111).Value = "Previous Close to Open"
    Rng.Offset(0, 112).Value = "Previous Close to Open"
    Rng.Offset(0, 113).Value = "Previous Close to Open"
    Rng.Offset(0, 114).Value = "Previous Close to Open"
    Rng.Offset(0, 115).Value = "Previous Close to Open"
    Rng.Offset(0, 116).Value = "Previous Close to Open"
    Rng.Offset(0, 117).Value = "Previous Close to Open"
    Rng.Offset(0, 118).Value = "Previous Close to Open"
    Rng.Offset(0, 119).Value = "Previous Close to Open"
    Rng.Offset(0, 120).Value = "Previous Close to Open"
    Rng.Offset(0, 121).Value = "Previous Close to Open"
    Rng.Offset(0, 122).Value = "Intraday Open to Close"
    Rng.Offset(0, 123).Value = "Intraday Open to Close"
    Rng.Offset(0, 124).Value = "Intraday Open to Close"
    Rng.Offset(0, 125).Value = "Intraday Open to Close"
    Rng.Offset(0, 126).Value = "Intraday Open to Close"
    Rng.Offset(0, 127).Value = "Intraday Open to Close"
    Rng.Offset(0, 128).Value = "Intraday Open to Close"
    Rng.Offset(0, 129).Value = "Intraday Open to Close"
    Rng.Offset(0, 130).Value = "Intraday Open to Close"
    Rng.Offset(0, 131).Value = "Intraday Open to Close"
    Rng.Offset(0, 132).Value = "Intraday Open to Close"
    Rng.Offset(0, 133).Value = "Intraday Open to Close"
    Rng.Offset(0, 134).Value = "Intraday Open to Close"
    Rng.Offset(0, 135).Value = "Intraday Open to Close"
    Rng.Offset(0, 136).Value = "Intraday Open to Close"
    Rng.Offset(0, 137).Value = "Intraday Open to Close"
    Rng.Offset(0, 138).Value = "Intraday Open to Close"
    Rng.Offset(0, 139).Value = "Intraday Open to Close"
    Rng.Offset(0, 140).Value = "Intraday Open to Close"
    Rng.Offset(0, 141).Value = "Intraday Open to Close"
    Rng.Offset(0, 142).Value = "Intraday Open to Close"
    Rng.Offset(0, 143).Value = "Intraday Open to Close"
    Rng.Offset(0, 144).Value = "Intraday Open to Close"
    Rng.Offset(0, 145).Value = "Intraday Open to Close"
    Rng.Offset(0, 146).Value = "Intraday Open to Close"
    Rng.Offset(0, 147).Value = "Intraday Open to Close"
    Rng.Offset(0, 148).Value = "Intraday Open to Close"
    Rng.Offset(0, 149).Value = "Intraday Open to Close"
    Rng.Offset(0, 150).Value = "Intraday Open to Close"
    Rng.Offset(0, 151).Value = "Intraday Open to Close"
    Rng.Offset(0, 152).Value = "Intraday Open to Close"
    Rng.Offset(0, 153).Value = "Intraday Open to Close"
    Rng.Offset(0, 154).Value = "Intraday Open to Close"
    Rng.Offset(0, 155).Value = "Intraday Open to Close"


    'Sub headers
    Rng.Offset(1, 0).Value = "Actual"
    Rng.Offset(1, 1).Value = "Actual"
    Rng.Offset(1, 2).Value = "Actual"
    Rng.Offset(1, 3).Value = "Actual"
    Rng.Offset(1, 4).Value = "Actual"
    Rng.Offset(1, 5).Value = "Actual"
    Rng.Offset(1, 6).Value = "Actual"
    Rng.Offset(1, 7).Value = "Actual"
    Rng.Offset(1, 8).Value = "Actual"
    Rng.Offset(1, 9).Value = "Actual"
    Rng.Offset(1, 10).Value = "Actual"
    Rng.Offset(1, 11).Value = "Actual"
    Rng.Offset(1, 12).Value = "Actual"
    Rng.Offset(1, 13).Value = "Actual"
    Rng.Offset(1, 14).Value = "Actual"
    Rng.Offset(1, 15).Value = "Actual"
    Rng.Offset(1, 16).Value = "Actual"
    Rng.Offset(1, 17).Value = "Actual"
    Rng.Offset(1, 18).Value = "Actual"
    Rng.Offset(1, 19).Value = "Actual"
    Rng.Offset(1, 20).Value = "Actual"
    Rng.Offset(1, 21).Value = "Actual"
    Rng.Offset(1, 22).Value = "Actual"
    Rng.Offset(1, 23).Value = "Actual"
    Rng.Offset(1, 24).Value = "Actual"
    Rng.Offset(1, 25).Value = "Actual"
    Rng.Offset(1, 26).Value = "Actual"
    Rng.Offset(1, 27).Value = "Actual"
    Rng.Offset(1, 28).Value = "Actual"
    Rng.Offset(1, 29).Value = "Actual"
    Rng.Offset(1, 30).Value = "Actual"
    Rng.Offset(1, 31).Value = "Actual"
    Rng.Offset(1, 32).Value = "Actual"
    Rng.Offset(1, 33).Value = "Actual"
    Rng.Offset(1, 34).Value = "Actual"
    Rng.Offset(1, 35).Value = "Actual"
    Rng.Offset(1, 36).Value = "Actual"
    Rng.Offset(1, 37).Value = "Percent"
    Rng.Offset(1, 38).Value = "Percent"
    Rng.Offset(1, 39).Value = "Percent"
    Rng.Offset(1, 40).Value = "Percent"
    Rng.Offset(1, 41).Value = "Percent"
    Rng.Offset(1, 42).Value = "Percent"
    Rng.Offset(1, 43).Value = "Percent"
    Rng.Offset(1, 44).Value = "Percent"
    Rng.Offset(1, 45).Value = "Percent"
    Rng.Offset(1, 46).Value = "Percent"
    Rng.Offset(1, 47).Value = "Percent"
    Rng.Offset(1, 48).Value = "Percent"
    Rng.Offset(1, 49).Value = "Percent"
    Rng.Offset(1, 50).Value = "Percent"
    Rng.Offset(1, 51).Value = "Percent"
    Rng.Offset(1, 52).Value = "Percent"
    Rng.Offset(1, 53).Value = "Percent"
    Rng.Offset(1, 54).Value = "Actual"
    Rng.Offset(1, 55).Value = "Actual"
    Rng.Offset(1, 56).Value = "Actual"
    Rng.Offset(1, 57).Value = "Actual"
    Rng.Offset(1, 58).Value = "Actual"
    Rng.Offset(1, 59).Value = "Actual"
    Rng.Offset(1, 60).Value = "Actual"
    Rng.Offset(1, 61).Value = "Actual"
    Rng.Offset(1, 62).Value = "Actual"
    Rng.Offset(1, 63).Value = "Actual"
    Rng.Offset(1, 64).Value = "Actual"
    Rng.Offset(1, 65).Value = "Actual"
    Rng.Offset(1, 66).Value = "Actual"
    Rng.Offset(1, 67).Value = "Actual"
    Rng.Offset(1, 68).Value = "Actual"
    Rng.Offset(1, 69).Value = "Actual"
    Rng.Offset(1, 70).Value = "Actual"
    Rng.Offset(1, 71).Value = "Percent"
    Rng.Offset(1, 72).Value = "Percent"
    Rng.Offset(1, 73).Value = "Percent"
    Rng.Offset(1, 74).Value = "Percent"
    Rng.Offset(1, 75).Value = "Percent"
    Rng.Offset(1, 76).Value = "Percent"
    Rng.Offset(1, 77).Value = "Percent"
    Rng.Offset(1, 78).Value = "Percent"
    Rng.Offset(1, 79).Value = "Percent"
    Rng.Offset(1, 80).Value = "Percent"
    Rng.Offset(1, 81).Value = "Percent"
    Rng.Offset(1, 82).Value = "Percent"
    Rng.Offset(1, 83).Value = "Percent"
    Rng.Offset(1, 84).Value = "Percent"
    Rng.Offset(1, 85).Value = "Percent"
    Rng.Offset(1, 86).Value = "Percent"
    Rng.Offset(1, 87).Value = "Percent"
    Rng.Offset(1, 88).Value = "Actual"
    Rng.Offset(1, 89).Value = "Actual"
    Rng.Offset(1, 90).Value = "Actual"
    Rng.Offset(1, 91).Value = "Actual"
    Rng.Offset(1, 92).Value = "Actual"
    Rng.Offset(1, 93).Value = "Actual"
    Rng.Offset(1, 94).Value = "Actual"
    Rng.Offset(1, 95).Value = "Actual"
    Rng.Offset(1, 96).Value = "Actual"
    Rng.Offset(1, 97).Value = "Actual"
    Rng.Offset(1, 98).Value = "Actual"
    Rng.Offset(1, 99).Value = "Actual"
    Rng.Offset(1, 100).Value = "Actual"
    Rng.Offset(1, 101).Value = "Actual"
    Rng.Offset(1, 102).Value = "Actual"
    Rng.Offset(1, 103).Value = "Actual"
    Rng.Offset(1, 104).Value = "Actual"
    Rng.Offset(1, 105).Value = "Percent"
    Rng.Offset(1, 106).Value = "Percent"
    Rng.Offset(1, 107).Value = "Percent"
    Rng.Offset(1, 108).Value = "Percent"
    Rng.Offset(1, 109).Value = "Percent"
    Rng.Offset(1, 110).Value = "Percent"
    Rng.Offset(1, 111).Value = "Percent"
    Rng.Offset(1, 112).Value = "Percent"
    Rng.Offset(1, 113).Value = "Percent"
    Rng.Offset(1, 114).Value = "Percent"
    Rng.Offset(1, 115).Value = "Percent"
    Rng.Offset(1, 116).Value = "Percent"
    Rng.Offset(1, 117).Value = "Percent"
    Rng.Offset(1, 118).Value = "Percent"
    Rng.Offset(1, 119).Value = "Percent"
    Rng.Offset(1, 120).Value = "Percent"
    Rng.Offset(1, 121).Value = "Percent"
    Rng.Offset(1, 122).Value = "Actual"
    Rng.Offset(1, 123).Value = "Actual"
    Rng.Offset(1, 124).Value = "Actual"
    Rng.Offset(1, 125).Value = "Actual"
    Rng.Offset(1, 126).Value = "Actual"
    Rng.Offset(1, 127).Value = "Actual"
    Rng.Offset(1, 128).Value = "Actual"
    Rng.Offset(1, 129).Value = "Actual"
    Rng.Offset(1, 130).Value = "Actual"
    Rng.Offset(1, 131).Value = "Actual"
    Rng.Offset(1, 132).Value = "Actual"
    Rng.Offset(1, 133).Value = "Actual"
    Rng.Offset(1, 134).Value = "Actual"
    Rng.Offset(1, 135).Value = "Actual"
    Rng.Offset(1, 136).Value = "Actual"
    Rng.Offset(1, 137).Value = "Actual"
    Rng.Offset(1, 138).Value = "Actual"
    Rng.Offset(1, 139).Value = "Percent"
    Rng.Offset(1, 140).Value = "Percent"
    Rng.Offset(1, 141).Value = "Percent"
    Rng.Offset(1, 142).Value = "Percent"
    Rng.Offset(1, 143).Value = "Percent"
    Rng.Offset(1, 144).Value = "Percent"
    Rng.Offset(1, 145).Value = "Percent"
    Rng.Offset(1, 146).Value = "Percent"
    Rng.Offset(1, 147).Value = "Percent"
    Rng.Offset(1, 148).Value = "Percent"
    Rng.Offset(1, 149).Value = "Percent"
    Rng.Offset(1, 150).Value = "Percent"
    Rng.Offset(1, 151).Value = "Percent"
    Rng.Offset(1, 152).Value = "Percent"
    Rng.Offset(1, 153).Value = "Percent"
    Rng.Offset(1, 154).Value = "Percent"
    Rng.Offset(1, 155).Value = "Percent"

'Line headers
    Rng.Offset(2, 0).Value = "Symbol"
    Rng.Offset(2, 1).Value = "Start Date"
    Rng.Offset(2, 2).Value = "End Date"
    Rng.Offset(2, 3).Value = "N"
    Rng.Offset(2, 4).Value = "Minimum"
    Rng.Offset(2, 5).Value = "First Quintile"
    Rng.Offset(2, 6).Value = "First Decile"
    Rng.Offset(2, 7).Value = "Lower Quartile"
    Rng.Offset(2, 8).Value = "Median"
    Rng.Offset(2, 9).Value = "Upper Quartile"
    Rng.Offset(2, 10).Value = "Last Decile"
    Rng.Offset(2, 11).Value = "Last Quintile"
    Rng.Offset(2, 12).Value = "Maximum"
    Rng.Offset(2, 13).Value = "Mode"
    Rng.Offset(2, 14).Value = "Arithmetic Mean"
    Rng.Offset(2, 15).Value = "Variance"
    Rng.Offset(2, 16).Value = "Standard Deviation"
    Rng.Offset(2, 17).Value = "Coefficient Of Variation"
    Rng.Offset(2, 18).Value = "Kurtosis"
    Rng.Offset(2, 19).Value = "Skewness"
    Rng.Offset(2, 20).Value = "N"
    Rng.Offset(2, 21).Value = "Minimum"
    Rng.Offset(2, 22).Value = "First Quintile"
    Rng.Offset(2, 23).Value = "First Decile"
    Rng.Offset(2, 24).Value = "Lower Quartile"
    Rng.Offset(2, 25).Value = "Median"
    Rng.Offset(2, 26).Value = "Upper Quartile"
    Rng.Offset(2, 27).Value = "Last Decile"
    Rng.Offset(2, 28).Value = "Last Quintile"
    Rng.Offset(2, 29).Value = "Maximum"
    Rng.Offset(2, 30).Value = "Mode"
    Rng.Offset(2, 31).Value = "Arithmetic Mean"
    Rng.Offset(2, 32).Value = "Variance"
    Rng.Offset(2, 33).Value = "Standard Deviation"
    Rng.Offset(2, 34).Value = "Coefficient Of Variation"
    Rng.Offset(2, 35).Value = "Kurtosis"
    Rng.Offset(2, 36).Value = "Skewness"
    Rng.Offset(2, 37).Value = "N"
    Rng.Offset(2, 38).Value = "Minimum"
    Rng.Offset(2, 39).Value = "First Quintile"
    Rng.Offset(2, 40).Value = "First Decile"
    Rng.Offset(2, 41).Value = "Lower Quartile"
    Rng.Offset(2, 42).Value = "Median"
    Rng.Offset(2, 43).Value = "Upper Quartile"
    Rng.Offset(2, 44).Value = "Last Decile"
    Rng.Offset(2, 45).Value = "Last Quintile"
    Rng.Offset(2, 46).Value = "Maximum"
    Rng.Offset(2, 47).Value = "Mode"
    Rng.Offset(2, 48).Value = "Arithmetic Mean"
    Rng.Offset(2, 49).Value = "Variance"
    Rng.Offset(2, 50).Value = "Standard Deviation"
    Rng.Offset(2, 51).Value = "Coefficient Of Variation"
    Rng.Offset(2, 52).Value = "Kurtosis"
    Rng.Offset(2, 53).Value = "Skewness"
    Rng.Offset(2, 54).Value = "N"
    Rng.Offset(2, 55).Value = "Minimum"
    Rng.Offset(2, 56).Value = "First Quintile"
    Rng.Offset(2, 57).Value = "First Decile"
    Rng.Offset(2, 58).Value = "Lower Quartile"
    Rng.Offset(2, 59).Value = "Median"
    Rng.Offset(2, 60).Value = "Upper Quartile"
    Rng.Offset(2, 61).Value = "Last Decile"
    Rng.Offset(2, 62).Value = "Last Quintile"
    Rng.Offset(2, 63).Value = "Maximum"
    Rng.Offset(2, 64).Value = "Mode"
    Rng.Offset(2, 65).Value = "Arithmetic Mean"
    Rng.Offset(2, 66).Value = "Variance"
    Rng.Offset(2, 67).Value = "Standard Deviation"
    Rng.Offset(2, 68).Value = "Coefficient Of Variation"
    Rng.Offset(2, 69).Value = "Kurtosis"
    Rng.Offset(2, 70).Value = "Skewness"
    Rng.Offset(2, 71).Value = "N"
    Rng.Offset(2, 72).Value = "Minimum"
    Rng.Offset(2, 73).Value = "First Quintile"
    Rng.Offset(2, 74).Value = "First Decile"
    Rng.Offset(2, 75).Value = "Lower Quartile"
    Rng.Offset(2, 76).Value = "Median"
    Rng.Offset(2, 77).Value = "Upper Quartile"
    Rng.Offset(2, 78).Value = "Last Decile"
    Rng.Offset(2, 79).Value = "Last Quintile"
    Rng.Offset(2, 80).Value = "Maximum"
    Rng.Offset(2, 81).Value = "Mode"
    Rng.Offset(2, 82).Value = "Arithmetic Mean"
    Rng.Offset(2, 83).Value = "Variance"
    Rng.Offset(2, 84).Value = "Standard Deviation"
    Rng.Offset(2, 85).Value = "Coefficient Of Variation"
    Rng.Offset(2, 86).Value = "Kurtosis"
    Rng.Offset(2, 87).Value = "Skewness"
    Rng.Offset(2, 88).Value = "N"
    Rng.Offset(2, 89).Value = "Minimum"
    Rng.Offset(2, 90).Value = "First Quintile"
    Rng.Offset(2, 91).Value = "First Decile"
    Rng.Offset(2, 92).Value = "Lower Quartile"
    Rng.Offset(2, 93).Value = "Median"
    Rng.Offset(2, 94).Value = "Upper Quartile"
    Rng.Offset(2, 95).Value = "Last Decile"
    Rng.Offset(2, 96).Value = "Last Quintile"
    Rng.Offset(2, 97).Value = "Maximum"
    Rng.Offset(2, 98).Value = "Mode"
    Rng.Offset(2, 99).Value = "Arithmetic Mean"
    Rng.Offset(2, 100).Value = "Variance"
    Rng.Offset(2, 101).Value = "Standard Deviation"
    Rng.Offset(2, 102).Value = "Coefficient Of Variation"
    Rng.Offset(2, 103).Value = "Kurtosis"
    Rng.Offset(2, 104).Value = "Skewness"
    Rng.Offset(2, 105).Value = "N"
    Rng.Offset(2, 106).Value = "Minimum"
    Rng.Offset(2, 107).Value = "First Quintile"
    Rng.Offset(2, 108).Value = "First Decile"
    Rng.Offset(2, 109).Value = "Lower Quartile"
    Rng.Offset(2, 110).Value = "Median"
    Rng.Offset(2, 111).Value = "Upper Quartile"
    Rng.Offset(2, 112).Value = "Last Decile"
    Rng.Offset(2, 113).Value = "Last Quintile"
    Rng.Offset(2, 114).Value = "Maximum"
    Rng.Offset(2, 115).Value = "Mode"
    Rng.Offset(2, 116).Value = "Arithmetic Mean"
    Rng.Offset(2, 117).Value = "Variance"
    Rng.Offset(2, 118).Value = "Standard Deviation"
    Rng.Offset(2, 119).Value = "Coefficient Of Variation"
    Rng.Offset(2, 120).Value = "Kurtosis"
    Rng.Offset(2, 121).Value = "Skewness"
    Rng.Offset(2, 122).Value = "N"
    Rng.Offset(2, 123).Value = "Minimum"
    Rng.Offset(2, 124).Value = "First Quintile"
    Rng.Offset(2, 125).Value = "First Decile"
    Rng.Offset(2, 126).Value = "Lower Quartile"
    Rng.Offset(2, 127).Value = "Median"
    Rng.Offset(2, 128).Value = "Upper Quartile"
    Rng.Offset(2, 129).Value = "Last Decile"
    Rng.Offset(2, 130).Value = "Last Quintile"
    Rng.Offset(2, 131).Value = "Maximum"
    Rng.Offset(2, 132).Value = "Mode"
    Rng.Offset(2, 133).Value = "Arithmetic Mean"
    Rng.Offset(2, 134).Value = "Variance"
    Rng.Offset(2, 135).Value = "Standard Deviation"
    Rng.Offset(2, 136).Value = "Coefficient Of Variation"
    Rng.Offset(2, 137).Value = "Kurtosis"
    Rng.Offset(2, 138).Value = "Skewness"
    Rng.Offset(2, 139).Value = "N"
    Rng.Offset(2, 140).Value = "Minimum"
    Rng.Offset(2, 141).Value = "First Quintile"
    Rng.Offset(2, 142).Value = "First Decile"
    Rng.Offset(2, 143).Value = "Lower Quartile"
    Rng.Offset(2, 144).Value = "Median"
    Rng.Offset(2, 145).Value = "Upper Quartile"
    Rng.Offset(2, 146).Value = "Last Decile"
    Rng.Offset(2, 147).Value = "Last Quintile"
    Rng.Offset(2, 148).Value = "Maximum"
    Rng.Offset(2, 149).Value = "Mode"
    Rng.Offset(2, 150).Value = "Arithmetic Mean"
    Rng.Offset(2, 151).Value = "Variance"
    Rng.Offset(2, 152).Value = "Standard Deviation"
    Rng.Offset(2, 153).Value = "Coefficient Of Variation"
    Rng.Offset(2, 154).Value = "Kurtosis"
    Rng.Offset(2, 155).Value = "Skewness"



'center and auto-width those three rows
    
    Range("A1:EZ3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Selection.Columns.AutoFit

'color cells
    Range("D1:T1").Interior.ColorIndex = 6
    Range("U1:BB1").Interior.ColorIndex = 39
    Range("BC1:CJ1").Interior.ColorIndex = 38
    Range("CK1:DR1").Interior.ColorIndex = 35
    Range("DS1:EZ1").Interior.ColorIndex = 34

'freeze panes
    Range("B4").Select
    ActiveWindow.FreezePanes = True
    
'reset cursor
    Range("A1").Select

    
End Function
Function orderDataForGraphing()
    
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
    Selection.NumberFormat = "_(* #,##0.0_);_(* (#,##0.0);_(* ""-""??_);_(@_)"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    Range("C2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "Currency"
    Range("A1").Select
    
End Function
Function manipulateData()
    Dim Rng As Range
    Dim LastRow As Integer
    
    'find the last row
    
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    LastRow = Selection.Rows.Count
    
    'day average average
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Day Average"
    Range("G2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=AVERAGE(RC[-4]:RC[-1])"
    Range("G2").Select
    Selection.AutoFill Destination:=Range("G2:G" & LastRow)
    
    'data manipulations
    Range("H1").Value = "Previous Close to Close"
    Range("H3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-2]-R[-1]C[-2]"
    Range("I3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]/R[-1]C[-3]"
    Range("J1").Value = "Previous Open to Open"
    Range("J3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-7]-R[-1]C[-7]"
    Range("K3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]/R[-1]C[-8]"
    Range("L1").Value = "Previous Close to Open"
    Range("L3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-9]-R[-1]C[-6]"
    Range("M3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]/R[-1]C[-7]"
    Range("N1").Value = "Intraday Open to Close"
    Range("N3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-8]-RC[-11]"
    Range("O3").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]/RC[-12]"
    Range("H3:O3").Select
    Selection.AutoFill Destination:=Range("H3:O" & LastRow)
    Range("I:I,K:K,M:M,O:O").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.000%"
   
End Function
Function populateSummary(SummaryRng As Range)

    Dim WS As Worksheet
    Dim StartDate, EndDate As Date
    Dim Count, LastRow As Integer
    Dim Rng As Range
       
    Dim VolN, VolMinimumVal, VolFirstQuintile, VolFirstDecile, VolLowerQuartile, VolMedian, VolUpperQuartile, VolLastDecile, VolLastQuintile, VolMaximumVal, VolModeVal, VolArithmeticMean, VolVariance, VolStandardDeviation, VolCoefficientOfVariation, VolKurtosis, VolSkewness As Double
    Dim CtoCNActual, CtoCMinimumValActual, CtoCFirstQuintileActual, CtoCFirstDecileActual, CtoCLowerQuartileActual, CtoCMedianActual, CtoCUpperQuartileActual, CtoCLastDecileActual, CtoCLastQuintileActual, CtoCMaximumValActual, CtoCModeValActual, CtoCArithmeticMeanActual, CtoCVarianceActual, CtoCStandardDeviationActual, CtoCCoefficientOfVariationActual, CtoCKurtosisActual, CtoCSkewnessActual, CtoCNPercent, CtoCMinimumValPercent, CtoCFirstQuintilePercent, CtoCFirstDecilePercent, CtoCLowerQuartilePercent, CtoCMedianPercent, CtoCUpperQuartilePercent, CtoCLastDecilePercent, CtoCLastQuintilePercent, CtoCMaximumValPercent, CtoCModeValPercent, CtoCArithmeticMeanPercent, CtoCVariancePercent, CtoCStandardDeviationPercent, CtoCCoefficientOfVariationPercent, CtoCKurtosisPercent, CtoCSkewnessPercent As Double
    Dim OtoONActual, OtoOMinimumValActual, OtoOFirstQuintileActual, OtoOFirstDecileActual, OtoOLowerQuartileActual, OtoOMedianActual, OtoOUpperQuartileActual, OtoOLastDecileActual, OtoOLastQuintileActual, OtoOMaximumValActual, OtoOModeValActual, OtoOArithmeticMeanActual, OtoOVarianceActual, OtoOStandardDeviationActual, OtoOCoefficientOfVariationActual, OtoOKurtosisActual, OtoOSkewnessActual, OtoONPercent, OtoOMinimumValPercent, OtoOFirstQuintilePercent, OtoOFirstDecilePercent, OtoOLowerQuartilePercent, OtoOMedianPercent, OtoOUpperQuartilePercent, OtoOLastDecilePercent, OtoOLastQuintilePercent, OtoOMaximumValPercent, OtoOModeValPercent, OtoOArithmeticMeanPercent, OtoOVariancePercent, OtoOStandardDeviationPercent, OtoOCoefficientOfVariationPercent, OtoOKurtosisPercent, OtoOSkewnessPercent As Double
    Dim CtoONActual, CtoOMinimumValActual, CtoOFirstQuintileActual, CtoOFirstDecileActual, CtoOLowerQuartileActual, CtoOMedianActual, CtoOUpperQuartileActual, CtoOLastDecileActual, CtoOLastQuintileActual, CtoOMaximumValActual, CtoOModeValActual, CtoOArithmeticMeanActual, CtoOVarianceActual, CtoOStandardDeviationActual, CtoOCoefficientOfVariationActual, CtoOKurtosisActual, CtoOSkewnessActual, CtoONPercent, CtoOMinimumValPercent, CtoOFirstQuintilePercent, CtoOFirstDecilePercent, CtoOLowerQuartilePercent, CtoOMedianPercent, CtoOUpperQuartilePercent, CtoOLastDecilePercent, CtoOLastQuintilePercent, CtoOMaximumValPercent, CtoOModeValPercent, CtoOArithmeticMeanPercent, CtoOVariancePercent, CtoOStandardDeviationPercent, CtoOCoefficientOfVariationPercent, CtoOKurtosisPercent, CtoOSkewnessPercent As Double
    Dim OtoCNActual, OtoCMinimumValActual, OtoCFirstQuintileActual, OtoCFirstDecileActual, OtoCLowerQuartileActual, OtoCMedianActual, OtoCUpperQuartileActual, OtoCLastDecileActual, OtoCLastQuintileActual, OtoCMaximumValActual, OtoCModeValActual, OtoCArithmeticMeanActual, OtoCVarianceActual, OtoCStandardDeviationActual, OtoCCoefficientOfVariationActual, OtoCKurtosisActual, OtoCSkewnessActual, OtoCNPercent, OtoCMinimumValPercent, OtoCFirstQuintilePercent, OtoCFirstDecilePercent, OtoCLowerQuartilePercent, OtoCMedianPercent, OtoCUpperQuartilePercent, OtoCLastDecilePercent, OtoCLastQuintilePercent, OtoCMaximumValPercent, OtoCModeValPercent, OtoCArithmeticMeanPercent, OtoCVariancePercent, OtoCStandardDeviationPercent, OtoCCoefficientOfVariationPercent, OtoCKurtosisPercent, OtoCSkewnessPercent As Double

    Set WS = ActiveSheet
    
    WS.Activate
    
    'get dates
    Set Rng = Range("A2")
    Rng.Select
    StartDate = Selection
    Selection.End(xlDown).Select
    EndDate = Selection

    'paste dates
    Worksheets("Summary").Activate
    SummaryRng.Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=WS.Name & "!A1", TextToDisplay:=WS.Name
    SummaryRng.Offset(0, 1).Value = StartDate
    SummaryRng.Offset(0, 2).Value = EndDate
    
    'define volume actual range
    WS.Activate
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set Rng = Selection
    LastRow = Selection.Rows.Count
    
    'calculate volume actual stats
    On Error Resume Next
    VolN = LastRow
    VolMinimumVal = Application.WorksheetFunction.Min(Rng)
    VolFirstQuintile = Application.WorksheetFunction.Percentile(Rng, 0.05)
    VolFirstDecile = Application.WorksheetFunction.Percentile(Rng, 0.1)
    VolLowerQuartile = Application.WorksheetFunction.Percentile(Rng, 0.25)
    VolMedian = Application.WorksheetFunction.Median(Rng)
    VolUpperQuartile = Application.WorksheetFunction.Percentile(Rng, 0.75)
    VolLastDecile = Application.WorksheetFunction.Percentile(Rng, 0.9)
    VolLastQuintile = Application.WorksheetFunction.Percentile(Rng, 0.95)
    VolMaximumVal = Application.WorksheetFunction.Max(Rng)
    VolModeVal = Application.WorksheetFunction.Mode(Rng)
    VolArithmeticMean = Application.WorksheetFunction.Average(Rng)
    VolStandardDeviation = Application.WorksheetFunction.StDev_P(Rng)
    VolVariance = VolStandardDeviation * VolStandardDeviation
    VolCoefficientOfVariation = VolStandardDeviation / VolArithmeticMean
    VolKurtosis = Application.WorksheetFunction.Kurt(Rng)
    VolSkewness = Application.WorksheetFunction.Skew_p(Rng)

    'paste volume actual stats
    Worksheets("Summary").Activate
    SummaryRng.Offset(0, 3).Value = VolN
    SummaryRng.Offset(0, 4).Value = VolMinimumVal
    SummaryRng.Offset(0, 5).Value = VolFirstQuintile
    SummaryRng.Offset(0, 6).Value = VolFirstDecile
    SummaryRng.Offset(0, 7).Value = VolLowerQuartile
    SummaryRng.Offset(0, 8).Value = VolMedian
    SummaryRng.Offset(0, 9).Value = VolUpperQuartile
    SummaryRng.Offset(0, 10).Value = VolLastDecile
    SummaryRng.Offset(0, 11).Value = VolLastQuintile
    SummaryRng.Offset(0, 12).Value = VolMaximumVal
    SummaryRng.Offset(0, 13).Value = VolModeVal
    SummaryRng.Offset(0, 14).Value = VolArithmeticMean
    SummaryRng.Offset(0, 15).Value = VolVariance
    SummaryRng.Offset(0, 16).Value = VolStandardDeviation
    SummaryRng.Offset(0, 17).Value = VolCoefficientOfVariation
    SummaryRng.Offset(0, 18).Value = VolKurtosis
    SummaryRng.Offset(0, 19).Value = VolSkewness


    
    
    
    'define Previous Close to Close actual Range
    WS.Activate
    Range("H3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set Rng = Selection
    LastRow = Selection.Rows.Count
    
    'calculate Previous Close to Close actual stats
    CtoCNActual = LastRow
    CtoCMinimumValActual = Application.WorksheetFunction.Min(Rng)
    CtoCFirstQuintileActual = Application.WorksheetFunction.Percentile(Rng, 0.05)
    CtoCFirstDecileActual = Application.WorksheetFunction.Percentile(Rng, 0.1)
    CtoCLowerQuartileActual = Application.WorksheetFunction.Percentile(Rng, 0.25)
    CtoCMedianActual = Application.WorksheetFunction.Median(Rng)
    CtoCUpperQuartileActual = Application.WorksheetFunction.Percentile(Rng, 0.75)
    CtoCLastDecileActual = Application.WorksheetFunction.Percentile(Rng, 0.9)
    CtoCLastQuintileActual = Application.WorksheetFunction.Percentile(Rng, 0.95)
    CtoCMaximumValActual = Application.WorksheetFunction.Max(Rng)
    CtoCModeValActual = Application.WorksheetFunction.Mode(Rng)
    CtoCArithmeticMeanActual = Application.WorksheetFunction.Average(Rng)
    CtoCStandardDeviationActual = Application.WorksheetFunction.StDev_P(Rng)
    CtoCVarianceActual = CtoCStandardDeviationActual * CtoCStandardDeviationActual
    CtoCCoefficientOfVariationActual = CtoCStandardDeviationActual / CtoCArithmeticMeanActual
    CtoCKurtosisActual = Application.WorksheetFunction.Kurt(Rng)
    CtoCSkewnessActual = Application.WorksheetFunction.Skew_p(Rng)

    'paste Previous Close to Close actual stats
    Worksheets("Summary").Activate
    SummaryRng.Offset(0, 20).Value = CtoCNActual
    SummaryRng.Offset(0, 21).Value = CtoCMinimumValActual
    SummaryRng.Offset(0, 22).Value = CtoCFirstQuintileActual
    SummaryRng.Offset(0, 23).Value = CtoCFirstDecileActual
    SummaryRng.Offset(0, 24).Value = CtoCLowerQuartileActual
    SummaryRng.Offset(0, 25).Value = CtoCMedianActual
    SummaryRng.Offset(0, 26).Value = CtoCUpperQuartileActual
    SummaryRng.Offset(0, 27).Value = CtoCLastDecileActual
    SummaryRng.Offset(0, 28).Value = CtoCLastQuintileActual
    SummaryRng.Offset(0, 29).Value = CtoCMaximumValActual
    SummaryRng.Offset(0, 30).Value = CtoCModeValActual
    SummaryRng.Offset(0, 31).Value = CtoCArithmeticMeanActual
    SummaryRng.Offset(0, 32).Value = CtoCVarianceActual
    SummaryRng.Offset(0, 33).Value = CtoCStandardDeviationActual
    SummaryRng.Offset(0, 34).Value = CtoCCoefficientOfVariationActual
    SummaryRng.Offset(0, 35).Value = CtoCKurtosisActual
    SummaryRng.Offset(0, 36).Value = CtoCSkewnessActual

    
    'define Previous Close to Close percent Range
    WS.Activate
    Range("I3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set Rng = Selection
    LastRow = Selection.Rows.Count
    
    'calculate Previous Close to Close percent stats
    CtoCNPercent = LastRow
    CtoCMinimumValPercent = Application.WorksheetFunction.Min(Rng)
    CtoCFirstQuintilePercent = Application.WorksheetFunction.Percentile(Rng, 0.05)
    CtoCFirstDecilePercent = Application.WorksheetFunction.Percentile(Rng, 0.1)
    CtoCLowerQuartilePercent = Application.WorksheetFunction.Percentile(Rng, 0.25)
    CtoCMedianPercent = Application.WorksheetFunction.Median(Rng)
    CtoCUpperQuartilePercent = Application.WorksheetFunction.Percentile(Rng, 0.75)
    CtoCLastDecilePercent = Application.WorksheetFunction.Percentile(Rng, 0.9)
    CtoCLastQuintilePercent = Application.WorksheetFunction.Percentile(Rng, 0.95)
    CtoCMaximumValPercent = Application.WorksheetFunction.Max(Rng)
    CtoCArithmeticMeanPercent = Application.WorksheetFunction.Average(Rng)
    CtoCStandardDeviationPercent = Application.WorksheetFunction.StDev_P(Rng)
    CtoCVariancePercent = CtoCStandardDeviationPercent * CtoCStandardDeviationPercent
    CtoCCoefficientOfVariationPercent = CtoCStandardDeviationPercent / CtoCArithmeticMeanPercent
    CtoCKurtosisPercent = Application.WorksheetFunction.Kurt(Rng)
    CtoCSkewnessPercent = Application.WorksheetFunction.Skew_p(Rng)

    'paste Previous Close to Close percent stats
    Worksheets("Summary").Activate
    SummaryRng.Offset(0, 37).Value = CtoCNPercent
    SummaryRng.Offset(0, 38).Value = CtoCMinimumValPercent
    SummaryRng.Offset(0, 39).Value = CtoCFirstQuintilePercent
    SummaryRng.Offset(0, 40).Value = CtoCFirstDecilePercent
    SummaryRng.Offset(0, 41).Value = CtoCLowerQuartilePercent
    SummaryRng.Offset(0, 42).Value = CtoCMedianPercent
    SummaryRng.Offset(0, 43).Value = CtoCUpperQuartilePercent
    SummaryRng.Offset(0, 44).Value = CtoCLastDecilePercent
    SummaryRng.Offset(0, 45).Value = CtoCLastQuintilePercent
    SummaryRng.Offset(0, 46).Value = CtoCMaximumValPercent
    SummaryRng.Offset(0, 47).Value = CtoCModeValPercent
    SummaryRng.Offset(0, 48).Value = CtoCArithmeticMeanPercent
    SummaryRng.Offset(0, 49).Value = CtoCVariancePercent
    SummaryRng.Offset(0, 50).Value = CtoCStandardDeviationPercent
    SummaryRng.Offset(0, 51).Value = CtoCCoefficientOfVariationPercent
    SummaryRng.Offset(0, 52).Value = CtoCKurtosisPercent
    SummaryRng.Offset(0, 53).Value = CtoCSkewnessPercent




    'define Previous Open to Open actual Range
    WS.Activate
    Range("J3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set Rng = Selection
    LastRow = Selection.Rows.Count
    
    'calculate Previous Open to Open actual stats
    OtoONActual = LastRow
    OtoOMinimumValActual = Application.WorksheetFunction.Min(Rng)
    OtoOFirstQuintileActual = Application.WorksheetFunction.Percentile(Rng, 0.05)
    OtoOFirstDecileActual = Application.WorksheetFunction.Percentile(Rng, 0.1)
    OtoOLowerQuartileActual = Application.WorksheetFunction.Percentile(Rng, 0.25)
    OtoOMedianActual = Application.WorksheetFunction.Median(Rng)
    OtoOUpperQuartileActual = Application.WorksheetFunction.Percentile(Rng, 0.75)
    OtoOLastDecileActual = Application.WorksheetFunction.Percentile(Rng, 0.9)
    OtoOLastQuintileActual = Application.WorksheetFunction.Percentile(Rng, 0.95)
    OtoOMaximumValActual = Application.WorksheetFunction.Max(Rng)
    OtoOModeValActual = Application.WorksheetFunction.Mode(Rng)
    OtoOArithmeticMeanActual = Application.WorksheetFunction.Average(Rng)
    OtoOVarianceActual = OtoOStandardDeviationActual * OtoOStandardDeviationActual
    OtoOCoefficientOfVariationActual = OtoOStandardDeviationActual / OtoOArithmeticMeanActual
    OtoOStandardDeviationActual = Application.WorksheetFunction.StDev_P(Rng)
    OtoOKurtosisActual = Application.WorksheetFunction.Kurt(Rng)
    OtoOSkewnessActual = Application.WorksheetFunction.Skew_p(Rng)

    
    'paste Previous Open to Open actual stats
    Worksheets("Summary").Activate
    SummaryRng.Offset(0, 54).Value = OtoONActual
    SummaryRng.Offset(0, 55).Value = OtoOMinimumValActual
    SummaryRng.Offset(0, 56).Value = OtoOFirstQuintileActual
    SummaryRng.Offset(0, 57).Value = OtoOFirstDecileActual
    SummaryRng.Offset(0, 58).Value = OtoOLowerQuartileActual
    SummaryRng.Offset(0, 59).Value = OtoOMedianActual
    SummaryRng.Offset(0, 60).Value = OtoOUpperQuartileActual
    SummaryRng.Offset(0, 61).Value = OtoOLastDecileActual
    SummaryRng.Offset(0, 62).Value = OtoOLastQuintileActual
    SummaryRng.Offset(0, 63).Value = OtoOMaximumValActual
    SummaryRng.Offset(0, 64).Value = OtoOModeValActual
    SummaryRng.Offset(0, 65).Value = OtoOArithmeticMeanActual
    SummaryRng.Offset(0, 66).Value = OtoOVarianceActual
    SummaryRng.Offset(0, 67).Value = OtoOStandardDeviationActual
    SummaryRng.Offset(0, 68).Value = OtoOCoefficientOfVariationActual
    SummaryRng.Offset(0, 69).Value = OtoOKurtosisActual
    SummaryRng.Offset(0, 70).Value = OtoOSkewnessActual


    'define Previous Open to Open percent Range
    WS.Activate
    Range("K3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set Rng = Selection
    LastRow = Selection.Rows.Count
    
    'calculate Previous Open to Open percent stats
    OtoONPercent = LastRow
    OtoOMinimumValPercent = Application.WorksheetFunction.Min(Rng)
    OtoOFirstQuintilePercent = Application.WorksheetFunction.Percentile(Rng, 0.05)
    OtoOFirstDecilePercent = Application.WorksheetFunction.Percentile(Rng, 0.1)
    OtoOLowerQuartilePercent = Application.WorksheetFunction.Percentile(Rng, 0.25)
    OtoOMedianPercent = Application.WorksheetFunction.Median(Rng)
    OtoOUpperQuartilePercent = Application.WorksheetFunction.Percentile(Rng, 0.75)
    OtoOLastDecilePercent = Application.WorksheetFunction.Percentile(Rng, 0.9)
    OtoOLastQuintilePercent = Application.WorksheetFunction.Percentile(Rng, 0.95)
    OtoOMaximumValPercent = Application.WorksheetFunction.Max(Rng)
    OtoOModeValPercent = Application.WorksheetFunction.Mode(Rng)
    OtoOArithmeticMeanPercent = Application.WorksheetFunction.Average(Rng)
    OtoOStandardDeviationPercent = Application.WorksheetFunction.StDev_P(Rng)
    OtoOVariancePercent = OtoOStandardDeviationPercent * OtoOStandardDeviationPercent
    OtoOCoefficientOfVariationPercent = OtoOStandardDeviationPercent / OtoOArithmeticMeanPercent
    OtoOKurtosisPercent = Application.WorksheetFunction.Kurt(Rng)
    OtoOSkewnessPercent = Application.WorksheetFunction.Skew_p(Rng)

    
    'paste Previous Open to Open percent stats
    Worksheets("Summary").Activate
    SummaryRng.Offset(0, 71).Value = OtoONPercent
    SummaryRng.Offset(0, 72).Value = OtoOMinimumValPercent
    SummaryRng.Offset(0, 73).Value = OtoOFirstQuintilePercent
    SummaryRng.Offset(0, 74).Value = OtoOFirstDecilePercent
    SummaryRng.Offset(0, 75).Value = OtoOLowerQuartilePercent
    SummaryRng.Offset(0, 76).Value = OtoOMedianPercent
    SummaryRng.Offset(0, 77).Value = OtoOUpperQuartilePercent
    SummaryRng.Offset(0, 78).Value = OtoOLastDecilePercent
    SummaryRng.Offset(0, 79).Value = OtoOLastQuintilePercent
    SummaryRng.Offset(0, 80).Value = OtoOMaximumValPercent
    SummaryRng.Offset(0, 81).Value = OtoOModeValPercent
    SummaryRng.Offset(0, 82).Value = OtoOArithmeticMeanPercent
    SummaryRng.Offset(0, 83).Value = OtoOVariancePercent
    SummaryRng.Offset(0, 84).Value = OtoOStandardDeviationPercent
    SummaryRng.Offset(0, 85).Value = OtoOCoefficientOfVariationPercent
    SummaryRng.Offset(0, 86).Value = OtoOKurtosisPercent
    SummaryRng.Offset(0, 87).Value = OtoOSkewnessPercent



    'define Previous Close to Open actual Range
    WS.Activate
    Range("L3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set Rng = Selection
    LastRow = Selection.Rows.Count
    
    'calculate Previous Close to Open actual stats
    CtoONActual = LastRow
    CtoOMinimumValActual = Application.WorksheetFunction.Min(Rng)
    CtoOFirstQuintileActual = Application.WorksheetFunction.Percentile(Rng, 0.05)
    CtoOFirstDecileActual = Application.WorksheetFunction.Percentile(Rng, 0.1)
    CtoOLowerQuartileActual = Application.WorksheetFunction.Percentile(Rng, 0.25)
    CtoOMedianActual = Application.WorksheetFunction.Median(Rng)
    CtoOUpperQuartileActual = Application.WorksheetFunction.Percentile(Rng, 0.75)
    CtoOLastDecileActual = Application.WorksheetFunction.Percentile(Rng, 0.9)
    CtoOLastQuintileActual = Application.WorksheetFunction.Percentile(Rng, 0.95)
    CtoOMaximumValActual = Application.WorksheetFunction.Max(Rng)
    CtoOModeValActual = Application.WorksheetFunction.Mode(Rng)
    CtoOArithmeticMeanActual = Application.WorksheetFunction.Average(Rng)
    CtoOStandardDeviationActual = Application.WorksheetFunction.StDev_P(Rng)
    CtoOVarianceActual = CtoOStandardDeviationActual * CtoOStandardDeviationActual
    CtoOCoefficientOfVariationActual = CtoOStandardDeviationActual / CtoOArithmeticMeanActual
    CtoOKurtosisActual = Application.WorksheetFunction.Kurt(Rng)
    CtoOSkewnessActual = Application.WorksheetFunction.Skew_p(Rng)

    
    'paste Previous Close to Open actual stats
    Worksheets("Summary").Activate
    SummaryRng.Offset(0, 88).Value = CtoONActual
    SummaryRng.Offset(0, 89).Value = CtoOMinimumValActual
    SummaryRng.Offset(0, 90).Value = CtoOFirstQuintileActual
    SummaryRng.Offset(0, 91).Value = CtoOFirstDecileActual
    SummaryRng.Offset(0, 92).Value = CtoOLowerQuartileActual
    SummaryRng.Offset(0, 93).Value = CtoOMedianActual
    SummaryRng.Offset(0, 94).Value = CtoOUpperQuartileActual
    SummaryRng.Offset(0, 95).Value = CtoOLastDecileActual
    SummaryRng.Offset(0, 96).Value = CtoOLastQuintileActual
    SummaryRng.Offset(0, 97).Value = CtoOMaximumValActual
    SummaryRng.Offset(0, 98).Value = CtoOModeValActual
    SummaryRng.Offset(0, 99).Value = CtoOArithmeticMeanActual
    SummaryRng.Offset(0, 100).Value = CtoOVarianceActual
    SummaryRng.Offset(0, 101).Value = CtoOStandardDeviationActual
    SummaryRng.Offset(0, 102).Value = CtoOCoefficientOfVariationActual
    SummaryRng.Offset(0, 103).Value = CtoOKurtosisActual
    SummaryRng.Offset(0, 104).Value = CtoOSkewnessActual



    'define Previous Close to Open percent Range
    WS.Activate
    Range("M3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set Rng = Selection
    LastRow = Selection.Rows.Count
    
    'calculate Previous Close to Open percent stats
    CtoONPercent = LastRow
    CtoOMinimumValPercent = Application.WorksheetFunction.Min(Rng)
    CtoOFirstQuintilePercent = Application.WorksheetFunction.Percentile(Rng, 0.05)
    CtoOFirstDecilePercent = Application.WorksheetFunction.Percentile(Rng, 0.1)
    CtoOLowerQuartilePercent = Application.WorksheetFunction.Percentile(Rng, 0.25)
    CtoOMedianPercent = Application.WorksheetFunction.Median(Rng)
    CtoOUpperQuartilePercent = Application.WorksheetFunction.Percentile(Rng, 0.75)
    CtoOLastDecilePercent = Application.WorksheetFunction.Percentile(Rng, 0.9)
    CtoOLastQuintilePercent = Application.WorksheetFunction.Percentile(Rng, 0.95)
    CtoOMaximumValPercent = Application.WorksheetFunction.Max(Rng)
    CtoOModeValPercent = Application.WorksheetFunction.Mode(Rng)
    CtoOArithmeticMeanPercent = Application.WorksheetFunction.Average(Rng)
    CtoOStandardDeviationPercent = Application.WorksheetFunction.StDev_P(Rng)
    CtoOVariancePercent = CtoOStandardDeviationPercent * CtoOStandardDeviationPercent
    CtoOCoefficientOfVariationPercent = CtoOStandardDeviationPercent / CtoOArithmeticMeanPercent
    CtoOKurtosisPercent = Application.WorksheetFunction.Kurt(Rng)
    CtoOSkewnessPercent = Application.WorksheetFunction.Skew_p(Rng)

    
    'paste Previous Close to Open percent stats
    Worksheets("Summary").Activate
    SummaryRng.Offset(0, 105).Value = CtoONPercent
    SummaryRng.Offset(0, 106).Value = CtoOMinimumValPercent
    SummaryRng.Offset(0, 107).Value = CtoOFirstQuintilePercent
    SummaryRng.Offset(0, 108).Value = CtoOFirstDecilePercent
    SummaryRng.Offset(0, 109).Value = CtoOLowerQuartilePercent
    SummaryRng.Offset(0, 110).Value = CtoOMedianPercent
    SummaryRng.Offset(0, 111).Value = CtoOUpperQuartilePercent
    SummaryRng.Offset(0, 112).Value = CtoOLastDecilePercent
    SummaryRng.Offset(0, 113).Value = CtoOLastQuintilePercent
    SummaryRng.Offset(0, 114).Value = CtoOMaximumValPercent
    SummaryRng.Offset(0, 115).Value = CtoOModeValPercent
    SummaryRng.Offset(0, 116).Value = CtoOArithmeticMeanPercent
    SummaryRng.Offset(0, 117).Value = CtoOVariancePercent
    SummaryRng.Offset(0, 118).Value = CtoOStandardDeviationPercent
    SummaryRng.Offset(0, 119).Value = CtoOCoefficientOfVariationPercent
    SummaryRng.Offset(0, 120).Value = CtoOKurtosisPercent
    SummaryRng.Offset(0, 121).Value = CtoOSkewnessPercent



    
    'define Intraday Open to Close actual Range
    WS.Activate
    Range("N3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set Rng = Selection
    LastRow = Selection.Rows.Count
    
    'calculate Intraday Open to Close actual stats
    OtoCNActual = LastRow
    OtoCMinimumValActual = Application.WorksheetFunction.Min(Rng)
    OtoCFirstQuintileActual = Application.WorksheetFunction.Percentile(Rng, 0.05)
    OtoCFirstDecileActual = Application.WorksheetFunction.Percentile(Rng, 0.1)
    OtoCLowerQuartileActual = Application.WorksheetFunction.Percentile(Rng, 0.25)
    OtoCMedianActual = Application.WorksheetFunction.Median(Rng)
    OtoCUpperQuartileActual = Application.WorksheetFunction.Percentile(Rng, 0.75)
    OtoCLastDecileActual = Application.WorksheetFunction.Percentile(Rng, 0.9)
    OtoCLastQuintileActual = Application.WorksheetFunction.Percentile(Rng, 0.95)
    OtoCMaximumValActual = Application.WorksheetFunction.Max(Rng)
    OtoCModeValActual = Application.WorksheetFunction.Mode(Rng)
    OtoCArithmeticMeanActual = Application.WorksheetFunction.Average(Rng)
    OtoCStandardDeviationActual = Application.WorksheetFunction.StDev_P(Rng)
    OtoCVarianceActual = OtoCStandardDeviationActual * OtoCStandardDeviationActual
    OtoCCoefficientOfVariationActual = OtoCStandardDeviationActual / OtoCArithmeticMeanActual
    OtoCKurtosisActual = Application.WorksheetFunction.Kurt(Rng)
    OtoCSkewnessActual = Application.WorksheetFunction.Skew_p(Rng)

    
    'paste Intraday Open to Close actual stats
    Worksheets("Summary").Activate
    SummaryRng.Offset(0, 122).Value = OtoCNActual
    SummaryRng.Offset(0, 123).Value = OtoCMinimumValActual
    SummaryRng.Offset(0, 124).Value = OtoCFirstQuintileActual
    SummaryRng.Offset(0, 125).Value = OtoCFirstDecileActual
    SummaryRng.Offset(0, 126).Value = OtoCLowerQuartileActual
    SummaryRng.Offset(0, 127).Value = OtoCMedianActual
    SummaryRng.Offset(0, 128).Value = OtoCUpperQuartileActual
    SummaryRng.Offset(0, 129).Value = OtoCLastDecileActual
    SummaryRng.Offset(0, 130).Value = OtoCLastQuintileActual
    SummaryRng.Offset(0, 131).Value = OtoCMaximumValActual
    SummaryRng.Offset(0, 132).Value = OtoCModeValActual
    SummaryRng.Offset(0, 133).Value = OtoCArithmeticMeanActual
    SummaryRng.Offset(0, 134).Value = OtoCVarianceActual
    SummaryRng.Offset(0, 135).Value = OtoCStandardDeviationActual
    SummaryRng.Offset(0, 136).Value = OtoCCoefficientOfVariationActual
    SummaryRng.Offset(0, 137).Value = OtoCKurtosisActual
    SummaryRng.Offset(0, 138).Value = OtoCSkewnessActual

    
    'define Intraday Open to Close percent Range
    WS.Activate
    Range("O3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set Rng = Selection
    LastRow = Selection.Rows.Count
    
    'calculate Previous Open to Close percent stats
    OtoCNPercent = LastRow
    OtoCMinimumValPercent = Application.WorksheetFunction.Min(Rng)
    OtoCFirstQuintilePercent = Application.WorksheetFunction.Percentile(Rng, 0.05)
    OtoCFirstDecilePercent = Application.WorksheetFunction.Percentile(Rng, 0.1)
    OtoCLowerQuartilePercent = Application.WorksheetFunction.Percentile(Rng, 0.25)
    OtoCMedianPercent = Application.WorksheetFunction.Median(Rng)
    OtoCUpperQuartilePercent = Application.WorksheetFunction.Percentile(Rng, 0.75)
    OtoCLastDecilePercent = Application.WorksheetFunction.Percentile(Rng, 0.9)
    OtoCLastQuintilePercent = Application.WorksheetFunction.Percentile(Rng, 0.95)
    OtoCMaximumValPercent = Application.WorksheetFunction.Max(Rng)
    OtoCModeValPercent = Application.WorksheetFunction.Mode(Rng)
    OtoCArithmeticMeanPercent = Application.WorksheetFunction.Average(Rng)
    OtoCStandardDeviationPercent = Application.WorksheetFunction.StDev_P(Rng)
    OtoCVariancePercent = OtoCStandardDeviationPercent * OtoCStandardDeviationPercent
    OtoCCoefficientOfVariationPercent = OtoCStandardDeviationPercent / OtoCArithmeticMeanPercent
    OtoCKurtosisPercent = Application.WorksheetFunction.Kurt(Rng)
    OtoCSkewnessPercent = Application.WorksheetFunction.Skew_p(Rng)

    
    'paste Intraday Open to Close percent stats
    Worksheets("Summary").Activate
    SummaryRng.Offset(0, 139).Value = OtoCNPercent
    SummaryRng.Offset(0, 140).Value = OtoCMinimumValPercent
    SummaryRng.Offset(0, 141).Value = OtoCFirstQuintilePercent
    SummaryRng.Offset(0, 142).Value = OtoCFirstDecilePercent
    SummaryRng.Offset(0, 143).Value = OtoCLowerQuartilePercent
    SummaryRng.Offset(0, 144).Value = OtoCMedianPercent
    SummaryRng.Offset(0, 145).Value = OtoCUpperQuartilePercent
    SummaryRng.Offset(0, 146).Value = OtoCLastDecilePercent
    SummaryRng.Offset(0, 147).Value = OtoCLastQuintilePercent
    SummaryRng.Offset(0, 148).Value = OtoCMaximumValPercent
    SummaryRng.Offset(0, 149).Value = OtoCModeValPercent
    SummaryRng.Offset(0, 150).Value = OtoCArithmeticMeanPercent
    SummaryRng.Offset(0, 151).Value = OtoCVariancePercent
    SummaryRng.Offset(0, 152).Value = OtoCStandardDeviationPercent
    SummaryRng.Offset(0, 153).Value = OtoCCoefficientOfVariationPercent
    SummaryRng.Offset(0, 154).Value = OtoCKurtosisPercent
    SummaryRng.Offset(0, 155).Value = OtoCSkewnessPercent
    
End Function
Function CorrelationTable()

    Dim Count As Integer
    Dim Rng1 As Range
    Dim Rng2 As Range
    Dim Rng3 As Range
    Dim Rng4 As Range
    Dim Company1 As String
    Dim Company2 As String
    Dim TrgtRng As Range
    Dim i As Integer
    Dim j As Integer
    Dim CorrelationVar As Single
    
        Worksheets("Summary").Select
        Range("A4").Select
        Range(Selection, Selection.End(xlDown)).Select
        Count = Selection.Rows.Count
        Selection.Copy
        
        
        Sheets.Add.Name = "CorrelationPage"
        Worksheets("CorrelationPage").Select
        Range("A2").Select
        ActiveSheet.Paste
        Range("A2").Select
        Set Rng1 = Selection
        Company1 = Rng1.Value
        
        
        Range("B1").Select
        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Range("B1").Select
        Set Rng2 = Selection
        Company2 = Rng1.Value
        
        Range("B2").Select
        Set TrgtRng = Selection
        
        For i = 1 To Count:
            Worksheets(Company1).Select
            Range("O3").Select
            Range(Selection, Selection.End(xlDown)).Select
            Set Rng3 = Selection
            For j = 1 To Count
            On Error GoTo -1
                Worksheets(Company2).Select
                Range("O3").Select
                Range(Selection, Selection.End(xlDown)).Select
                Set Rng4 = Selection
                If (Rng3.Rows.Count <> Rng4.Rows.Count) Then
                    Worksheets("CorrelationPage").Select
                    TrgtRng = "N/A"
                Else
                    CorrelationVar = Application.WorksheetFunction.Correl(Rng3, Rng4)
                    Worksheets("CorrelationPage").Select
                    TrgtRng = CorrelationVar
                End If
                TrgtRng.Offset(0, 1).Select
                Set TrgtRng = Selection
                Rng2.Offset(0, 1).Select
                Set Rng2 = Selection
                Company2 = Rng2
            Next j
            Worksheets("CorrelationPage").Select
            Range("B1").Select
            Set Rng2 = Selection
            Company2 = Rng2.Value
            Rng1.Offset(1, 1).Select
            Set TrgtRng = Selection
            Rng1.Offset(1, 0).Select
            Set Rng1 = Selection
            Company1 = Rng1
        Next i
        
        'Color code shit
        
        Range("B2").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        
        Selection.FormatConditions.AddColorScale ColorScaleType:=3
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        Selection.FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueNumber
        Selection.FormatConditions(1).ColorScaleCriteria(1).Value = -1
        
        With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
            .Color = 7039480
            .TintAndShade = 0
        End With
        
        Selection.FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValueNumber
        Selection.FormatConditions(1).ColorScaleCriteria(2).Value = 0
        
        With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        
        Selection.FormatConditions(1).ColorScaleCriteria(3).Type = xlConditionValueNumber
        Selection.FormatConditions(1).ColorScaleCriteria(3).Value = 1
        
        With Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor
            .Color = 8109667
            .TintAndShade = 0
        End With
        
        'bold first row and column
        Rows("1:1").Select
        Selection.Font.Bold = True
        Columns("A:A").Select
        Selection.Font.Bold = True
        
        'Freeze Panes
        Range("B2").Select
        ActiveWindow.FreezePanes = True
        
        'Reset selected cell for good measure
        Range("A1").Select

End Function
