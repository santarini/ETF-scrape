Dim CurrentSheet As Worksheet

For Each CurrentSheet In Worksheets
	If INSTR(1, Current.Name, "(Mon)")>0 then
		Call dailySummary
	If INSTR(1, Current.Name, "(Day)")>0 then
		Call monthlySummary	
	EndIf
Next
