Dim Excel
OfficeBitness=0
Set Excel = CreateObject("Excel.Application")
Excel.Visible = False
If InStr(Excel.OperatingSystem,"64") > 0 Then
    OfficeBitness=2
Else
    OfficeBitness=1
End if
Excel.Quit
Set Excel = Nothing
wscript.Quit(OfficeBitness)