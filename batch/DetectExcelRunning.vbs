' DetectExcelRunning.vbs
' Find an invisible instance of Excel

Dim objXL, strMessage, ExcelRunning

ExcelRunning = 0

On Error Resume Next

' Try to grab a running instance of Excel:
Set objXL = GetObject(, "Excel.Application")

' What have we found?
If Not TypeName(objXL) = "Empty" Then
	'Excel is running
	ExcelRunning=1	
	strMessage = "Excel Running."
Else
	'Excel is not running
    ExcelRunning=2
	strMessage = "Excel Not Running."
End if

' Feedback to user...
'MsgBox strMessage, vbInformation, "Excel Status"

wscript.Quit(ExcelRunning)
' End of VBS code
