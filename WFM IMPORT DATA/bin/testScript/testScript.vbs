Dim path, WshShell
Dim curDate, curTime, TimeExev
Set WshShell = CreateObject("WScript.Shell")

curDate = Date()
curTime = Time()
MsgBox curDate
MsgBox curTime

TimeExev = curDate + CDate("17:00:00")
MsgBox TimeExev

If Now() > TimeExev then
	MsgBox "Data is Greatest when Interval was Open"
Else
	MsgBox "Data is Less when Interval was Open"
End if

Btn1 = WshShell.Popup("Loading is " & FormatDateTime(TimeExev), 3, "Handle", 0+64)