Dim stDate, OpenTimeToExec, ClosedTimeToExec
stDate = Now()

OpenTimeToExec = Date() + CDate("07:00:00")
ClosedTimeToExec = Date() + CDate("21:00:00")

If stDate >= OpenTimeToExec And stDate <= ClosedTimeToExec Then
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("C:\ivolobuev\WFM IMPORT DATA\bin\NICE WMF get intraday.xlsm")
objExcel.Visible = True
objExcel.Application.Run "wMainRun"
objWorkbook.Close(False)
Set objWorkbook = Nothing
objExcel.Quit
Set objExcel = Nothing
Set WshShell = CreateObject("WScript.Shell")
Dim path
path = "C:\ivolobuev\WFM IMPORT DATA\bin\"
WshShell.Run Chr(34) & path & "uploadMagicOut.cmd" & Chr(34), 1, True
WScript.Sleep 5000
Btn1 = WshShell.Popup("Loading is " & FormatDateTime(Now() - stDate, vbLongTime), 2, "Handle", 0+64)
Set objIE = CreateObject("InternetExplorer.Application")
objIE.Visible = True
objIE.Navigate("http://cs/magic/frame/cmk-1796.php?stat=good")
While objIE.ReadyState <> 4 : WScript.Sleep 5000 : Wend
objIE.Quit
set objIE = Nothing
Btn2 = WshShell.Popup("Execution Time is " & FormatDateTime(Now() - stDate, vbLongTime), 3, "FINISH ( . )( . )", 0+64)
Set Btn1 = Nothing
Set Btn2 = Nothing
Set WshShell = Nothing
End If