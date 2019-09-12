Dim stDate
stDate = Now()

Set objShell = CreateObject("Wscript.Shell")
strPath = Wscript.ScriptFullName
pathProj = Left(strPath, InStrRev(strPath, "\"))

Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(pathProj & "NICE WMF get intraday.xlsm")
objExcel.Visible = True
objExcel.Application.Run "deleteReportWFM"
objWorkbook.Close(False)
Set objWorkbook = Nothing
objExcel.Quit
Set objExcel = Nothing

Btn1 = objShell.Popup("Loading is " & FormatDateTime(Now() - stDate, vbLongTime), 2, "all cleared", 0+64)
Set Btn1 = Nothing
Set objShell = Nothing