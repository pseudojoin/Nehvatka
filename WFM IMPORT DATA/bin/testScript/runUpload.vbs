Set WshShell = CreateObject("WScript.Shell")
WshShell.Run("uploadMagicOut.cmd")
Btn = WshShell.Popup("Message", 3, "Handle", 0+64)
Set Btn = Nothing
Set	WshShell = Nothing
Set objIE = CreateObject("InternetExplorer.Application")
objIE.Visible = True
objIE.Navigate("http://cs/magic/frame/cmk-1796.php?stat=good")
While objIE.ReadyState <> 4 : WScript.Sleep 3000 : Wend
objIE.Quit
set objIE = Nothing