set WshShell = WScript.CreateObject("WScript.Shell")

WScript.Sleep( 3000 )

WshShell.SendKeys "%{TAB}"
WshShell.SendKeys "{ENTER}"
