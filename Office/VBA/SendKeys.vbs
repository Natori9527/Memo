Set WshShell = WScript.CreateObject("WScript.Shell")
'Cltr + Enter
WshShell.SendKeys "^{ENTER}"
'Enter
WshShell.SendKeys "{ENTER}"
'TAB
WshShell.SendKeys "{TAB}"
'↓
WshShell.SendKeys "{DOWN}"
WshShell.SendKeys "{ENTER}"
