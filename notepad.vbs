WScript.Sleep 10
WScript.Sleep 10
Set WshShell = WScript.CreateObject("WScript.Shell")
do
WshShell.Run "notepad"
WScript.Sleep 100
WshShell.AppActivate "notepad"
loop 