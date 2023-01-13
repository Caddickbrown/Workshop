'Initialise
Set WshShell = WScript.CreateObject("WScript.Shell")
'Create KeepAwake
Call WshShell.Run("%windir%\system32\notepad.exe")
WScript.Sleep(500)
WshShell.SendKeys("Set WshShell = WScript.CreateObject" & "{(}" & chr(34) & "WScript.Shell" & chr(34) & "{)}")
WshShell.SendKeys("{ENTER}")
WshShell.SendKeys("Do While True")
WshShell.SendKeys("{ENTER}")
WshShell.SendKeys("{TAB}" & "WshShell.SendKeys" & "{(}" & chr(34) & "{{}" & "F15" & "{}}" & chr(34) & "{)}" )
WshShell.SendKeys("{ENTER}")
WshShell.SendKeys("{TAB}" & "WshShell.SendKeys" & "{(}" & "55000" & "{)}")
WshShell.SendKeys("{ENTER}")
WshShell.SendKeys("Loop")
 
'Save KeepAwake
WshShell.SendKeys("^S")
WshShell.SendKeys("KeepAwake.vbs")
WshShell.SendKeys("{TAB}")
WshShell.SendKeys("{DOWN}{DOWN}")
WshShell.SendKeys("{ENTER}{ENTER}")
WScript.Sleep(500)
WshShell.SendKeys("%{F4}")
 
'Open Websites
'Website 1
WshShell.Run """https://www.gmail.com/""", 0, TRUE
'Website 2
WshShell.Run """https://www.gmail.com/""", 0, TRUE
 
 