'This VBScript Puts together the initial files and opens websites ready to setup a computer for the first time.
'Initialise
Set WshShell = WScript.CreateObject("WScript.Shell")
'Create KeepAwake
Call WshShell.Run("%windir%\system32\notepad.exe")
WScript.Sleep(500)
WshShell.SendKeys("Set WshShell = WScript.CreateObject" & "{(}" & chr(34) & "WScript.Shell" & chr(34) & "{)}" & "{ENTER}" & "Do While True" & "{ENTER}" & "{TAB}" & "WshShell.SendKeys" & "{(}" & chr(34) & "{{}" & "F15" & "{}}" & chr(34) & "{)}" & "{ENTER}" & "{TAB}" & "WshShell.SendKeys" & "{(}" & "55000" & "{)}" & "{ENTER}" & "Loop")

WScript.Sleep(5000)
 
'Save KeepAwake
WshShell.SendKeys("^S")
WshShell.SendKeys("KeepAwake.vbs")
WshShell.SendKeys("{TAB}")
WshShell.SendKeys("{DOWN}{DOWN}")
WshShell.SendKeys("{ENTER}{ENTER}")
WScript.Sleep(500)
WshShell.SendKeys("%{F4}")

'Open Websites
'VS Code
'WshShell.Run """https://code.visualstudio.com/""", 0, TRUE
'Obsidian
'WshShell.Run """https://www.obsidian.md/""", 0, TRUE
'Jrnl
'WshShell.Run """https://jrnl.sh/en/stable/""", 0, TRUE
 
 