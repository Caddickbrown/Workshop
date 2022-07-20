'This script will keep your computer awake indefinitely - be aware, this can delete some cells out of Google Sheets

Set WshShell = WScript.CreateObject("WScript.Shell")
Do While True
        WshShell.SendKeys("{F15}")
        WScript.Sleep(55000)
Loop
