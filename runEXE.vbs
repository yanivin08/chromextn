Set WshShell = WScript.CreateObject("WScript.Shell")

Dim exeName

exeName = "skype.exe"
  '"C:\\Users\arvin76560\Desktop\Notepad++\Notepad++\notepad++.exe"

WshShell.Run exeName, 1, true
