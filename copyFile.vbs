Dim oFSO
Set oFSO = CreateObject("Scripting.FileSystemObject")

' Create a new folder
  'oFSO.CreateFolder "C:\\Users\arvin76560\Desktop\Notepad++"

' Copy a file into the new folder
' Note that the destination folder path must end with a path separator (\)
  'oFSO.CopyFolder "\\fs01\quality & process excellence\Process Excellence\Audit-QPE\TTL Philippines\Tools\Notepad++", "C:\\Users\arvin76560\Desktop\Notepad++"

  oFSO.CopyFile "\\fs01\quality & process excellence\Process Excellence\Audit-QPE\TTL Philippines\Coder\abayani\backup\XPO\usZip.b2h","C:\\Users\arvin76560\Documents\XPO\"
  
