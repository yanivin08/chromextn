Dim oFSO
Set oFSO = CreateObject("Scripting.FileSystemObject")

' Create a new folder
  oFSO.CreateFolder "C:\\Users\arvin76560\Documents\XPO\Humit v14.08"

' Copy a file into the new folder
' Note that the destination folder path must end with a path separator (\)
  oFSO.CopyFolder "\\fs01\quality & process excellence\Process Excellence\Audit-QPE\TTL Philippines\Projects\XPO SpeedyG Buddy\HUMIT v14.08", "C:\\Users\arvin76560\Documents\XPO\"
