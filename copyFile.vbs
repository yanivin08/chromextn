Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

fso.CreateFolder "C:\\Users\arvin76560\Documents\XPO
fso.CreateFolder "C:\\Users\arvin76560\Documents\XPO\Humit v14.8"

fso.CopyFolder "\\fs01\quality & process excellence\Process Excellence\Audit-QPE\TTL Philippines\Projects\XPO SpeedyG Buddy\HUMIT v14.08", "C:\\Users\arvin76560\Documents\XPO\Humit v14.8"

