Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

' จับที่อยู่โฟลเดอร์ปัจจุบันแบบเต็มรูปแบบ (E:\AUTOEXCEL)
currentFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)
WshShell.CurrentDirectory = currentFolder

' 1. สั่งรัน Python Server แบบซ่อนหน้าต่าง
WshShell.Run "cmd /c python main.py", 0, False

' หน่วงเวลา 2 วินาทีให้เซิร์ฟเวอร์เตรียมตัว
WScript.Sleep 2000

' 2. สั่งเปิดไฟล์ Excel ด้วย Path แบบเต็ม (แก้ไขชื่อไฟล์ตรงนี้ให้ตรงกับที่คุณตั้ง)
excelFilePath = currentFolder & "\Report.xlsx"
WshShell.Run """" & excelFilePath & """", 1, False