Set WshShell = CreateObject("WScript.Shell")
' Mendapatkan lokasi folder tempat script berada
strPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
WshShell.CurrentDirectory = strPath

' Memanggil file bat yang sudah berhasil tadi tanpa memunculkan jendela hitam (0)
WshShell.Run chr(34) & "run_kurasi.bat" & Chr(34), 0
Set WshShell = Nothing