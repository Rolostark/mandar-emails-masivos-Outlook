Set WshShell = CreateObject("WScript.Shell")
' Ejecuta el archivo MailMergeStudio.exe en segundo plano de forma totalmente invisible
WshShell.Run "MailMergeStudio.exe", 0, false
