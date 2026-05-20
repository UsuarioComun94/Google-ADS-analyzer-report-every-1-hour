Set WshShell = CreateObject("WScript.Shell")
baseDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
ps1 = baseDir & "\scripts\hma_manager_gui_v2.ps1"
cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & ps1 & """"
WshShell.Run cmd, 0, False
