Set WshShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

BaseDir = FSO.GetParentFolderName(WScript.ScriptFullName)
BatPath = BaseDir & "\hma_desktop_start.bat"

WshShell.Run """" & BatPath & """", 0, False
