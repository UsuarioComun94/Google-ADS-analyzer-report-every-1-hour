Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

logPath = "D:\Proyectos\hma-system\logs\hidden_task_runner.log"
Set logFile = fso.OpenTextFile(logPath, 8, True)
logFile.WriteLine Now & " | downloader_hidden_vbs_start"
logFile.Close

shell.CurrentDirectory = "D:\Proyectos\hma-system"
shell.Run "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File ""D:\Proyectos\hma-system\download_latest_hma_artifact.ps1""", 0, True

