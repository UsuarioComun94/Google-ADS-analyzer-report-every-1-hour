Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

logPath = "D:\Proyectos\hma-system\logs\hidden_task_runner.log"
Set logFile = fso.OpenTextFile(logPath, 8, True)
logFile.WriteLine Now & " | promote_pending_hidden_vbs_start"
logFile.Close

shell.CurrentDirectory = "D:\Proyectos\hma-system"
shell.Run """D:\Proyectos\hma-system\.venv\Scripts\pythonw.exe"" ""D:\Proyectos\hma-system\scripts\promote_hma_pending.py""", 0, True

