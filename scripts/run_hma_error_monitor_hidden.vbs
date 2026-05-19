Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

logPath = "D:\Proyectos\hma-system\logs\hidden_task_runner.log"
Set logFile = fso.OpenTextFile(logPath, 8, True)
logFile.WriteLine Now & " | error_monitor_hidden_vbs_start"
logFile.Close

shell.CurrentDirectory = "D:\Proyectos\hma-system"
shell.Run """D:\Proyectos\hma-system\.venv\Scripts\pythonw.exe"" ""D:\Proyectos\hma-system\scripts\hma_error_monitor.py""", 0, True

