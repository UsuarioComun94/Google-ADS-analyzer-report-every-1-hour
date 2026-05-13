Set shell = CreateObject("WScript.Shell")
shell.CurrentDirectory = "D:\Proyectos\hma-system"
shell.Run """D:\Proyectos\hma-system\.venv\Scripts\pythonw.exe"" ""D:\Proyectos\hma-system\scripts\hma_error_monitor.py""", 0, False
