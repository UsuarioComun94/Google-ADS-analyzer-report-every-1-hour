Set shell = CreateObject("WScript.Shell")
shell.CurrentDirectory = "D:\Proyectos\hma-system"
shell.Run """D:\Proyectos\hma-system\.venv\Scripts\pythonw.exe"" ""D:\Proyectos\hma-system\scripts\promote_hma_pending.py""", 0, False
