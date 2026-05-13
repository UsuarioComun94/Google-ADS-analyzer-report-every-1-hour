Set shell = CreateObject("WScript.Shell")
shell.CurrentDirectory = "D:\Proyectos\hma-system"
shell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -File ""D:\Proyectos\hma-system\download_latest_hma_artifact.ps1""", 0, False
