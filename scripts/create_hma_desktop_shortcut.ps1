$ErrorActionPreference = "Stop"

$BaseDir = "D:\Proyectos\hma-system"
$VbsPath = Join-Path $BaseDir "hma_manager.vbs"
$ShortcutPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "HMA Manager.lnk"

if (-not (Test-Path $VbsPath)) {
    [System.Windows.Forms.MessageBox]::Show("No existe hma_manager.vbs en $BaseDir", "HMA Shortcut", "OK", "Error") | Out-Null
    exit 1
}

$Shell = New-Object -ComObject WScript.Shell
$Shortcut = $Shell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = "wscript.exe"
$Shortcut.Arguments = "`"$VbsPath`""
$Shortcut.WorkingDirectory = $BaseDir
$Shortcut.WindowStyle = 7
$Shortcut.Description = "Abrir HMA Manager sin consola"
$Shortcut.Save()

Write-Host "SHORTCUT_OK: $ShortcutPath"
