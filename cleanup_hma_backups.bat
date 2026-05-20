@echo off
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\cleanup_hma_backups.ps1" -KeepLast 8
