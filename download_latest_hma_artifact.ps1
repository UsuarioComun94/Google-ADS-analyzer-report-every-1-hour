$Repo = "UsuarioComun94/Google-ADS-analyzer-report-every-1-hour"
$WorkflowName = "HMA Hourly Demo"
$BaseProjectDir = "D:\Proyectos\hma-system"
$BaseDownloadDir = Join-Path $BaseProjectDir "downloads"
$PythonExe = Join-Path $BaseProjectDir ".venv\Scripts\python.exe"
$UpdateMasterScript = Join-Path $BaseProjectDir "scripts\update_hma_master.py"

New-Item -ItemType Directory -Force -Path $BaseDownloadDir | Out-Null

$RunId = gh run list `
  --repo $Repo `
  --workflow $WorkflowName `
  --status success `
  --limit 1 `
  --json databaseId `
  --jq ".[0].databaseId"

if (-not $RunId) {
    Write-Host "No se encontró una ejecución exitosa para descargar."
    exit 1
}

$Today = Get-Date -Format "yyyy-MM-dd"
$TodayDir = Join-Path $BaseDownloadDir $Today

New-Item -ItemType Directory -Force -Path $TodayDir | Out-Null

$TimeStamp = Get-Date -Format "HH-mm"
$DownloadDirName = "${TimeStamp}_run-${RunId}"
$DownloadDir = Join-Path $TodayDir $DownloadDirName

if (Test-Path $DownloadDir) {
    Write-Host "Este run ya fue descargado anteriormente:"
    Write-Host $DownloadDir
} else {
    New-Item -ItemType Directory -Force -Path $DownloadDir | Out-Null

    gh run download $RunId `
      --repo $Repo `
      --dir $DownloadDir

    Write-Host "Artifact descargado correctamente."
    Write-Host "Carpeta del día: $TodayDir"
    Write-Host "Carpeta de esta descarga: $DownloadDir"
}

if (Test-Path $PythonExe) {
    Write-Host "Actualizando HMA_Master.xlsx..."
    & $PythonExe $UpdateMasterScript
} else {
    Write-Host "No se encontró Python del entorno virtual:"
    Write-Host $PythonExe
    exit 1
}
