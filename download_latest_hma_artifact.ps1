$Repo = "UsuarioComun94/Google-ADS-analyzer-report-every-1-hour"
$WorkflowFile = "hma-hourly.yml"
$WorkflowName = "HMA Hourly Demo"
$BaseProjectDir = "D:\Proyectos\hma-system"
$BaseDownloadDir = Join-Path $BaseProjectDir "downloads"
$PythonExe = Join-Path $BaseProjectDir ".venv\Scripts\python.exe"
$UpdateMasterScript = Join-Path $BaseProjectDir "scripts\update_hma_master.py"

# Inicio limpio del histórico local.
# El downloader ignorará cualquier run generado antes de esta fecha/hora local.
$StartFromLocal = "2026-05-12T19:30:00"
$StartFromDateTime = [datetime]::Parse($StartFromLocal)

# Cantidad de runs recientes que se revisan para recuperar backlog.
$LookbackRuns = 50

# IMPORTANTE:
# La PC local NO dispara workflows.
# La PC local solo descarga artifacts ya existentes y actualiza el histórico.
$TriggerWorkflowIfMissingCurrentHour = $false

New-Item -ItemType Directory -Force -Path $BaseDownloadDir | Out-Null

function Convert-RunCreatedAtToLocal {
    param([object]$Run)
    return ([datetime]::Parse($Run.createdAt)).ToLocalTime()
}

function Get-RunHourKey {
    param([object]$Run)

    $createdLocal = Convert-RunCreatedAtToLocal -Run $Run
    return $createdLocal.ToString("yyyy-MM-dd HH:00")
}

function Is-RunEligible {
    param([object]$Run)

    try {
        $createdLocal = Convert-RunCreatedAtToLocal -Run $Run
        return ($createdLocal -ge $StartFromDateTime)
    } catch {
        return $false
    }
}

function Set-GeneratedTimestamp {
    param(
        [string]$TargetPath,
        [datetime]$GeneratedAt
    )

    try {
        if (-not (Test-Path -LiteralPath $TargetPath)) {
            return
        }

        Get-ChildItem -LiteralPath $TargetPath -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                $_.CreationTime = $GeneratedAt
                $_.LastWriteTime = $GeneratedAt
                $_.LastAccessTime = $GeneratedAt
            } catch {
                # No bloquear el flujo por metadatos.
            }
        }

        $targetItem = Get-Item -LiteralPath $TargetPath -Force
        $targetItem.CreationTime = $GeneratedAt
        $targetItem.LastWriteTime = $GeneratedAt
        $targetItem.LastAccessTime = $GeneratedAt
    } catch {
        Write-Host "No se pudo ajustar la fecha de modificación de: $TargetPath"
    }
}

function Get-ArtifactCount {
    param([string]$RunId)

    try {
        $countText = gh api "repos/$Repo/actions/runs/$RunId/artifacts" --jq ".total_count" 2>$null

        if (-not $countText) {
            return 0
        }

        return [int]$countText
    } catch {
        return 0
    }
}

function Has-DownloadedHour {
    param([string]$HourKey)

    try {
        $folders = Get-ChildItem -Path $BaseDownloadDir -Directory -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like "*_run-*" }

        foreach ($folder in $folders) {
            $parentDate = Split-Path (Split-Path $folder.FullName -Parent) -Leaf

            if ($folder.Name -match "^(\d{2})-(\d{2})_run-\d+$") {
                $hourFromFolder = "$parentDate $($Matches[1]):00"

                if ($hourFromFolder -eq $HourKey) {
                    return $true
                }
            }
        }

        return $false
    } catch {
        return $false
    }
}

Write-Host "Inicio limpio configurado desde: $StartFromDateTime"
Write-Host "Modo local: SOLO DESCARGA. La generación horaria corresponde a GitHub Actions o trigger externo."

if ((Get-Date) -lt $StartFromDateTime) {
    Write-Host "Todavía no llegó la hora de inicio configurada."
    Write-Host "No se descargan runs anteriores."
    exit 0
}

Write-Host "Buscando runs exitosos recientes para recuperar backlog elegible..."

$RunsJson = gh run list `
  --repo $Repo `
  --workflow $WorkflowName `
  --status success `
  --limit $LookbackRuns `
  --json databaseId,createdAt,event,displayTitle,status,conclusion

if (-not $RunsJson) {
    Write-Host "No se encontraron runs exitosos para revisar."
    exit 1
}

$Runs = $RunsJson | ConvertFrom-Json

if (-not $Runs -or $Runs.Count -eq 0) {
    Write-Host "No se encontraron runs exitosos para descargar."
    exit 1
}

$Runs = @($Runs | Where-Object { Is-RunEligible -Run $_ })

if (-not $Runs -or $Runs.Count -eq 0) {
    Write-Host "No hay runs elegibles desde $StartFromDateTime."
    Write-Host "No se actualiza HMA_Master.xlsx porque no hay datos nuevos."
    exit 0
}

# Procesar del más viejo al más nuevo para que el histórico quede natural.
$Runs = $Runs | Sort-Object { [datetime]$_.createdAt }

$DownloadedCount = 0
$SkippedCount = 0
$SkippedDuplicateHourCount = 0
$NoArtifactCount = 0
$FailedCount = 0
$SeenHourKeys = @{}

foreach ($Run in $Runs) {
    $RunId = [string]$Run.databaseId

    if (-not $RunId) {
        continue
    }

    $RunCreatedAtLocal = Convert-RunCreatedAtToLocal -Run $Run
    $RunHourKey = Get-RunHourKey -Run $Run

    $ExistingRunFolder = Get-ChildItem -Path $BaseDownloadDir -Directory -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "*_run-$RunId" } |
        Sort-Object FullName |
        Select-Object -First 1

    if ($ExistingRunFolder) {
        Set-GeneratedTimestamp -TargetPath $ExistingRunFolder.FullName -GeneratedAt $RunCreatedAtLocal

        Write-Host "Ya existe run $RunId. No se duplica:"
        Write-Host $ExistingRunFolder.FullName
        Write-Host "Fecha de carpeta ajustada a hora de generación: $RunCreatedAtLocal"

        $SeenHourKeys[$RunHourKey] = $true
        $SkippedCount += 1
        continue
    }

    if ($SeenHourKeys.ContainsKey($RunHourKey) -or (Has-DownloadedHour -HourKey $RunHourKey)) {
        Write-Host "Ya existe un artifact local para la hora lógica $RunHourKey. Se omite run duplicado:"
        Write-Host "Run: $RunId"
        Write-Host "Evento: $($Run.event)"
        Write-Host "Creado: $($Run.createdAt)"
        $SkippedDuplicateHourCount += 1
        continue
    }

    $ArtifactCount = Get-ArtifactCount -RunId $RunId

    if ($ArtifactCount -le 0) {
        Write-Host "Run $RunId no tiene artifacts descargables. Se omite, no es error crítico."
        Write-Host "Evento: $($Run.event)"
        Write-Host "Creado: $($Run.createdAt)"
        $NoArtifactCount += 1
        continue
    }

    $RunDay = $RunCreatedAtLocal.ToString("yyyy-MM-dd")
    $RunTime = $RunCreatedAtLocal.ToString("HH-mm")

    $TodayDir = Join-Path $BaseDownloadDir $RunDay
    New-Item -ItemType Directory -Force -Path $TodayDir | Out-Null

    $DownloadDirName = "${RunTime}_run-${RunId}"
    $DownloadDir = Join-Path $TodayDir $DownloadDirName

    New-Item -ItemType Directory -Force -Path $DownloadDir | Out-Null

    Write-Host "Descargando run pendiente $RunId..."
    Write-Host "Evento: $($Run.event)"
    Write-Host "Creado: $($Run.createdAt)"
    Write-Host "Hora local de generación: $RunCreatedAtLocal"
    Write-Host "Hora lógica local: $RunHourKey"
    Write-Host "Artifacts disponibles: $ArtifactCount"
    Write-Host "Destino: $DownloadDir"

    gh run download $RunId `
      --repo $Repo `
      --dir $DownloadDir

    if ($LASTEXITCODE -eq 0) {
        Set-GeneratedTimestamp -TargetPath $DownloadDir -GeneratedAt $RunCreatedAtLocal

        Write-Host "Run $RunId descargado correctamente."
        Write-Host "Fecha de carpeta ajustada a hora de generación: $RunCreatedAtLocal"

        $SeenHourKeys[$RunHourKey] = $true
        $DownloadedCount += 1
    } else {
        Write-Host "Falló la descarga del run $RunId."
        $FailedCount += 1

        try {
            if ((Test-Path $DownloadDir) -and -not (Get-ChildItem $DownloadDir -Recurse -ErrorAction SilentlyContinue)) {
                Remove-Item $DownloadDir -Force -Recurse
            }
        } catch {
            Write-Host "No se pudo limpiar la carpeta fallida: $DownloadDir"
        }
    }
}

Write-Host "Resumen de recuperación:"
Write-Host "Runs elegibles revisados: $($Runs.Count)"
Write-Host "Runs nuevos descargados: $DownloadedCount"
Write-Host "Runs ya existentes omitidos: $SkippedCount"
Write-Host "Runs duplicados por hora omitidos: $SkippedDuplicateHourCount"
Write-Host "Runs sin artifact omitidos: $NoArtifactCount"
Write-Host "Runs con error real: $FailedCount"

if (Test-Path $PythonExe) {
    Write-Host "Actualizando HMA_Master.xlsx con reportes horarios únicos..."
    & $PythonExe $UpdateMasterScript
} else {
    Write-Host "No se encontró Python del entorno virtual:"
    Write-Host $PythonExe
    exit 1
}
