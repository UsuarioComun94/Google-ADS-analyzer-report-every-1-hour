$Repo = "UsuarioComun94/Google-ADS-analyzer-report-every-1-hour"
$WorkflowFile = "hma-hourly.yml"
$WorkflowName = "HMA Hourly Demo"
$BaseProjectDir = "D:\Proyectos\hma-system"
$BaseDownloadDir = Join-Path $BaseProjectDir "downloads"
$PythonExe = Join-Path $BaseProjectDir ".venv\Scripts\python.exe"
$UpdateMasterScript = Join-Path $BaseProjectDir "scripts\update_hma_master.py"

# Cantidad de runs recientes que se revisan para recuperar backlog.
$LookbackRuns = 50

# Si GitHub schedule no generó un run en la hora actual,
# este script dispara workflow_dispatch para crear el reporte horario.
$TriggerWorkflowIfMissingCurrentHour = $true

New-Item -ItemType Directory -Force -Path $BaseDownloadDir | Out-Null

function Get-RecentRuns {
    param(
        [int]$Limit = 50
    )

    $runsJson = gh run list `
      --repo $Repo `
      --workflow $WorkflowName `
      --limit $Limit `
      --json databaseId,createdAt,event,status,conclusion,displayTitle

    if (-not $runsJson) {
        return @()
    }

    $runs = $runsJson | ConvertFrom-Json

    if (-not $runs) {
        return @()
    }

    if ($runs -isnot [System.Array]) {
        return @($runs)
    }

    return $runs
}

function Has-CurrentHourRun {
    param(
        [array]$Runs
    )

    $now = Get-Date
    $currentHourStart = Get-Date -Year $now.Year -Month $now.Month -Day $now.Day -Hour $now.Hour -Minute 0 -Second 0

    foreach ($run in $Runs) {
        try {
            $createdLocal = ([datetime]::Parse($run.createdAt)).ToLocalTime()
        } catch {
            continue
        }

        if ($createdLocal -ge $currentHourStart) {
            if (
                $run.status -eq "queued" -or
                $run.status -eq "in_progress" -or
                ($run.status -eq "completed" -and $run.conclusion -eq "success")
            ) {
                return $true
            }
        }
    }

    return $false
}

function Get-ArtifactCount {
    param(
        [string]$RunId
    )

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


function Set-GeneratedTimestamp {
    param(
        [string]$TargetPath,
        [datetime]$GeneratedAt
    )

    try {
        if (-not (Test-Path -LiteralPath $TargetPath)) {
            return
        }

        # Primero ajustar archivos y subcarpetas internas.
        # Después ajustar la carpeta principal para que Windows Explorer muestre
        # la hora de generación del run, no la hora local de descarga.
        Get-ChildItem -LiteralPath $TargetPath -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                $_.CreationTime = $GeneratedAt
                $_.LastWriteTime = $GeneratedAt
                $_.LastAccessTime = $GeneratedAt
            } catch {
                # No bloquear el flujo por metadatos de archivo.
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

function Trigger-Workflow-And-Wait {
    Write-Host "No hay run exitoso/en curso para la hora actual."
    Write-Host "Disparando GitHub Actions manualmente para generar reporte horario..."

    $beforeRuns = Get-RecentRuns -Limit 5
    $beforeIds = @{}
    foreach ($run in $beforeRuns) {
        $beforeIds[[string]$run.databaseId] = $true
    }

    gh workflow run $WorkflowFile `
      --repo $Repo `
      --ref main

    if ($LASTEXITCODE -ne 0) {
        Write-Host "No se pudo disparar el workflow."
        return $null
    }

    $newRunId = $null

    for ($i = 0; $i -lt 24; $i++) {
        Start-Sleep -Seconds 5

        $afterRuns = Get-RecentRuns -Limit 10

        foreach ($run in $afterRuns) {
            $candidateId = [string]$run.databaseId

            if (-not $beforeIds.ContainsKey($candidateId)) {
                if ($run.event -eq "workflow_dispatch") {
                    $newRunId = $candidateId
                    break
                }
            }
        }

        if ($newRunId) {
            break
        }

        Write-Host "Esperando que GitHub cree el run nuevo..."
    }

    if (-not $newRunId) {
        Write-Host "No se pudo identificar el nuevo run. Se continuará con descarga de backlog."
        return $null
    }

    Write-Host "Run nuevo detectado: $newRunId"
    Write-Host "Esperando finalización del run..."

    gh run watch $newRunId `
      --repo $Repo `
      --exit-status `
      --interval 5

    if ($LASTEXITCODE -eq 0) {
        Write-Host "Run $newRunId finalizado correctamente."
    } else {
        Write-Host "El run $newRunId no finalizó correctamente. Se continuará con descarga de runs exitosos disponibles."
    }

    return $newRunId
}

Write-Host "Buscando runs recientes en GitHub Actions..."

$InitialRuns = Get-RecentRuns -Limit $LookbackRuns

if ($TriggerWorkflowIfMissingCurrentHour) {
    $hasCurrentHourRun = Has-CurrentHourRun -Runs $InitialRuns

    if ($hasCurrentHourRun) {
        Write-Host "Ya existe un run exitoso/en curso para la hora actual. No se dispara otro."
    } else {
        Trigger-Workflow-And-Wait | Out-Null
    }
}

Write-Host "Buscando runs exitosos recientes para recuperar backlog..."

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

# Procesar del más viejo al más nuevo para que el histórico quede natural.
$Runs = $Runs | Sort-Object { [datetime]$_.createdAt }

$DownloadedCount = 0
$SkippedCount = 0
$NoArtifactCount = 0
$FailedCount = 0

foreach ($Run in $Runs) {
    $RunId = [string]$Run.databaseId

    if (-not $RunId) {
        continue
    }

    $ExistingRunFolder = Get-ChildItem -Path $BaseDownloadDir -Directory -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "*_run-$RunId" } |
        Sort-Object FullName |
        Select-Object -First 1

    if ($ExistingRunFolder) {
        $RunCreatedAtLocal = ([datetime]::Parse($Run.createdAt)).ToLocalTime()
        Set-GeneratedTimestamp -TargetPath $ExistingRunFolder.FullName -GeneratedAt $RunCreatedAtLocal

        Write-Host "Ya existe run $RunId. No se duplica:"
        Write-Host $ExistingRunFolder.FullName
        Write-Host "Fecha de carpeta ajustada a hora de generación: $RunCreatedAtLocal"
        $SkippedCount += 1
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

    $RunCreatedAtLocal = ([datetime]::Parse($Run.createdAt)).ToLocalTime()
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
    Write-Host "Artifacts disponibles: $ArtifactCount"
    Write-Host "Destino: $DownloadDir"

    gh run download $RunId `
      --repo $Repo `
      --dir $DownloadDir

    if ($LASTEXITCODE -eq 0) {
        Set-GeneratedTimestamp -TargetPath $DownloadDir -GeneratedAt $RunCreatedAtLocal

        Write-Host "Run $RunId descargado correctamente."
        Write-Host "Fecha de carpeta ajustada a hora de generación: $RunCreatedAtLocal"
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
Write-Host "Runs revisados: $($Runs.Count)"
Write-Host "Runs nuevos descargados: $DownloadedCount"
Write-Host "Runs ya existentes omitidos: $SkippedCount"
Write-Host "Runs sin artifact omitidos: $NoArtifactCount"
Write-Host "Runs con error real: $FailedCount"

# El Excel se reconstruye desde reportes horarios únicos.
# No genera reportes nuevos. Solo consolida lo que ya existe en downloads/.
if (Test-Path $PythonExe) {
    Write-Host "Actualizando HMA_Master.xlsx con reportes horarios únicos..."
    & $PythonExe $UpdateMasterScript
} else {
    Write-Host "No se encontró Python del entorno virtual:"
    Write-Host $PythonExe
    exit 1
}
