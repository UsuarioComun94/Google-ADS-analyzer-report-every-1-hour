$Repo = "UsuarioComun94/Google-ADS-analyzer-report-every-1-hour"
$WorkflowFile = "hma-hourly.yml"
$WorkflowName = "HMA Hourly Demo"
$BaseProjectDir = "D:\Proyectos\hma-system"
$BaseDownloadDir = Join-Path $BaseProjectDir "downloads"
$PythonExe = Join-Path $BaseProjectDir ".venv\Scripts\python.exe"
$UpdateMasterScript = Join-Path $BaseProjectDir "scripts\update_hma_master.py"
$MetricLogicScript = Join-Path $BaseProjectDir "scripts\fix_metric_comparison_logic.py"
$BuildCrossesScript = Join-Path $BaseProjectDir "scripts\build_metric_crosses.py"
$RecommendationsScript = Join-Path $BaseProjectDir "scripts\rebuild_recommendations_hourly.py"

# Inicio limpio del histÃƒÆ’Ã‚Â³rico local.
# El downloader ignorarÃƒÆ’Ã‚Â¡ cualquier run generado antes de esta fecha/hora local.
$StartFromLocal = "2026-05-12T19:30:00"
$StartFromDateTime = [datetime]::Parse($StartFromLocal)

# Cantidad de runs recientes que se revisan para recuperar backlog.
$LookbackRuns = 50

# IMPORTANTE:
# La PC local NO dispara workflows.
# La PC local solo descarga artifacts ya existentes y actualiza el histÃƒÆ’Ã‚Â³rico.
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
        Write-Host "No se pudo ajustar la fecha de modificaciÃƒÆ’Ã‚Â³n de: $TargetPath"
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
Write-Host "Modo local: SOLO DESCARGA. La generaciÃƒÆ’Ã‚Â³n horaria corresponde a GitHub Actions o trigger externo."

if ((Get-Date) -lt $StartFromDateTime) {
    Write-Host "TodavÃƒÆ’Ã‚Â­a no llegÃƒÆ’Ã‚Â³ la hora de inicio configurada."
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

# Procesar del mÃƒÆ’Ã‚Â¡s viejo al mÃƒÆ’Ã‚Â¡s nuevo para que el histÃƒÆ’Ã‚Â³rico quede natural.
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
        Write-Host "Fecha de carpeta ajustada a hora de generaciÃƒÆ’Ã‚Â³n: $RunCreatedAtLocal"

        $SeenHourKeys[$RunHourKey] = $true
        $SkippedCount += 1
        continue
    }

    if ($SeenHourKeys.ContainsKey($RunHourKey) -or (Has-DownloadedHour -HourKey $RunHourKey)) {
        Write-Host "Ya existe un artifact local para la hora lÃƒÆ’Ã‚Â³gica $RunHourKey. Se omite run duplicado:"
        Write-Host "Run: $RunId"
        Write-Host "Evento: $($Run.event)"
        Write-Host "Creado: $($Run.createdAt)"
        $SkippedDuplicateHourCount += 1
        continue
    }

    $ArtifactCount = Get-ArtifactCount -RunId $RunId

    if ($ArtifactCount -le 0) {
        Write-Host "Run $RunId no tiene artifacts descargables. Se omite, no es error crÃƒÆ’Ã‚Â­tico."
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
    Write-Host "Hora local de generaciÃƒÆ’Ã‚Â³n: $RunCreatedAtLocal"
    Write-Host "Hora lÃƒÆ’Ã‚Â³gica local: $RunHourKey"
    Write-Host "Artifacts disponibles: $ArtifactCount"
    Write-Host "Destino: $DownloadDir"

    gh run download $RunId `
      --repo $Repo `
      --dir $DownloadDir

    if ($LASTEXITCODE -eq 0) {
        Set-GeneratedTimestamp -TargetPath $DownloadDir -GeneratedAt $RunCreatedAtLocal

        Write-Host "Run $RunId descargado correctamente."
        Write-Host "Fecha de carpeta ajustada a hora de generaciÃƒÆ’Ã‚Â³n: $RunCreatedAtLocal"

        $SeenHourKeys[$RunHourKey] = $true
        $DownloadedCount += 1
    } else {
        Write-Host "FallÃƒÆ’Ã‚Â³ la descarga del run $RunId."
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

Write-Host "Resumen de recuperaciÃƒÆ’Ã‚Â³n:"
Write-Host "Runs elegibles revisados: $($Runs.Count)"
Write-Host "Runs nuevos descargados: $DownloadedCount"
Write-Host "Runs ya existentes omitidos: $SkippedCount"
Write-Host "Runs duplicados por hora omitidos: $SkippedDuplicateHourCount"
Write-Host "Runs sin artifact omitidos: $NoArtifactCount"
Write-Host "Runs con error real: $FailedCount"

if (Test-Path $PythonExe) {
    Write-Host "Actualizando HMA_Master.xlsx con reportes horarios ?nicos..."

    $MasterFile = "D:\Proyectos\hma-system\historico\HMA_Master.xlsx"
    $StyleScript = Join-Path $BaseProjectDir "scripts\style_hma_tabs_and_bold.py"
$EnsureCrossesScript = Join-Path $BaseProjectDir "scripts\ensure_metric_crosses_present.py"

    $MasterBefore = if (Test-Path $MasterFile) {
        (Get-Item $MasterFile).LastWriteTimeUtc
    } else {
        $null
    }

    & $PythonExe $UpdateMasterScript
    $UpdateExitCode = $LASTEXITCODE

    if ($UpdateExitCode -ne 0) {
        Write-Host "update_hma_master.py termin? con error. No se aplica l?gica ni estilo."
        exit $UpdateExitCode
    }

    $MasterAfter = if (Test-Path $MasterFile) {
        (Get-Item $MasterFile).LastWriteTimeUtc
    } else {
        $null
    }

    if ($MasterBefore -eq $MasterAfter) {
        Write-Host "Sin cambios en HMA_Master.xlsx; no se aplica l?gica ni estilo."
    } else {
        if (Test-Path $MetricLogicScript) {
            Write-Host "Corrigiendo l?gica de metric_comparison..."
            & $PythonExe $MetricLogicScript
            $MetricLogicExitCode = $LASTEXITCODE

            if ($MetricLogicExitCode -ne 0) {
                Write-Host "fix_metric_comparison_logic.py termin? con error."
                exit $MetricLogicExitCode
            }
        }
    if (Test-Path $BuildCrossesScript) {
        Write-Host "Construyendo metric_crosses..."
        & $PythonExe $BuildCrossesScript
        $BuildCrossesExitCode = $LASTEXITCODE

        if ($BuildCrossesExitCode -ne 0) {
            Write-Host "build_metric_crosses.py terminó con error."
            exit $BuildCrossesExitCode
        }
    } else {
        Write-Host "No se encontró build_metric_crosses.py. Frenar."
        exit 91
    }


        if (Test-Path $RecommendationsScript) {
            Write-Host "Construyendo recommendations de alto impacto..."
            & $PythonExe $RecommendationsScript
            $RecommendationsExitCode = $LASTEXITCODE

            if ($RecommendationsExitCode -ne 0) {
                Write-Host "build_high_impact_recommendations.py terminÃ³ con error."
                exit $RecommendationsExitCode
            }
        }

        if (Test-Path $StyleScript) {
            Write-Host "Aplicando estilo visual a HMA_Master.xlsx..."
            & $PythonExe $StyleScript
            $StyleExitCode = $LASTEXITCODE

            if ($StyleExitCode -ne 0) {
                Write-Host "style_hma_tabs_and_bold.py termin? con error."
                exit $StyleExitCode
            }
        }
    }
} else {
    Write-Host "No se encontr? Python del entorno virtual:"
    Write-Host $PythonExe
    exit 1
}

Write-Host "Verificando metric_crosses como guard final..."
if ((Test-Path $PythonExe) -and (Test-Path $EnsureCrossesScript)) {
    & $PythonExe $EnsureCrossesScript
    $EnsureExitCode = $LASTEXITCODE

    if ($EnsureExitCode -ne 0) {
        Write-Host "ensure_metric_crosses_present.py termin? con error."
        exit $EnsureExitCode
    }
} else {
    Write-Host "No se pudo ejecutar guard final de metric_crosses. Falta Python o ensure_metric_crosses_present.py."
}


Write-Host "Post-procesado analítico obligatorio..."
$BuildCrossesScript = Join-Path $BaseProjectDir "scripts\build_metric_crosses.py"
$RecommendationsScript = Join-Path $BaseProjectDir "scripts\rebuild_recommendations_hourly.py"
$StyleScript = Join-Path $BaseProjectDir "scripts\style_hma_tabs_and_bold.py"
$FormatRecommendationsScript = Join-Path $BaseProjectDir "scripts\format_recommendations_final.py"
$EnsureCrossesScript = Join-Path $BaseProjectDir "scripts\ensure_metric_crosses_present.py"

if ((Test-Path $PythonExe) -and (Test-Path $BuildCrossesScript)) {
    & $PythonExe $BuildCrossesScript
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

if ((Test-Path $PythonExe) -and (Test-Path $RecommendationsScript)) {
    & $PythonExe $RecommendationsScript
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

if ((Test-Path $PythonExe) -and (Test-Path $StyleScript)) {
    & $PythonExe $StyleScript
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

if ((Test-Path $PythonExe) -and (Test-Path $FormatRecommendationsScript)) {
    & $PythonExe $FormatRecommendationsScript
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

if ((Test-Path $PythonExe) -and (Test-Path $EnsureCrossesScript)) {
    & $PythonExe $EnsureCrossesScript
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

Write-Host "Post-procesado analítico obligatorio terminado."

