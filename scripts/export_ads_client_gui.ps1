Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "Stop"

$BaseDir = Split-Path -Parent $PSScriptRoot
$Python = Join-Path $BaseDir ".venv\Scripts\python.exe"
$ClientsRoot = Join-Path $BaseDir "clientes"
$GoogleExporter = Join-Path $BaseDir "scripts\export_google_ads_client.py"
$MetaExporter = Join-Path $BaseDir "scripts\export_meta_ads_client.py"

function Show-Info($msg) {
    [System.Windows.Forms.MessageBox]::Show($msg, "HMA Export Ads", "OK", "Information") | Out-Null
}

function Show-ErrorBox($msg) {
    [System.Windows.Forms.MessageBox]::Show($msg, "HMA Export Ads - Error", "OK", "Error") | Out-Null
}

function Get-HmaClients {
    $clients = @()

    Get-ChildItem $ClientsRoot -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -ne "_template" } |
    ForEach-Object {
        $configPath = Join-Path $_.FullName "config\client_config.json"

        if (Test-Path $configPath) {
            try {
                $cfg = Get-Content $configPath -Raw | ConvertFrom-Json
                $clients += [PSCustomObject]@{
                    Id = $cfg.client_id
                    Name = $cfg.client_name
                    Path = $_.FullName
                    Label = "$($cfg.client_name) [$($cfg.client_id)]"
                    GoogleEnabled = [bool]$cfg.platforms.google_ads.enabled
                    MetaEnabled = [bool]$cfg.platforms.meta_ads.enabled
                }
            } catch {}
        }
    }

    return $clients
}

function Run-Exporter($script, $clientPath, $extraArgs) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $Python
    $psi.Arguments = "`"$script`" --client-dir `"$clientPath`" $extraArgs"
    $psi.WorkingDirectory = $BaseDir
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true

    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi
    $p.Start() | Out-Null
    $stdout = $p.StandardOutput.ReadToEnd()
    $stderr = $p.StandardError.ReadToEnd()
    $p.WaitForExit()

    return @{
        ExitCode = $p.ExitCode
        StdOut = $stdout
        StdErr = $stderr
    }
}

try {
    if (-not (Test-Path $Python)) {
        Show-ErrorBox "No existe Python:`n$Python"
        exit 1
    }

    $clients = @(Get-HmaClients)

    if ($clients.Count -eq 0) {
        Show-ErrorBox "No hay clientes creados."
        exit 1
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Exportar Ads por cliente"
    $form.Size = New-Object System.Drawing.Size(560,300)
    $form.StartPosition = "CenterScreen"

    $clientLabel = New-Object System.Windows.Forms.Label
    $clientLabel.Text = "Cliente:"
    $clientLabel.Location = New-Object System.Drawing.Point(20,25)
    $clientLabel.Size = New-Object System.Drawing.Size(500,22)
    $form.Controls.Add($clientLabel)

    $clientCombo = New-Object System.Windows.Forms.ComboBox
    $clientCombo.Location = New-Object System.Drawing.Point(20,50)
    $clientCombo.Size = New-Object System.Drawing.Size(500,25)
    $clientCombo.DropDownStyle = "DropDownList"
    foreach ($c in $clients) { [void]$clientCombo.Items.Add($c.Label) }
    $clientCombo.SelectedIndex = 0
    $form.Controls.Add($clientCombo)

    $platformLabel = New-Object System.Windows.Forms.Label
    $platformLabel.Text = "Plataforma:"
    $platformLabel.Location = New-Object System.Drawing.Point(20,90)
    $platformLabel.Size = New-Object System.Drawing.Size(500,22)
    $form.Controls.Add($platformLabel)

    $platformCombo = New-Object System.Windows.Forms.ComboBox
    $platformCombo.Location = New-Object System.Drawing.Point(20,115)
    $platformCombo.Size = New-Object System.Drawing.Size(500,25)
    $platformCombo.DropDownStyle = "DropDownList"
    [void]$platformCombo.Items.Add("Google Ads")
    [void]$platformCombo.Items.Add("Meta Ads")
    $platformCombo.SelectedIndex = 0
    $form.Controls.Add($platformCombo)

    $periodLabel = New-Object System.Windows.Forms.Label
    $periodLabel.Text = "Período:"
    $periodLabel.Location = New-Object System.Drawing.Point(20,155)
    $periodLabel.Size = New-Object System.Drawing.Size(500,22)
    $form.Controls.Add($periodLabel)

    $periodCombo = New-Object System.Windows.Forms.ComboBox
    $periodCombo.Location = New-Object System.Drawing.Point(20,180)
    $periodCombo.Size = New-Object System.Drawing.Size(500,25)
    $periodCombo.DropDownStyle = "DropDownList"
    [void]$periodCombo.Items.Add("today")
    [void]$periodCombo.Items.Add("yesterday")
    [void]$periodCombo.Items.Add("last_7d")
    $periodCombo.SelectedIndex = 0
    $form.Controls.Add($periodCombo)

    $exportButton = New-Object System.Windows.Forms.Button
    $exportButton.Text = "Exportar CSV"
    $exportButton.Location = New-Object System.Drawing.Point(300,225)
    $exportButton.Size = New-Object System.Drawing.Size(105,30)
    $form.Controls.Add($exportButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancelar"
    $cancelButton.Location = New-Object System.Drawing.Point(415,225)
    $cancelButton.Size = New-Object System.Drawing.Size(105,30)
    $cancelButton.Add_Click({ $form.Close() })
    $form.Controls.Add($cancelButton)

    $exportButton.Add_Click({
        try {
            $client = $clients[$clientCombo.SelectedIndex]
            $platform = $platformCombo.SelectedItem.ToString()
            $period = $periodCombo.SelectedItem.ToString()

            if ($platform -eq "Google Ads") {
                if (-not $client.GoogleEnabled) {
                    Show-ErrorBox "Google Ads no está conectado para este cliente."
                    return
                }

                $dateRange = switch ($period) {
                    "today" { "TODAY" }
                    "yesterday" { "YESTERDAY" }
                    "last_7d" { "LAST_7_DAYS" }
                }

                $result = Run-Exporter $GoogleExporter $client.Path "--date-range $dateRange"
            }

            if ($platform -eq "Meta Ads") {
                if (-not $client.MetaEnabled) {
                    Show-ErrorBox "Meta Ads no está conectado para este cliente."
                    return
                }

                $datePreset = switch ($period) {
                    "today" { "today" }
                    "yesterday" { "yesterday" }
                    "last_7d" { "last_7d" }
                }

                $result = Run-Exporter $MetaExporter $client.Path "--date-preset $datePreset"
            }

            if ($result.ExitCode -eq 0) {
                Show-Info "Export OK.`n`n$($result.StdOut)"
            } else {
                Show-ErrorBox "Falló export.`n`nSTDOUT:`n$($result.StdOut)`n`nSTDERR:`n$($result.StdErr)"
            }
        } catch {
            Show-ErrorBox $_.Exception.Message
        }
    })

    [void]$form.ShowDialog()

} catch {
    Show-ErrorBox $_.Exception.Message
    exit 1
}
