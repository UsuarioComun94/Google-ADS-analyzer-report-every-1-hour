Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
$ClientsRoot = Join-Path $ProjectRoot "clientes"
$CreateScript = Join-Path $PSScriptRoot "create_hma_client.ps1"

function Show-Info($msg) {
    [System.Windows.Forms.MessageBox]::Show($msg, "HMA Crear Cliente", "OK", "Information") | Out-Null
}

function Show-ErrorBox($msg) {
    [System.Windows.Forms.MessageBox]::Show($msg, "HMA Crear Cliente - Error", "OK", "Error") | Out-Null
}

function Get-NextClientId {
    New-Item -ItemType Directory -Force -Path $ClientsRoot | Out-Null

    $max = 0

    Get-ChildItem $ClientsRoot -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -ne "_template" } |
    ForEach-Object {
        if ($_.Name -match "^CL-(\d{3})-") {
            $n = [int]$matches[1]
            if ($n -gt $max) { $max = $n }
        } else {
            $cfgPath = Join-Path $_.FullName "config\client_config.json"
            if (Test-Path $cfgPath) {
                try {
                    $cfg = Get-Content $cfgPath -Raw | ConvertFrom-Json
                    if ($cfg.client_id -match "^CL-(\d{3})$") {
                        $n = [int]$matches[1]
                        if ($n -gt $max) { $max = $n }
                    }
                } catch {}
            }
        }
    }

    return "CL-{0:D3}" -f ($max + 1)
}

function ConvertTo-SafeFolderName {
    param([string]$Name)

    $safe = $Name.Trim()
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()

    foreach ($ch in $invalid) {
        $safe = $safe.Replace($ch, "-")
    }

    $safe = $safe -replace "\s+", " "
    $safe = $safe.Trim(" .-_")

    return $safe
}

try {
    if (-not (Test-Path $CreateScript)) {
        Show-ErrorBox "No existe:`n$CreateScript"
        exit 1
    }

    $clientId = Get-NextClientId

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Crear nuevo cliente HMA"
    $form.Size = New-Object System.Drawing.Size(620,330)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false

    $title = New-Object System.Windows.Forms.Label
    $title.Text = "Nuevo cliente HMA"
    $title.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
    $title.Location = New-Object System.Drawing.Point(20,20)
    $title.Size = New-Object System.Drawing.Size(560,30)
    $form.Controls.Add($title)

    $idLabel = New-Object System.Windows.Forms.Label
    $idLabel.Text = "ID asignado automáticamente:"
    $idLabel.Location = New-Object System.Drawing.Point(20,65)
    $idLabel.Size = New-Object System.Drawing.Size(220,22)
    $form.Controls.Add($idLabel)

    $idBox = New-Object System.Windows.Forms.TextBox
    $idBox.Text = $clientId
    $idBox.Location = New-Object System.Drawing.Point(250,62)
    $idBox.Size = New-Object System.Drawing.Size(320,24)
    $idBox.ReadOnly = $true
    $form.Controls.Add($idBox)

    $nameLabel = New-Object System.Windows.Forms.Label
    $nameLabel.Text = "Nombre del cliente:"
    $nameLabel.Location = New-Object System.Drawing.Point(20,105)
    $nameLabel.Size = New-Object System.Drawing.Size(220,22)
    $form.Controls.Add($nameLabel)

    $nameBox = New-Object System.Windows.Forms.TextBox
    $nameBox.Location = New-Object System.Drawing.Point(250,102)
    $nameBox.Size = New-Object System.Drawing.Size(320,24)
    $form.Controls.Add($nameBox)

    $previewLabel = New-Object System.Windows.Forms.Label
    $previewLabel.Text = "Carpeta que se va a crear:"
    $previewLabel.Location = New-Object System.Drawing.Point(20,145)
    $previewLabel.Size = New-Object System.Drawing.Size(220,22)
    $form.Controls.Add($previewLabel)

    $previewBox = New-Object System.Windows.Forms.TextBox
    $previewBox.Location = New-Object System.Drawing.Point(20,170)
    $previewBox.Size = New-Object System.Drawing.Size(550,24)
    $previewBox.ReadOnly = $true
    $form.Controls.Add($previewBox)

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = "Escribí el nombre del cliente. El ID se calcula solo."
    $statusLabel.Location = New-Object System.Drawing.Point(20,210)
    $statusLabel.Size = New-Object System.Drawing.Size(550,28)
    $form.Controls.Add($statusLabel)

    function Update-Preview {
        $safeName = ConvertTo-SafeFolderName $nameBox.Text
        if ([string]::IsNullOrWhiteSpace($safeName)) {
            $previewBox.Text = ""
        } else {
            $previewBox.Text = Join-Path $ClientsRoot "$clientId-$safeName"
        }
    }

    $nameBox.Add_TextChanged({ Update-Preview })

    $createButton = New-Object System.Windows.Forms.Button
    $createButton.Text = "Crear cliente"
    $createButton.Location = New-Object System.Drawing.Point(350,245)
    $createButton.Size = New-Object System.Drawing.Size(105,32)
    $form.Controls.Add($createButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancelar"
    $cancelButton.Location = New-Object System.Drawing.Point(465,245)
    $cancelButton.Size = New-Object System.Drawing.Size(105,32)
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancelButton)

    $form.CancelButton = $cancelButton

    $createButton.Add_Click({
        try {
            $clientName = $nameBox.Text.Trim()

            if ([string]::IsNullOrWhiteSpace($clientName)) {
                Show-ErrorBox "Falta escribir el nombre del cliente."
                return
            }

            $safeName = ConvertTo-SafeFolderName $clientName
            $targetFolder = Join-Path $ClientsRoot "$clientId-$safeName"

            if (Test-Path $targetFolder) {
                Show-ErrorBox "Ya existe la carpeta:`n$targetFolder"
                return
            }

            & $CreateScript -ClientId $clientId -ClientName $clientName

            Show-Info "Cliente creado correctamente:`n`n$clientId - $clientName`n`n$targetFolder"

            $form.Close()
        } catch {
            Show-ErrorBox $_.Exception.Message
        }
    })

    Update-Preview
    [void]$form.ShowDialog()

} catch {
    Show-ErrorBox $_.Exception.Message
    exit 1
}
