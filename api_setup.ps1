Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Define paths
$storedPath = ".\Tools\Scripts\Integrations\Stored"
$inusePath = ".\Tools\Scripts\Integrations\Inuse"

# Functions to manage API key
function Redact-ApiKeyInScript {
    param ($scriptPath)
    $scriptContent = Get-Content -Path $scriptPath -Raw
    $redactedScriptContent = $scriptContent -replace '\$apiKey\s*=\s*"([^"]+)"', '$apiKey = "redacted"'
    Set-Content -Path $scriptPath -Value $redactedScriptContent
}

function Get-ApiKeyFromScript {
    param ($scriptPath)
    $scriptContent = Get-Content -Path $scriptPath -Raw
    if ($scriptContent -match '\$apiKey\s*=\s*"([^"]+)"') {
        return $matches[1]
    } else {
        return $null
    }
}

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'API Credentials'
$form.StartPosition = 'CenterScreen'
$form.ForeColor = [System.Drawing.Color]::lightseagreen   
$form.BackColor = [System.Drawing.Color]::Black
$Form.FormBorderStyle = 'Fixed3D'

$apiProviders = @('Anomali', 'VMRay', 'VirusTotal', 'PulseDive', 'Joe-Sandbox', 'Manalyzer', 'Opswat', 'Hybrid-Analysis', 'Cuckoo')
$paddingTop = 10

foreach ($apiProvider in $apiProviders) {
    $mainCheckbox = New-Object System.Windows.Forms.CheckBox
    $mainCheckbox.Name = "${apiProvider}-mainCheckbox"
    $mainCheckbox.Location = New-Object System.Drawing.Point(10, $paddingTop)
    $mainCheckbox.Size = New-Object System.Drawing.Size(20,20)
    
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(40, $paddingTop)
    $label.Size = New-Object System.Drawing.Size(100,20)
    $label.Text = $apiProvider

    $labelURLDetonation = New-Object System.Windows.Forms.Label
    $labelURLDetonation.Location = New-Object System.Drawing.Point(150, $paddingTop)
    $labelURLDetonation.Size = New-Object System.Drawing.Size(100,20)
    $labelURLDetonation.Text = "URL Detonation?"

    $checkboxURLDetonation = New-Object System.Windows.Forms.CheckBox
    $checkboxURLDetonation.Location = New-Object System.Drawing.Point(250, $paddingTop)
    $checkboxURLDetonation.Size = New-Object System.Drawing.Size(20,20)
    $checkboxURLDetonation.Name = "${apiProvider}-URLDetonation"

    $labelFileDetonation = New-Object System.Windows.Forms.Label
    $labelFileDetonation.Location = New-Object System.Drawing.Point(280, $paddingTop)
    $labelFileDetonation.Size = New-Object System.Drawing.Size(100,20)
    $labelFileDetonation.Text = "File Detonation?"

    $checkboxFileDetonation = New-Object System.Windows.Forms.CheckBox
    $checkboxFileDetonation.Location = New-Object System.Drawing.Point(380, $paddingTop)
    $checkboxFileDetonation.Size = New-Object System.Drawing.Size(20,20)
    $checkboxFileDetonation.Name = "${apiProvider}-FileDetonation"

    $labelDirectory = New-Object System.Windows.Forms.Label
    $labelDirectory.Location = New-Object System.Drawing.Point(410, $paddingTop)
    $labelDirectory.Size = New-Object System.Drawing.Size(130,20)
    $labelDirectory.Text = "Directory Detonation?"

    $checkboxDirectory = New-Object System.Windows.Forms.CheckBox
    $checkboxDirectory.Location = New-Object System.Drawing.Point(540, $paddingTop)
    $checkboxDirectory.Size = New-Object System.Drawing.Size(20,20)
    $checkboxDirectory.Name = "${apiProvider}-Directory"
    $checkboxDirectory.Enabled = $false

    $labelIntel = New-Object System.Windows.Forms.Label
    $labelIntel.Location = New-Object System.Drawing.Point(575, $paddingTop)
    $labelIntel.Size = New-Object System.Drawing.Size(35,20)
    $labelIntel.Text = "Intel?"

    $checkboxIntel = New-Object System.Windows.Forms.CheckBox
    $checkboxIntel.Location = New-Object System.Drawing.Point(620, $paddingTop)
    $checkboxIntel.Size = New-Object System.Drawing.Size(20,20)
    $checkboxIntel.Name = "${apiProvider}-Intel"
    $checkboxIntel.Enabled = $false

    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Name = "${apiProvider}-textbox"
    $textbox.Location = New-Object System.Drawing.Point(650, $paddingTop)
    $textbox.Size = New-Object System.Drawing.Size(210,20)
    $textbox.Enabled = $false
    $textbox.BackColor = [System.Drawing.Color]::Black
    $textbox.ForeColor = [System.Drawing.Color]::lightseagreen
    $mainCheckbox.Tag = @($checkboxURLDetonation, $checkboxFileDetonation, $checkboxDirectory, $checkboxIntel, $textbox)

    $checkboxURLDetonation.Enabled = $false
    $checkboxFileDetonation.Enabled = $false
    $checkboxDirectory.Enabled = $false
    $checkboxIntel.Enabled = $false
    $textbox.Enabled = $false


    # When main checkbox is clicked
    $mainCheckbox.Add_Click({
        # The controls are stored in the Tag property of the checkbox
        $relatedControls = $this.Tag
        # Enable or disable the related controls based on checkbox state
        $relatedControls | ForEach-Object { $_.Enabled = $this.Checked }
    })    

    # Populate controls based on existing scripts
    $scriptTypes = @("URL", "File", "Directory", "Intel")
    foreach ($type in $scriptTypes) {
        $scriptName = "Integration_${apiProvider}_${type}.py"
        $scriptPath = Join-Path -Path $inusePath -ChildPath $scriptName
        if (Test-Path $scriptPath) {
            switch ($type) {
                'URL' { $checkboxURLDetonation.Checked = $true }
                'File' { $checkboxFileDetonation.Checked = $true }
                'Directory' { $checkboxDirectory.Checked = $true }
                'Intel' { $checkboxIntel.Checked = $true }
            }

            $apiKey = Get-ApiKeyFromScript -scriptPath $scriptPath
            if ($apiKey) {
                $textbox.Text = $apiKey
            }
        }
    }

    $form.Controls.AddRange(@($mainCheckbox, $label, $labelURLDetonation, $checkboxURLDetonation, $labelFileDetonation, $checkboxFileDetonation, $labelDirectory, $checkboxDirectory, $labelIntel, $checkboxIntel, $textbox))
    $paddingTop += 30
}

$executeButton = New-Object System.Windows.Forms.Button
$executeButton.Location = New-Object System.Drawing.Point(320, $paddingTop)
$executeButton.Size = New-Object System.Drawing.Size(240, 30)
$executeButton.Text = "Activate"
$executeButton.BackColor = [System.Drawing.Color]::DarkCyan
$executeButton.ForeColor = [System.Drawing.Color]::Black
$executeButton.Add_Click({
    foreach ($provider in $apiProviders) {
        $scriptTypes = @{
            "URL" = $form.Controls["${provider}-URLDetonation"]
            "File" = $form.Controls["${provider}-FileDetonation"]
            "Directory" = $form.Controls["${provider}-Directory"]
            "Intel" = $form.Controls["${provider}-Intel"]
        }

        foreach ($entry in $scriptTypes.GetEnumerator()) {
            $type = $entry.Key
            $checkboxControl = $entry.Value
            $scriptName = "Integration_${provider}_${type}.py"
            $inuseScript = Join-Path -Path $inusePath -ChildPath $scriptName
            $storedScript = Join-Path -Path $storedPath -ChildPath $scriptName

            if ($form.Controls["${provider}-mainCheckbox"].Checked) {
                if ($checkboxControl.Checked) {
                    if (-not (Test-Path $inuseScript)) {
                        # Delete the inuse script if it exists before moving
                        if (Test-Path $inuseScript) {
                            Remove-Item -Path $inuseScript -Force
                        }
                        Move-Item -Path $storedScript -Destination $inusePath
                        $apiKey = $form.Controls["${provider}-textbox"].Text
                        if ($apiKey) {
                            $scriptContent = Get-Content -Path $inuseScript -Raw
                            $updatedScriptContent = $scriptContent -replace '\$apiKey\s*=\s*"([^"]+)"', "`$apiKey = `"$apiKey`""
                            Set-Content -Path $inuseScript -Value $updatedScriptContent
                        }
                    }
                } else {
                    if (Test-Path $inuseScript) {
                        # Delete the stored script if it exists before moving
                        if (Test-Path $storedScript) {
                            Remove-Item -Path $storedScript -Force
                        }
                        Move-Item -Path $inuseScript -Destination $storedPath -Force
                        Redact-ApiKeyInScript -scriptPath $storedScript
                    }
                }
            } else {
                if (Test-Path $inuseScript) {
                    # Delete the stored script if it exists before moving
                    if (Test-Path $storedScript) {
                        Remove-Item -Path $storedScript -Force
                    }
                    Move-Item -Path $inuseScript -Destination $storedPath -Force
                    Redact-ApiKeyInScript -scriptPath $storedScript
                }
            }
        }
    }
    [System.Windows.Forms.MessageBox]::Show("Scripts updated successfully.")
})
$form.Controls.Add($executeButton)
$paddingTopup = $paddingTop + 85
$form.Size = New-Object System.Drawing.Size(890, $paddingTopup)
$form.ShowDialog()
