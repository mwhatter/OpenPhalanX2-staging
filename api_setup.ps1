Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Define paths
$storedPath = ".\Tools\Scripts\Integrations\Stored"
$inusePath = ".\Tools\Scripts\Integrations\Inuse"

function Get-ApiKeyFromScript {
    param ($scriptPath)
    $scriptContent = Get-Content -Path $scriptPath -Raw
    if ($scriptContent -match '\$apiKey\s*=\s*"([^"]+)"') {
        return $matches[1]
    }
    return $null
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'API Credentials'
$form.Size = New-Object System.Drawing.Size(865, 360)  # Adjusted form width
$form.StartPosition = 'CenterScreen'

$apiProviders = @('Anomali', 'VMRay', 'VirusTotal', 'PulseDive', 'Joe-Sandbox', 'Manalyzer', 'Opswat', 'Hybrid-Analysis', 'Cuckoo')
$apiTypes = @('URL', 'File', 'Directory', 'Intel')

$paddingTop = 10

foreach ($apiProvider in $apiProviders) {
    $mainCheckbox = New-Object System.Windows.Forms.CheckBox
    $mainCheckbox.Location = New-Object System.Drawing.Point(10, $paddingTop)
    $mainCheckbox.Size = New-Object System.Drawing.Size(20,20)
    
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(40, $paddingTop)
    $label.Size = New-Object System.Drawing.Size(100,20)
    $label.Text = $apiProvider

    $labelURLDetonation = New-Object System.Windows.Forms.Label
    $labelURLDetonation.Location = New-Object System.Drawing.Point(140, $paddingTop)
    $labelURLDetonation.Size = New-Object System.Drawing.Size(100,20)
    $labelURLDetonation.Text = "URL Detonation?"

    $checkboxURLDetonation = New-Object System.Windows.Forms.CheckBox
    $checkboxURLDetonation.Location = New-Object System.Drawing.Point(240, $paddingTop)
    $checkboxURLDetonation.Size = New-Object System.Drawing.Size(20,20)
    $checkboxURLDetonation.Name = "$apiProvider-URLDetonation?"

    $labelFileDetonation = New-Object System.Windows.Forms.Label
    $labelFileDetonation.Location = New-Object System.Drawing.Point(270, $paddingTop)
    $labelFileDetonation.Size = New-Object System.Drawing.Size(100,20)
    $labelFileDetonation.Text = "File Detonation?"

    $checkboxFileDetonation = New-Object System.Windows.Forms.CheckBox
    $checkboxFileDetonation.Location = New-Object System.Drawing.Point(370, $paddingTop)
    $checkboxFileDetonation.Size = New-Object System.Drawing.Size(20,20)
    $checkboxFileDetonation.Name = "$apiProvider-FileDetonation?"

    $labelDirectory = New-Object System.Windows.Forms.Label
    $labelDirectory.Location = New-Object System.Drawing.Point(400, $paddingTop)
    $labelDirectory.Size = New-Object System.Drawing.Size(130,20)
    $labelDirectory.Text = "Directory Detonation?"

    $checkboxDirectory = New-Object System.Windows.Forms.CheckBox
    $checkboxDirectory.Location = New-Object System.Drawing.Point(530, $paddingTop)  # Adjusted location
    $checkboxDirectory.Size = New-Object System.Drawing.Size(20,20)
    $checkboxDirectory.Name = "$apiProvider-Directory"
    $checkboxDirectory.Enabled = $false

    # Store this checkbox with the others in the Tag property of the mainCheckbox
    $mainCheckbox.Tag += $checkboxDirectory

    $labelIntel = New-Object System.Windows.Forms.Label
    $labelIntel.Location = New-Object System.Drawing.Point(560, $paddingTop)
    $labelIntel.Size = New-Object System.Drawing.Size(40,20)
    $labelIntel.Text = "Intel?"

    $checkboxIntel = New-Object System.Windows.Forms.CheckBox
    $checkboxIntel.Location = New-Object System.Drawing.Point(600, $paddingTop)
    $checkboxIntel.Size = New-Object System.Drawing.Size(20,20)
    $checkboxIntel.Name = "$apiProvider-Intel?"

    $checkboxURLDetonation.Enabled = $false
    $checkboxFileDetonation.Enabled = $false
    $checkboxIntel.Enabled = $false

    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Location = New-Object System.Drawing.Point(630, $paddingTop)
    $textbox.Size = New-Object System.Drawing.Size(210,20)
    $textbox.Enabled = $false

    # Store the related controls in an array in the Tag property
    $mainCheckbox.Tag = @($textbox, $checkboxURLDetonation, $checkboxFileDetonation, $checkboxIntel)

    # Using Add_Click instead of Add_CheckedChanged
    $mainCheckbox.Add_Click({
        # The controls are stored in the Tag property of the checkbox
        $relatedControls = $this.Tag
        # Enable or disable the related controls based on checkbox state
        $relatedControls | ForEach-Object { $_.Enabled = $this.Checked }
    })

    # Check if corresponding script exists in "Inuse" directory
    foreach ($apiType in $apiTypes) {
        $scriptName = "Integration_${apiProvider}_${apiType}.py"
        $scriptPath = Join-Path -Path $inusePath -ChildPath $scriptName
        if (Test-Path $scriptPath) {
            switch ($apiType) {
                'URL' {
                    $checkboxURLDetonation.Checked = $true
                }
                'File' {
                    $checkboxFileDetonation.Checked = $true
                }
                'Directory' {
                    # Assuming you have another checkbox for 'Directory'. If not, adjust accordingly.
                    $checkboxDirectory.Checked = $true
                }
                'Intel' {
                    $checkboxIntel.Checked = $true
                }
            }

            # Populate the textbox with the API key from the script
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
$executeButton.Location = New-Object System.Drawing.Point(320, 280)
$executeButton.Size = New-Object System.Drawing.Size(150, 30)
$executeButton.Text = "Activate"
$executeButton.Add_Click({
    # Clear 'Inuse' directory first
    Get-ChildItem $inUsePath -Filter "Integration_*.py" | ForEach-Object {
        Move-Item $_.FullName (Join-Path $storedPath $_.Name) -Force
    }

    foreach ($apiProvider in $apiProviders) {
        $mainCheckbox = $form.Controls | Where-Object { $_.Name -eq "$apiProvider-mainCheckbox" }
        $textbox = $form.Controls | Where-Object { $_.Name -eq "$apiProvider-textbox" }

        if ($mainCheckbox.Checked) {
            $apiKeyValue = $textbox.Text

            # Determine which scripts to move based on checkboxes
            foreach ($apiType in $apiTypes) {
                $checkBox = $form.Controls | Where-Object { $_.Name -eq "$apiProvider-$apiType" }
                if ($checkBox.Checked) {
                    $scriptName = "Integration_${apiProvider}_${apiType}.py"
                    $scriptPath = Join-Path $storedPath $scriptName
                    $destinationPath = Join-Path $inUsePath $scriptName

                    Move-Item $scriptPath $destinationPath -Force
                    # Replace the API key in the moved script
                    (Get-Content $destinationPath) -replace '\$apiKey = "redacted"', "`$apiKey = `"$apiKeyValue`"" | Set-Content $destinationPath
                }
            }
        }
    }
})

$form.Controls.Add($executeButton)
$form.ShowDialog()
