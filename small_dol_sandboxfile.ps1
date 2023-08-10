#Aligned and ready for testing
param (
    [string]$SampleFileFromMain
)

$directoryForScanning = "./Integrations/Inuse"
$pythonExe = "python"

# Fetch all 'Integration_Provider_File.py' files
$integrationFiles = Get-ChildItem -Path $directoryForScanning -Filter "Integration_*_File.py"

foreach ($file in $integrationFiles) {
    # Extract the provider name using regex.
    if ($file.Name -match "Integration_(.+?)_File\.py") {
        $providerName = $matches[1]
        
        # Ask the user if they want to use this provider
        $confirmation = [System.Windows.Forms.MessageBox]::Show(
            "Do you want to use the $providerName API provider to submit the file?", 
            "Confirm Provider Usage", 
            [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question)

        if ($confirmation -eq 'Cancel') {
            return "Action cancelled by user."
        } elseif ($confirmation -eq 'Yes') {
            # Execute the respective python script with necessary arguments
            $scriptPath = Join-Path $directoryForScanning $file.Name
            $output = & $pythonExe $scriptPath $SampleFileFromMain | Out-String

            # Post-processing based on the content of $output
            $textboxResults.AppendText("Python script output for $providerName $output `r`n")

            $outputObject = $output | ConvertFrom-Json
            $script:Note = $outputObject.Note
            $textboxResults.AppendText("Tags added: $script:Note `r`n")

            if ($null -ne $script:Note) {
                $buttonRetrieveReportfile.Enabled = $true
            }
        }
    }
}
