#Aligned and ready for testing
$directoryForScanning = "./Integrations/Inuse"
$directoryPath = ".\CopiedFiles"
$pythonExe = "python"

# Fetch all 'Integration_Provider_Directory.py' files
$integrationFiles = Get-ChildItem -Path $directoryForScanning -Filter "Integration_*_Directory.py"

foreach ($file in $integrationFiles) {
    # Extract the provider name using regex for readability. 
    # The provider name should be between "Integration_" and "_Directory.py"
    if ($file.Name -match "Integration_(.+?)_Directory\.py") {
        $providerName = $matches[1]
        
        # Ask the user if they want to use this provider
        $confirmation = [System.Windows.Forms.MessageBox]::Show(
            "Do you want to use the $providerName API provider for submissions?", 
            "Confirm Provider Usage", 
            [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question)

        if ($confirmation -eq 'Cancel') {
            return "Action cancelled by user."
        } elseif ($confirmation -eq 'Yes') {
            # Execute the respective python script with necessary arguments
            $scriptPath = Join-Path $directoryForScanning $file.Name
            $output = & $pythonExe $scriptPath $directoryPath | Out-String
            return "Python script output for $providerName $output"
        }
    }
}
