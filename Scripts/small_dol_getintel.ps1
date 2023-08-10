#Aligned and ready for testing
param (
    [string]$IndicatorFromMain
)

$directoryForScanning = "./Integrations/Inuse"
$pythonExe = "python"

# Fetch all 'Integration_Provider_Intel.py' files
$integrationFiles = Get-ChildItem -Path $directoryForScanning -Filter "Integration_*_Intel.py"

foreach ($file in $integrationFiles) {
    # Extract the provider name using regex for readability. 
    # The provider name should be between "Integration_" and "_Intel.py"
    if ($file.Name -match "Integration_(.+?)_Intel\.py") {
        $providerName = $matches[1]
        
        # Ask the user if they want to use this provider
        $confirmation = [System.Windows.Forms.MessageBox]::Show(
            "Do you want to use the $providerName API provider to fetch intel for the indicator?", 
            "Confirm Provider Usage", 
            [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question)

        if ($confirmation -eq 'Cancel') {
            return "Action cancelled by user."
        } elseif ($confirmation -eq 'Yes') {
            # Execute the respective python script with necessary arguments
            $scriptPath = Join-Path $directoryForScanning $file.Name
            $output = & $pythonExe $scriptPath $IndicatorFromMain | Out-String

            # Here you can perform the post-processing of the output.
            # For instance:
            if ($output -contains "specific-string") {
                # Do something specific based on the content of $output
            }
            return "Python script output for $providerName $output"
        }
    }
}
