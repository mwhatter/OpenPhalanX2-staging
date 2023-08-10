#Aligned and ready for testing
param (
    [string]$UrlFromMain
)

$directoryForScanning = "./Integrations/Inuse"
$pythonExe = "python"

# Fetch all '*_URL.py' files
$integrationFiles = Get-ChildItem -Path $directoryForScanning -Filter "Integration_*_URL.py"

foreach ($file in $integrationFiles) {
    # Extract the provider name using regex.
    if ($file.Name -match "(.+?)_URL\.py") {
        $providerName = $matches[1]
        
        # Ask the user if they want to use this provider
        $confirmation = [System.Windows.Forms.MessageBox]::Show(
            "Do you want to use the $providerName API provider to submit the URL?", 
            "Confirm Provider Usage", 
            [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question)

        if ($confirmation -eq 'Cancel') {
            return "Action cancelled by user."
        } elseif ($confirmation -eq 'Yes') {
            # Execute the respective python script with necessary arguments
            $scriptPath = Join-Path $directoryForScanning $file.Name
            $output = & $pythonExe $scriptPath $UrlFromMain | Out-String

            # In the original script, the output was appended to a textbox and other variables were populated based on the response.
            $textboxResults.AppendText("Python script output for $providerName $output `r`n")
        }
    }
}





    