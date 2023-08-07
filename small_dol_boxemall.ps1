$directoryPath = ".\CopiedFiles"
    $pythonScript = "mass_submit_files_anomali.py"
    $pythonExe = "python"
    $apiKey = "redacted"
    
    # Get the number of files in the directory
    $fileCount = (Get-ChildItem -Path $directoryPath -File).Count

    # Ask for confirmation
    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "You are about to submit $fileCount files. Do you wish to continue?", 
        "Confirm Action", 
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question)

    if($confirmation -eq 'Yes') {
        $output = & $pythonExe $pythonScript $directoryPath $apiKey | Out-String
        $textboxResults.AppendText("Python script output: $output")
    }