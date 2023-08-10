$computerName = if ($comboBoxComputerName.SelectedItem) {
    $comboBoxComputerName.SelectedItem.ToString()
} else {
    $comboBoxComputerName.Text
}
$localFilePath = if ($comboboxlocalFilePath.SelectedItem) {
    $comboboxlocalFilePath.SelectedItem.ToString()
} else {
    $comboboxlocalFilePath.Text
}
$remoteFilePath = $textboxRemoteFilePath.Text
$additionalArgs = $textboxaddargs.Text

if (![string]::IsNullOrEmpty($computerName) -and ![string]::IsNullOrEmpty($localFilePath) -and ![string]::IsNullOrEmpty($remoteFilePath)) {
    $filename = [System.IO.Path]::GetFileName($localFilePath)
    $fileExtension = [System.IO.Path]::GetExtension($localFilePath).ToLower()
    $remoteDriveLetter = $remoteFilePath.Substring(0, 1)
    $remoteDirectoryPath = "\\$computerName\$remoteDriveLetter$" + $remoteFilePath.Substring(2)
    $remoteUncPath = Join-Path -Path $remoteDirectoryPath -ChildPath $filename

    try {
        if ($fileExtension -eq ".ps1") {
            $dialogResult = [System.Windows.Forms.MessageBox]::Show("Would you like to execute the script in-memory (fileless)?", "Fileless Execution", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($dialogResult -eq "Yes") {
                # Read PS1 file and convert to Base64
                $scriptContent = Get-Content $localFilePath -Raw
                # Create a script block that calls the original script with arguments
                $scriptWithArgs = "$scriptContent; & { $scriptContent } $additionalArgs"
                $bytes = [System.Text.Encoding]::Unicode.GetBytes($scriptWithArgs)
                $encodedCommand = [Convert]::ToBase64String($bytes)
                $commandLine = "powershell.exe -EncodedCommand $encodedCommand"
            }
            else {
                # Normal execution
                Copy-Item -Path $localFilePath -Destination $remoteUncPath -Force
                $commandLine = "powershell.exe -File $remoteUncPath $additionalArgs"
            }
        }
        else {
            Copy-Item -Path $localFilePath -Destination $remoteUncPath -Force
            switch ($fileExtension) {
                ".cmd" { $commandLine = "cmd.exe /c $remoteUncPath $additionalArgs" }
                ".bat" { $commandLine = "cmd.exe /c $remoteUncPath $additionalArgs" }
                ".js" { $commandLine = "cscript.exe $remoteUncPath $additionalArgs" }
                ".vbs" { $commandLine = "cscript.exe $remoteUncPath $additionalArgs" }
                ".dll" { $commandLine = "rundll32.exe $remoteUncPath,$additionalArgs" }
                ".py" { $commandLine = "python $remoteUncPath $additionalArgs" }
                
                default { $commandLine = "$remoteUncPath $additionalArgs" }
            }
        }
        $session = New-CimSession -ComputerName $computerName
        $newProcess = $session | Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{CommandLine = $commandLine}
        $returnValue = $newProcess.ReturnValue
        if ($returnValue -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("File '$filename' executed successfully on the remote computer with additional arguments: $additionalArgs", "Execution Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $textboxResults.AppendText("File '$filename' executed successfully on the remote computer with additional arguments: $additionalArgs `r`n")
        } else {
            throw "Error executing file on remote computer. Error code: $returnValue"
        }
        $session | Remove-CimSession
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error executing file '$filename' on the remote computer: $($_.Exception.Message)", "Execution Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $textboxResults.AppendText("Error executing file '$filename' on the remote computer: $($_.Exception.Message)`r`n")
    }
}