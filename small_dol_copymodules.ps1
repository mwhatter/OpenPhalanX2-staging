function CopyModules {
    param (
        [string]$computerName,
        [string]$CopiedFilesPath
    )

    $uniqueModules = Invoke-Command -ComputerName $computerName -ScriptBlock {
        Get-Process | ForEach-Object { $_.Modules } | Select-Object -Unique -ExpandProperty FileName
    }

    $uniqueModules | ForEach-Object -ThrottleLimit 100 -Parallel {
        $modulePath = $_
        if (![string]::IsNullOrEmpty($modulePath)) {
            # Convert the drive letter path to UNC format
            $driveLetter = $modulePath.Substring(0, 1)
            $uncPath = "\\$using:computerName\$driveLetter$" + $modulePath.Substring(2)

            $moduleFilename = [System.IO.Path]::GetFileName($modulePath)
            $destinationPath = Join-Path -Path $using:CopiedFilesPath -ChildPath $moduleFilename

            Copy-Item $uncPath -Destination $destinationPath -Force -ErrorAction SilentlyContinue
        }
    }
}

$computerName = if ($comboBoxComputerName.SelectedItem) {
    $comboBoxComputerName.SelectedItem.ToString()
} else {
    $comboBoxComputerName.Text
}
$copytreestart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Write-Host "Started copying all modules from $computerName at $copytreestart" -ForegroundColor Cyan 
        Log_Message -logfile $logfile -Message "Copied all modules from $computerName"

$CopiedFilesPath = $CopiedFilesDir

if (![string]::IsNullOrEmpty($computerName) -and ![string]::IsNullOrEmpty($CopiedFilesPath)) {
    try {
        CopyModules -computerName $computerName -CopiedFilesPath $CopiedFilesPath

        
        $copytreesend = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Write-Host "Copied all modules from $computerName at $copytreesend" -ForegroundColor Cyan 
        [System.Windows.Forms.MessageBox]::Show("All uniquely pathed modules copied to the CopiedFiles folder.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        Log_Message -logfile $logfile -Message "Copied all modules from $computerName"
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error while trying to copy modules to the CopiedFiles folder: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
} else {
    [System.Windows.Forms.MessageBox]::Show("Please enter a valid computer name and CopiedFiles path.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}