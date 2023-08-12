param(
    [Parameter(Mandatory=$true)]
    [string]$CopiedFilesPath
)

$computerName = if ($comboBoxComputerName.SelectedItem) {$comboBoxComputerName.SelectedItem.ToString()} else {$comboBoxComputerName.Text}

function Get_RemoteDriveLetters {
    param(
        [string]$ComputerName
    )

    $driveLetters = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        Get-CimInstance -ClassName Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 } | ForEach-Object { $_.DeviceID.TrimEnd(':') }
    }

    return $driveLetters
}

function CopyModules {
    param (
        [string]$ComputerName,
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

            # Convert the modulePath to a valid filename by replacing backslashes with underscores
            $modifiedModuleName = $modulePath.Replace('\', '_').Replace(':', '')
            
            $destinationPath = Join-Path -Path $using:CopiedFilesPath -ChildPath $modifiedModuleName

            Copy-Item $uncPath -Destination $destinationPath\$ComputerName -Force -ErrorAction SilentlyContinue
        }
    }
}

$copytreestart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$textboxResults.AppendText("Started copying all modules from $computerName at $copytreestart`r`n")
write-host "Started copying all modules from $computerName at $copytreestart" -ForegroundColor Green

try {
    CopyModules -computerName $computerName -CopiedFilesPath $CopiedFilesPath
    $copytreesend = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $textboxResults.AppendText("Copied all modules from $computerName at $copytreesend`r`n")
    write-host "Copied all modules from $computerName at $copytreesend" -ForegroundColor Green
} catch {
    [System.Windows.Forms.MessageBox]::Show("Error while trying to copy modules to the CopiedFiles folder: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}
