$computerName = if ($comboBoxComputerName.SelectedItem) {
    $comboBoxComputerName.SelectedItem.ToString()
} else {
    $comboBoxComputerName.Text
}

# Call the new function to export all logs
Export_AllLogs -EVTXPath $EVTXPath -HostName $computerName

Process_Hayabusa -HostName $computerName 
Process_DeepBlueCLI -HostName $computerName

$colwinstart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "WinEventalyzer of $computerName complete at $colwinstart" -ForegroundColor Green
Log_Message -logfile $logfile -message "WinEventalyzer of $computerName complete"
$textboxResults.AppendText("WinEventalyzer of $computerName complete at $colwinstart `r`n")

function Export_AllLogs {
    param (
        $EVTXPath,
        $HostName
    )

    Add-Type -AssemblyName System.Windows.Forms # Only required in PowerShell Core (7+) 

    $colhoststart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "Wineventalyzing logs from $HostName to $EVTXPath at $colhoststart" -ForegroundColor Cyan 
    Log_Message -logfile $logfile -Message "Exporting logs from $HostName to $EVTXPath"
    $textboxResults.AppendText("Exporting logs from $HostName to $EVTXPath at $colhoststart `r`n")

    if (!(Test-Path $EVTXPath\$HostName)) {
        New-Item -Path $EVTXPath\$HostName -ItemType Directory
    }
    
    # Check and prompt for overwriting or proceeding with existing logs
    $localLogPath = Join-Path -Path $EVTXPath -ChildPath $HostName
    if(Test-Path $localLogPath){
        $files = Get-ChildItem -Path $localLogPath -File
        $fileCount = $files.Count
        Write-Host "There are $fileCount files in the $localLogPath directory." -ForegroundColor Cyan
        
        $userInput = [System.Windows.Forms.MessageBox]::Show("There are $fileCount files in the directory. Do you want to overwrite the logs?", "Confirm", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel)
        if ($userInput -eq 'No') { 
            Write-Host "Re-analyzing $fileCount logs." -ForegroundColor Cyan
            Log_Message -logfile $logfile -Message "Re-analyzing $fileCount logs."
            $textboxResults.AppendText("Re-analyzing $fileCount logs. `r`n")
            return 
        }
        if ($userInput -eq 'Cancel') { exit }
    }

    # Get remote drive letters
    $driveLetters = Get_RemoteDriveLetters -HostName $HostName

    $logsCopied = 0

    foreach ($driveLetter in $driveLetters) {
        # Remote log path
        $remoteLogPath1 = "\\$HostName\$driveLetter$\Windows\System32\winevt\Logs"
        $remoteLogPath2 = "\\$HostName\$driveLetter$\Windows\System32\winevt\EventLogs"
    
        if ((Test-Path $remoteLogPath1) -or (Test-Path $remoteLogPath2)) {
            # Get all logs from remote host
            $logs = Get-ChildItem -Path $remoteLogPath1 -Filter *.evtx -ErrorAction SilentlyContinue
            $logs += Get-ChildItem -Path $remoteLogPath2 -Filter *.evtx -ErrorAction SilentlyContinue
            
            # Copy the files
            $copiedLogs = $logs | ForEach-Object -ThrottleLimit 100 -Parallel {
                # Skip logs of size 68KB and logs with "configuration" in the name
                if ($_.Length -ne 69632 -and $_.Name -notlike "*onfiguration*") {
                    $localLogPath = Join-Path $using:EVTXPath $using:HostName "$($_.Name)"
                    Copy-Item -LiteralPath $_.FullName -Destination $localLogPath -Force -ErrorAction SilentlyContinue
                    return $localLogPath
                }
            }
    
            # Count the copied files
            $logsCopied += $copiedLogs.Count
        } 
    }
    
        $evTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Write-Host "Collected $logsCopied logs from $HostName at $evTime" -ForegroundColor Cyan
        Log_Message -Message "Collected $logsCopied logs from $HostName" -LogFilePath $LogFile
        $textboxResults.AppendText("Collected $logsCopied logs from $HostName at $evTime `r`n")
    } 


function Process_Hayabusa {
    param (
        [string]$HostName
    )
    New-Item -Path "$exportPath\$HostName" -ItemType Directory -Force | Out-Null
    New-Item -Path "$exportPath\$HostName\Hayabusa" -ItemType Directory -Force | Out-Null  
    $hayabusaPath = "$CWD\Tools\Hayabusa\hayabusa.exe"

    & $hayabusaPath logon-summary -d "$EVTXPath\$HostName" -C -o "$exportPath\$HostName\Hayabusa\logon-summary.csv" 
    & $hayabusaPath metrics -d "$EVTXPath\$HostName" -C -o "$exportPath\$HostName\Hayabusa\metrics.csv" 
    & $hayabusaPath pivot-keywords-list -d "$EVTXPath\$HostName" -o "$exportPath\$HostName\Hayabusa\pivot-keywords-list.csv" 
    & $hayabusaPath csv-timeline -d "$EVTXPath\$HostName" -C -o "$exportPath\$HostName\Hayabusa\csv-timeline.csv" -p super-verbose 
    $colhayastart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "Hayabusa analyzed logs from $HostName at $colhayastart" -ForegroundColor Cyan 
    Log_Message -logfile $logfile -Message "Hayabusa analyzed logs from $HostName"
    $textboxResults.AppendText("Hayabusa analyzed logs from $HostName at $colhayastart `r`n")
}

function Process_DeepBlueCLI {
    param (
        [string]$HostName
    )
    New-Item -Path "$exportPath\$HostName" -ItemType Directory -Force | Out-Null
    New-Item -Path "$exportPath\$HostName\DeepBlueCLI" -ItemType Directory -Force | Out-Null
    $deepBlueCLIPath = "$CWD\Tools\DeepBlueCLI\DeepBlue.ps1"
    
    $logFiles = @("security.evtx", "system.evtx", "Application.evtx", "Microsoft-Windows-Sysmon%4Operational.evtx", "Windows PowerShell.evtx", "Microsoft-Windows-AppLocker%4EXE and DLL.evtx")
    foreach($logFile in $logFiles) {
        $outFile = Join-Path "$exportPath\$HostName\DeepBlueCLI" "$($logFile -replace ".evtx", ".txt")"
        try {
            & $deepBlueCLIPath "$EVTXPath\$HostName\$logFile" -ErrorAction SilentlyContinue | Out-File -FilePath $outFile
        }
        catch {
            # Suppress the error message by doing nothing
        }
    }

    $coldeepstart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "DeepBlueCLI analyzed logs from $HostName at $coldeepstart" -ForegroundColor Cyan 
    Log_Message -logfile $logfile -Message "DeepBlueCLI analyzed logs from $HostName"
    $textboxResults.AppendText("DeepBlueCLI analyzed logs from $HostName at $coldeepstart `r`n")
}

function Get_RemoteDriveLetters {
    param(
        [string]$HostName
    )

    $driveLetters = Invoke-Command -ComputerName $HostName -ScriptBlock {
        Get-CimInstance -ClassName Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 } | ForEach-Object { $_.DeviceID.TrimEnd(':') }
    }

    return $driveLetters
}