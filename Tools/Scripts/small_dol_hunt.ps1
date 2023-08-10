$remoteComputer = if ($comboBoxComputerName.SelectedItem) {
    $comboBoxComputerName.SelectedItem.ToString()
} else {
    $comboBoxComputerName.Text
}
$remoteFilePath = $textboxRemoteFilePath.Text
$drivelessPath = Split-Path $remoteFilePath -NoQualifier
$driveLetter = Split-Path $remoteFilePath -Qualifier
$filehunt = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "Hunting for file: $remoteFilePath at $filehunt" -ForegroundColor Cyan
$textboxResults.AppendText("Hunting for file: $remoteFilePath at $filehunt `r`n")
Log_Message -logfile $logfile -Message "Hunting for file: $remoteFilePath"
if (![string]::IsNullOrEmpty($remoteComputer) -and ![string]::IsNullOrEmpty($remoteFilePath)) {
    $localPath = '.\CopiedFiles'
    if (!(Test-Path $localPath)) {
        New-Item -ItemType Directory -Force -Path $localPath
    }

    try {
        $fileExists = Invoke-Command -ComputerName $remoteComputer -ScriptBlock {
            param($path)
            Test-Path -Path $path
        } -ArgumentList $remoteFilePath -ErrorAction Stop
    } catch {
        Write-Error "Failed to check if the file exists on the remote computer: $_"
        return
    }

    if ($fileExists) {
        $destination = Join-Path -Path $localPath -ChildPath (Split-Path $remoteFilePath -Leaf)
        Copy-Item -Path "\\$remoteComputer\$($remoteFilePath.Replace(':', '$'))" -Destination $destination -Force
        $textboxResults.AppendText("$remoteFilePath copied from $remoteComputer.")
        Log_Message -Message "$remoteFilePath copied from $remoteComputer" -LogFilePath $LogFile
        Write-Host "$remoteFilePath copied from $remoteComputer" -ForegroundColor -Cyan
    } else {
        $foundInRecycleBin = $false

        try {
            $recycleBinItems = Invoke-Command -ComputerName $remoteComputer -ScriptBlock {
                Get-ChildItem 'C:\$Recycle.Bin' -Recurse -Force | Where-Object { $_.PSIsContainer -eq $false -and $_.Name -like "$I*" } | ForEach-Object {
                    $originalFilePath = Get-Content $_.FullName -ErrorAction SilentlyContinue | Select-Object -First 1
                    $originalFilePath = $originalFilePath -replace '^[^\:]*\:', '' # Remove non-printable characters before the colon
                    [PSCustomObject]@{
                        Name = $_.Name
                        FullName = $_.FullName
                        OriginalFilePath = $originalFilePath
                    }
                }
            } -ErrorAction Stop
        } catch {
            Write-Error "Failed to retrieve the items from the recycle bin: $_ `r`n"
            return
        }

        $recycleBinItem = $recycleBinItems | Where-Object { $_.OriginalFilePath -eq "$drivelessPath" }
        if ($recycleBinItem) {
            Write-Host "Match found in Recycle Bin: $($recycleBinItem.Name)" -Foregroundcolor Cyan
            $textboxResults.AppendText("Match found in Recycle Bin: $($recycleBinItem.Name) `r`n")
            $foundInRecycleBin = $true
        }

        if (!$foundInRecycleBin) {
            $textboxResults.AppendText("File not found in the remote computer or recycle bin. `r`n")
        } else {
            try {
                $vssServiceStatus = Invoke-Command -ComputerName $remoteComputer -ScriptBlock {
                    $service = Get-Service -Name VSS
                    $status = $service.Status
                    return $status
                } 
                
                $statusCodes = @{
                    1 = "Stopped"
                    2 = "Start Pending"
                    3 = "Stop Pending"
                    4 = "Running"
                    5 = "Continue Pending"
                    6 = "Pause Pending"
                    7 = "Paused"
                }
                
                # Use the hash table to get the corresponding status name
                $vssServiceStatusName = $statusCodes[$vssServiceStatus]
                
                # Print the status name
                Write-Host "VSS service status: $vssServiceStatusName" -ForegroundColor Cyan
                $textboxResults.AppendText("VSS service status on $remoteComputer $vssServiceStatusName `r`n")

                if ($vssServiceStatus -eq 'Running') {
                    $shadowCopyFileExists = Invoke-Command -ComputerName $remoteComputer -ScriptBlock {
                        param($path, $driveLetter)
                        $shadowCopies = vssadmin list shadows /for=$driveLetter | Where-Object { $_ -match 'GLOBALROOT\\Device\\HarddiskVolumeShadowCopy\\d+' }
                        foreach ($shadowCopy in $shadowCopies) {
                            $shadowCopyPath = $shadowCopy -replace '.*?(GLOBALROOT\\Device\\HarddiskVolumeShadowCopy\\d+).*', '$1'
                            if (Test-Path -Path "$shadowCopyPath\$path") {
                                return $shadowCopyPath
                            }
                        }
                    } -ArgumentList $drivelessPath, $driveLetter -ErrorAction Stop

                    if ($shadowCopyFileExists) {
                        Write-Host "Shadow copy found: $shadowCopyFileExists" -ForegroundColor Cyan
                        $textboxResults.AppendText("Shadow copy found: $shadowCopyFileExists `r`n")
                    } else {
                        Write-Host "No shadow copy found for the file." -ForegroundColor Red
                        $textboxResults.AppendText("No shadow copy found for the file. `r`n")
                    }
                }
            } catch {
                Write-Error "Failed to check VSS service or shadow copies: $_"
                return
            }
        $restorePoints = Invoke-Expression -Command "wmic /Namespace:\\root\default Path SystemRestore get * /format:list"
        if ($restorePoints) {
            Write-Host "System Restore points exist on $remoteComputer" -ForegroundColor Cyan
            $textboxResults.AppendText("System Restore points exist on $remoteComputer `r`n")
        } else {
            Write-Host "No System Restore points found." -ForegroundColor Red
            $textboxResults.AppendText("No System Restore points found. `r`n")
        }
        $lastBackup = Invoke-Expression -Command "wbadmin get versions"
        if ($lastBackup) {
            Write-Host "Backups exist on $remoteComputer" -ForegroundColor Cyan
            $textboxResults.AppendText("Backups exist on $remoteComputer `r`n")
        } else {
            Write-Host "No backups found." -ForegroundColor Red
            $textboxResults.AppendText("No backups found. `r`n")
        }
        }
    }
} else {
    [System.Windows.Forms.MessageBox]::Show("Please enter a remote computer name and file path.", "Missing Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
}
$fileput = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
write-host "File Hunt completed at $fileput" -ForegroundColor Green 
Log_Message -logfile $logfile -Message "File Hunt completed"
$textboxResults.AppendText("File Hunt completed at $fileput `r`n")