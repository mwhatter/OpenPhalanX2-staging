$computerName = if ($comboBoxComputerName.SelectedItem) {
    $comboBoxComputerName.SelectedItem.ToString()
} else {
    $comboBoxComputerName.Text
}
$RemoteDriveLetters = Get_RemoteDriveLetters -HostName $computerName
    foreach ($DriveLetter in $RemoteDriveLetters) {
        $result = Copy_USNJournal -HostName $computerName -Destination $exportPath -DriveLetter $DriveLetter
        if ($result -and $result.StartsWith("Error")) {
            Write-Host $result -ForegroundColor Red
        } elseif ($result) {
            Write-Host $result -ForegroundColor Green
        }
        }

function Copy_USNJournal {
    param(
        [string]$HostName,
        [string]$Destination,
        [string]$DriveLetter
    )
    New-Item -ItemType Directory -Path ".\Logs\Reports\$computerName\USN_Journal" -Force | Out-Null
    $LocalUSNJournalPath = Join-Path -Path ($Destination + '\' + "$HostName\USN_Journal") -ChildPath "UsnJrnl_$($HostName)-$DriveLetter.csv"
    $timestart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "Starting USN Journal Extraction at $timestart" -Foregroundcolor Cyan
    $textboxResults.AppendText("Starting USN Journal Extraction at $timestart `r`n") | Out-Null

    try {
        # Executing fsutil and processing its output
        $scriptBlock = {
            param($DriveLetter)
            & "C:\Windows\System32\fsutil.exe" "usn" "readjournal" "$DriveLetter`:" | Out-String -Stream 
        }
                
        $remoteOutput = Invoke-Command -ComputerName $HostName -ScriptBlock $scriptBlock -ArgumentList $DriveLetter
        $timeend = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Write-Host "Finished USN Journal Extraction from $DriveLetter at $timeend" -Foregroundcolor Cyan
        Write-Host "Tableling the data" -Foregroundcolor Cyan

        # Split the raw output into lines
        if ($null -ne $remoteOutput) {
            $lines = $remoteOutput -split "`r`n"
        } else {
            Write-Warning "No output from Invoke-Command."
            return
        }

        # Find the start of the journal entries
        $firstUsnLine = $lines | Where-Object { $_.StartsWith('Usn') } | Select-Object -First 1
        if ($null -ne $firstUsnLine) {
            $journalStartIndex = $lines.IndexOf($firstUsnLine)

        # Get only the journal entries (ignore metadata)
        $lines = $lines[$journalStartIndex..$lines.Count]

        # Prepare data collection
        $entry = @{}

        # Exporting data to CSV
        $lines | ForEach-Object {
            if ($_ -eq '') {
                if ($entry.Count -gt 0) {
                    $entry.PSObject.Copy() # Emit the entry
                    $entry.Clear()
                }
            } else {
                $parts = $_ -split ':', 2
                $entry[$parts[0].Trim()] = if ($parts.Length -gt 1) { $parts[1].Trim() } else { $null }
            }
        } | Export-Csv -Path $LocalUSNJournalPath -NoTypeInformation
        } else {
            Write-Warning "No lines starting with 'Usn' were found."
        }

        $timeend = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Write-Host "USN Journal from drive $DriveLetter on $HostName finished processing at $Timeend" -ForegroundColor Green
        $textboxResults.AppendText("USN Journal from drive $DriveLetter on $HostName finished processing at $Timeend `r`n")
    } catch {
        return "Error copying USN Journal from $HostName on drive $DriveLetter $($_.Exception.Message)"
    }
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