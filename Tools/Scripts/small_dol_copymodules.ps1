param(
    [Parameter(Mandatory=$true)]
    [string]$CopiedFilesPath
)

$computerName = if ($comboBoxComputerName.SelectedItem) {$comboBoxComputerName.SelectedItem.ToString()} else {$comboBoxComputerName.Text}

function Get-PECmdPath {
    $possibleLocations = @(".\Tools\EZTools\PECmd.exe", ".\Tools\EZTools\net*\PECmd.exe")
    foreach ($location in $possibleLocations) {
        $resolvedPaths = Resolve-Path $location -ErrorAction SilentlyContinue
        if ($resolvedPaths) {
            foreach ($path in $resolvedPaths) {
                if (Test-Path $path) {
                    return $path.Path
                }
            }
        }
    }
    throw "PECmd.exe not found in any of the known locations"
}
function Get_RemoteDriveLetters {
    param(
        [string]$ComputerName
    )

    $driveLetters = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        Get-CimInstance -ClassName Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 } | ForEach-Object { $_.DeviceID.TrimEnd(':') }
    }

    return $driveLetters
}

function Get_PrefetchMetadata {
    param(
        [string]$ComputerName,
        [string]$exportPath
    )

    if ([string]::IsNullOrEmpty($ComputerName)) {
        Write-Host "Error: ComputerName is null or empty" -ForegroundColor Red
        return
    }

    $driveLetters = Get_RemoteDriveLetters -ComputerName $ComputerName
    
    $copiedFilesPath = ".\CopiedFiles\$ComputerName\Prefetch"
    New-Item -Path $copiedFilesPath -ItemType Directory -Force

    $pecmdPath = Get-PECmdPath

    foreach ($driveLetter in $driveLetters) {
        # Convert drive letter path to UNC format
        $prefetchPath = "\\$ComputerName\$driveLetter$\Windows\Prefetch\*.pf"
        $prefetchFiles = Get-ChildItem -Path $prefetchPath -ErrorAction SilentlyContinue

        foreach ($file in $prefetchFiles) {
            # Directly copy the file using UNC path
            Copy-Item -Path $file.FullName -Destination $copiedFilesPath -Force -ErrorAction SilentlyContinue
        }
    }

    # Process the copied prefetch files with PECmd
    & $pecmdPath -d $copiedFilesPath --csv $copiedFilesPath

    # Now we'll scan the csv output
    $csvFiles = Get-ChildItem -Path $copiedFilesPath -Filter '*_PECmd_Output.csv'
    $allFilesLoaded = @()

    foreach ($csvFile in $csvFiles) {
        $csvContent = Import-Csv -Path $csvFile.FullName
        foreach ($row in $csvContent) {
            $filesLoaded = $row.FilesLoaded -split ', '
            $allFilesLoaded += $filesLoaded
        }
    }
    write-host "Found $($allFilesLoaded.Count) files loaded in prefetch files" -ForegroundColor Cyan
    # De-duplicate the list
    $uniqueFilesLoaded = $allFilesLoaded | Sort-Object | Get-Unique

    # Copy the unique files to the specified directory with their original paths in their names
    $destinationPath = ".\CopiedFiles\$ComputerName\"

    # Copy the unique files to the specified directory with their original paths in their names
    $destinationPath = ".\CopiedFiles\$ComputerName\"

    # Define the script block to handle the copy in parallel
    $scriptBlock = {
        param(
            [string]$uniqueFile,
            [string]$ComputerName,
            [string]$destinationPath
        )   
        param($uniqueFile, $ComputerName, $destinationPath)

        if ($uniqueFile -match '\\VOLUME{(.*?)}\\') {
            $volumeGUID = $matches[1]

            # Get the corresponding drive letter for the volume GUID
            $driveLetter = Get-DriveLetterFromVolumeGUID -ComputerName $ComputerName -VolumeGUID $volumeGUID

            # Replace the VOLUME{GUID} with the drive letter to form a valid path
            $filePath = $uniqueFile -replace '\\VOLUME{.*?}\\', "${driveLetter}:\" 

            # Convert the file path to UNC format
            $uncFilePath = "\\$ComputerName\$($filePath.Substring(0, 1))$" + $filePath.Substring(2)

            if (Test-Path $uncFilePath) {
                # Modify the filename to include its original path by replacing backslashes with underscores
                $modifiedFilename = $uniqueFile.Replace('\', '_').Replace(':', '')
                $destination = Join-Path -Path $destinationPath -ChildPath $modifiedFilename

                Copy-Item -Path $uncFilePath -Destination $destination -Force -ErrorAction SilentlyContinue
            }
        }
    }

    function Get-DriveLetterFromVolumeGUID {
        param (
            [string]$ComputerName,
            [string]$VolumeGUID
        )
        $volume = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            Get-WmiObject -Query "SELECT * FROM Win32_Volume WHERE DeviceID LIKE '%$VolumeGUID%'"
        }

        return $volume.DriveLetter
    }
    
    # Execute the copy in parallel
    $uniqueFilesLoaded | ForEach-Object -ThrottleLimit 100 -Parallel {
        $uniqueFile = $_
        $ComputerName = $using:computerName
        $destinationPath = $using:destinationPath
    
        if ($uniqueFile -match '\\VOLUME{(.*?)}\\') {
            $volumeGUID = $matches[1]
            $volume = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                Get-WmiObject -Query "SELECT * FROM Win32_Volume WHERE DeviceID LIKE '%$volumeGUID%'"
            }
    
            if ($volume) {
                $driveLetter = $volume.DriveLetter
    
                # Replace the VOLUME{GUID} with the drive letter to form a valid path
                $filePath = $uniqueFile -replace '\\VOLUME{.*?}\\', "${driveLetter}:\" 
    
                # Convert the file path to UNC format
                $uncFilePath = "\\$ComputerName\$($filePath.Substring(0, 1))$" + $filePath.Substring(2)
    
                if (Test-Path $uncFilePath) {
                    # Modify the filename to include its original path by replacing backslashes with underscores
                    $modifiedFilename = $uniqueFile.Replace('\', '_').Replace(':', '')
                    $destination = Join-Path -Path $destinationPath -ChildPath $modifiedFilename
    
                    Copy-Item -Path $uncFilePath -Destination $destination -Force -ErrorAction SilentlyContinue
                }
            }
        }
    }
    
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
    Get_PrefetchMetadata -ComputerName $computerName -exportPath $CopiedFilesPath
    CopyModules -computerName $computerName -CopiedFilesPath $CopiedFilesPath
    $copytreesend = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $textboxResults.AppendText("Copied all modules from $computerName at $copytreesend`r`n")
    write-host "Copied all modules from $computerName at $copytreesend" -ForegroundColor Green
} catch {
    [System.Windows.Forms.MessageBox]::Show("Error while trying to copy modules to the CopiedFiles folder: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}
