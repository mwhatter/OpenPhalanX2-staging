$computerName = if ($comboBoxComputerName.SelectedItem) {
    $comboBoxComputerName.SelectedItem.ToString()
} else {
    $comboBoxComputerName.Text
}
CollectForensicTimeline -HostName $computerName
Get_PrefetchMetadata -HostName $computerName -exportPath $exportPath
Copy_BrowserHistoryFiles -RemoteHost $computerName -LocalDestination "$exportPath\$computerName\RapidTriage\User_Browser&PSconsole"
$colallstart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "RapidTriage of $computerName completed at $colallstart" -ForegroundColor Green 
Log_Message -logfile $logfile -Message "RapidTriage of $computerName completed"
$textboxResults.AppendText("RapidTriage of $computerName completed at $colallstart `r`n")

function CollectForensicTimeline {
    param (
        [string[]]$Hostnames
    )
    
    foreach ($Hostname in $Hostnames) {
        try {
            $session = New-CimSession -ComputerName $Hostname -ErrorAction SilentlyContinue

            if (-not $session) {
                [System.Windows.Forms.MessageBox]::Show("Failed to connect to $Hostname. Please check the hostname and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                continue
            }

            New-Item -ItemType Directory -Path ".\Logs\Reports\$Hostname\RapidTriage" -Force | Out-Null
            $ExcelFile = Join-Path ".\Logs\Reports\$Hostname\RapidTriage" "$Hostname-RapidTriage.xlsx"

            $SystemUsers = Get-CimInstance -ClassName Win32_SystemUsers -CimSession $session
            if ($SystemUsers) {
                $colSystStart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Write-Host "Collected System Users from $HostName at $colSystStart" -ForegroundColor Cyan 
                Log_Message -logfile $logfile -Message "Collected System Users from $HostName"
                $textboxResults.AppendText("Collected System Users from $HostName at $colSystStart `r`n")
                $ExcelSystemUsers = $SystemUsers | Export-Excel -Path $ExcelFile -WorksheetName 'SystemUsers' -AutoSize -AutoFilter -TableStyle Medium6 -Append -PassThru
                Close-ExcelPackage $ExcelSystemUsers
            }

            $Processes = Get-CimInstance -ClassName Win32_Process -CimSession $session | select-object -Property CreationDate,CSName,ProcessName,CommandLine,Path,ProcessId,ParentProcessId
			if ($Processes) {
				$colProcStart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
				Write-Host "Collected Processes from $HostName at $colProcStart" -ForegroundColor Cyan 
				Log_Message -logfile $logfile -Message "Collected Processes from $HostName"
                $textboxResults.AppendText("Collected Processes from $HostName at $colProcStart `r`n")
				$ExcelProcesses = $Processes | Export-Excel -Path $ExcelFile -WorksheetName 'Processes' -AutoSize -AutoFilter -TableStyle Medium6 -Append -PassThru
				$worksheet = $ExcelProcesses.Workbook.Worksheets['Processes']
				$cmdLineCol = $worksheet.Dimension.Start.Column + ($worksheet.Dimension.End.Column - $worksheet.Dimension.Start.Column) - 3
				$pathCol = $worksheet.Dimension.Start.Column + ($worksheet.Dimension.End.Column - $worksheet.Dimension.Start.Column) - 2
				Set-Column -Worksheet $worksheet -Column $cmdLineCol -Width 75
				Set-Column -Worksheet $worksheet -Column $pathCol -Width 40
				Close-ExcelPackage $ExcelProcesses
				}

			$ScheduledTasks = Get-ScheduledTask -CimSession $session | select-object -Property Date,PSComputerName,Author,Description,TaskName,TaskPath
			if ($ScheduledTasks) {
				$colSchedTasksStart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
				Write-Host "Collected Scheduled Tasks from $HostName at $colSchedTasksStart" -ForegroundColor Cyan 
				Log_Message -logfile $logfile -Message "Collected Scheduled Tasks from $HostName"
                $textboxResults.AppendText("Collected Scheduled Tasks from $HostName at $colSchedTasksStart `r`n")
				$ExcelScheduledTasks = $ScheduledTasks | Export-Excel -Path $ExcelFile -WorksheetName 'ScheduledTasks' -AutoSize -AutoFilter -TableStyle Medium6 -Append -PassThru
				$worksheett = $ExcelScheduledTasks.Workbook.Worksheets['ScheduledTasks']
				$descrCol = $worksheett.Dimension.Start.Column + ($worksheett.Dimension.End.Column - $worksheett.Dimension.Start.Column) - 2
				Set-Column -Worksheet $worksheett -Column $descrCol -Width 40
				Close-ExcelPackage $ExcelScheduledTasks
			}

			$Services = Get-CimInstance -ClassName Win32_Service -CimSession $session | select-object -Property PSComputerName,Caption,Description,Name,StartMode,PathName,ProcessId,ServiceType,StartName,State
			if ($Services) {
				$colServStart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
				Write-Host "Collected Services from $HostName at $colServStart" -ForegroundColor Cyan 
				Log_Message -logfile $logfile -Message "Collected Services from $HostName"
                $textboxResults.AppendText("Collected Services from $HostName at $colServStart `r`n")
				$ExcelServices = $Services | Export-Excel -Path $ExcelFile -WorksheetName 'Services' -AutoSize -AutoFilter -TableStyle Medium6 -Append -PassThru
				$worksheets = $ExcelServices.Workbook.Worksheets['Services']
				$descr2Col = $worksheets.Dimension.Start.Column + ($worksheets.Dimension.End.Column - $worksheets.Dimension.Start.Column) - 7
				Set-Column -Worksheet $worksheets -Column $descr2Col -Width 40
				Close-ExcelPackage $ExcelServices
			}

            $WMIConsumer = Get-CimInstance -Namespace root/subscription -ClassName __EventConsumer -CimSession $session 
            if ($WMIConsumer) {
                $colWMIContStart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Write-Host "Collected WMI Consumer Events from $HostName at $colWMIContStart" -ForegroundColor Cyan 
                Log_Message -logfile $logfile -Message "Collected WMI Consumer Events from $HostName"
                $textboxResults.AppendText("Collected WMI Consumer Events from $HostName at $colWMIContStart `r`n")
                $ExcelWMIConsumer = $WMIConsumer | Export-Excel -Path $ExcelFile -WorksheetName 'WMIConsumer' -AutoSize -AutoFilter -TableStyle Medium6 -Append -PassThru
                Close-ExcelPackage $ExcelWMIConsumer
            }

            $WMIBindings = Get-CimInstance -Namespace root/subscription -ClassName __FilterToConsumerBinding -CimSession $session 
            if ($WMIBindings) {
                $colWMIBindStart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Write-Host "Collected WMI Bindings from $HostName at $colWMIBindStart" -ForegroundColor Cyan 
                Log_Message -logfile $logfile -Message "Collected WMI Bindings from $HostName"
                $textboxResults.AppendText("Collected WMI Bindings from $HostName at $colWMIBindStart `r`n")
                $ExcelWMIBindings = $WMIBindings | Export-Excel -Path $ExcelFile -WorksheetName 'WMIBindings' -AutoSize -AutoFilter -TableStyle Medium6 -Append -PassThru
                Close-ExcelPackage $ExcelWMIBindings
            }

			$WMIFilter = Get-CimInstance -Namespace root/subscription -ClassName __EventFilter -CimSession $session 
			if ($WMIFilter) {
				$colWMIFiltStart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
				Write-Host "Collected WMI Filters from $HostName at $colWMIFiltStart" -ForegroundColor Cyan 
				Log_Message -logfile $logfile -Message "Collected WMI Filters from $HostName"
                $textboxResults.AppendText("Collected WMI Filters from $HostName at $colWMIFiltStart `r`n")
				$ExcelWMIFilter = $WMIFilter | Export-Excel -Path $ExcelFile -WorksheetName 'WMIFilter' -AutoSize -AutoFilter -TableStyle Medium6 -Append -PassThru
				Close-ExcelPackage $ExcelWMIFilter
			}

			$Shares = Get-CimInstance -ClassName Win32_Share -CimSession $session
			if ($Shares) {
				$colSharesStart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
				Write-Host "Collected Shares from $HostName at $colSharesStart" -ForegroundColor Cyan 
				Log_Message -logfile $logfile -Message "Collected Shares from $HostName"
                $textboxResults.AppendText("Collected Shares from $HostName at $colSharesStart `r`n")
				$ExcelShares = $Shares | Export-Excel -Path $ExcelFile -WorksheetName 'Shares' -AutoSize -AutoFilter -TableStyle Medium6 -Append -PassThru
				Close-ExcelPackage $ExcelShares
			}

			$ShareToDirectory = Get-CimInstance -ClassName Win32_ShareToDirectory -CimSession $session
			if ($ShareToDirectory) {
				$colShareToDirStart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
				Write-Host "Collected ShareToDirectory from $HostName at $colShareToDirStart" -ForegroundColor Cyan 
				Log_Message -logfile $logfile -Message "Collected ShareToDirectory from $HostName"
                $textboxResults.AppendText("Collected ShareToDirectory from $HostName at $colShareToDirStart `r`n")
				$ExcelShareToDirectory = $ShareToDirectory | Export-Excel -Path $ExcelFile -WorksheetName 'ShareToDirectory' -AutoSize -AutoFilter -TableStyle Medium6 -Append -PassThru
				Close-ExcelPackage $ExcelShareToDirectory
			}

			$StartupCommand = Get-CimInstance -ClassName Win32_StartupCommand -CimSession $session | select-object -Property PSComputerName,User,UserSID,Name,Command,Location
			if ($StartupCommand) {
				$colStartupCmdStart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
				Write-Host "Collected Startup Commands from $HostName at $colStartupCmdStart" -ForegroundColor Cyan 
				Log_Message -logfile $logfile -Message "Collected Startup Commands from $HostName"
                $textboxResults.AppendText("Collected Startup Commands from $HostName at $colStartupCmdStart `r`n")
				$ExcelStartupCommand = $StartupCommand | Export-Excel -Path $ExcelFile -WorksheetName 'StartupCommand' -AutoSize -AutoFilter -TableStyle Medium6 -Append -PassThru
				Close-ExcelPackage $ExcelStartupCommand
			}

			$NetworkConnections = Invoke-Command -ComputerName $Hostname -ScriptBlock { Get-NetTCPConnection } -ErrorAction SilentlyContinue | select-object -Property CreationTime,State,LocalAddress,LocalPort,OwningProcess,RemoteAddress,RemotePort
            if ($NetworkConnections) {
                $colNetConnStart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Write-Host "Collected Network Connections from $HostName at $colNetConnStart" -ForegroundColor Cyan 
                Log_Message -logfile $logfile -Message "Collected Network Connections from $HostName"
                $textboxResults.AppendText("Collected Network Connections from $HostName at $colNetConnStart `r`n")
                $ExcelNetworkConnections = $NetworkConnections | Export-Excel -Path $ExcelFile -WorksheetName 'NetworkConnections' -AutoSize -AutoFilter -TableStyle Medium6 -Append -PassThru
                Close-ExcelPackage $ExcelNetworkConnections
            }
        }
			catch {
                Write-Host "An error occurred while collecting data from $HostName $($_.Exception.Message)" -ForegroundColor Red
                Log_Message -logfile $logfile -Message "An error occurred while collecting data from $HostName $($_.Exception.Message)"
                $textboxResults.AppendText("An error occurred while collecting data from $HostName $($_.Exception.Message) `r`n")
            }            
}
}

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

function Get_PrefetchMetadata {
    param(
        [string]$HostName,
        [string]$exportPath
    )

    if ([string]::IsNullOrEmpty($HostName)) {
        Write-Host "Error: ComputerName is null or empty" -ForegroundColor Red
        return
    }

    $driveLetters = Get_RemoteDriveLetters -HostName $HostName

    $ExcelFile = Join-Path $exportPath\$HostName\RapidTriage "RapidTriage.xlsx"
    
    $copiedFilesPath = ".\CopiedFiles\$HostName\Prefetch"
    New-Item -Path $copiedFilesPath -ItemType Directory -Force

    $pecmdPath = Get-PECmdPath
    $csvOutputDir = "$exportPath\$HostName\RapidTriage\CSVOutput"
    New-Item -Path $csvOutputDir -ItemType Directory -Force

    foreach ($driveLetter in $driveLetters) {
        # Convert drive letter path to UNC format
        $prefetchPath = "\\$HostName\$driveLetter$\Windows\Prefetch\*.pf"
        $prefetchFiles = Get-ChildItem -Path $prefetchPath -ErrorAction SilentlyContinue

        foreach ($file in $prefetchFiles) {
            # Directly copy the file using UNC path
            Copy-Item -Path $file.FullName -Destination $copiedFilesPath -Force -ErrorAction SilentlyContinue
        }
    }

    # Process the copied prefetch files with PECmd
    & $pecmdPath -d $copiedFilesPath --csv $csvOutputDir

    # Collect the CSV results and import to the Excel workbook
    $csvFiles = Get-ChildItem -Path $csvOutputDir -Filter "*.csv" -File
    foreach ($csvFile in $csvFiles) {
        $csvData = Import-Csv -Path $csvFile.FullName
        $csvData | Export-Excel -Path $ExcelFile -WorksheetName $csvFile.BaseName -AutoSize -AutoFilter -FreezeFirstColumn -BoldTopRow -TableName "$($csvFile.BaseName)Table"
    }

    Remove-Item -Path $csvOutputDir -Force
    $colprestart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "Collected Prefetch from $HostName at $colprestart" -ForegroundColor Cyan 
    Log_Message -logfile $logfile -Message "Collected Prefetch from $HostName"
    $textboxResults.AppendText("Collected Prefetch from $HostName at $colprestart `r`n")
}


function Copy_BrowserHistoryFiles {
    param (
        [string]$RemoteHost,
        [string]$LocalDestination
    )

    $session = New-CimSession -ComputerName $RemoteHost -ErrorAction SilentlyContinue

    if (-not $session) {
        [System.Windows.Forms.MessageBox]::Show("Failed to connect to $RemoteHost. Please check the hostname and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $UserProfiles = Get-CimInstance -ClassName Win32_UserProfile -CimSession $session
    $driveLetters = Get_RemoteDriveLetters -HostName $RemoteHost

    if (-not (Test-Path $LocalDestination)) {
        New-Item -ItemType Directory -Path $LocalDestination | Out-Null
    }

    foreach ($UserProfile in $UserProfiles) {
        $RemoteAccount = (Split-Path $UserProfile.LocalPath -Leaf)

        foreach ($driveLetter in $driveLetters) {
            $EdgeHistoryPath = "\\$RemoteHost\$driveLetter$\Users\$RemoteAccount\AppData\Local\Microsoft\Edge\User Data\Default\History"
            $FirefoxHistoryPath = "\\$RemoteHost\$driveLetter$\Users\$RemoteAccount\AppData\Roaming\Mozilla\Firefox\Profiles\*.default*\places.sqlite"
            $ChromeHistoryPath = "\\$RemoteHost\$driveLetter$\Users\$RemoteAccount\AppData\Local\Google\Chrome\User Data\Default\History"
            $ChromeProfilesPath = "\\$RemoteHost\$driveLetter$\Users\$RemoteAccount\AppData\Local\Google\Chrome\User Data"
            $PowerShellHistoryPath = "\\$RemoteHost\$driveLetter$\Users\$RemoteAccount\AppData\Roaming\Microsoft\Windows\PowerShell\PSReadline\ConsoleHost_history.txt"
            $OperaHistoryPath = "\\$RemoteHost\$driveLetter$\Users\$RemoteAccount\AppData\Roaming\Opera Software\Opera Stable\History"
            $BraveHistoryPath = "\\$RemoteHost\$driveLetter$\Users\$RemoteAccount\AppData\Local\BraveSoftware\Brave-Browser\User Data\Default\History"
            $VivaldiHistoryPath = "\\$RemoteHost\$driveLetter$\Users\$RemoteAccount\AppData\Local\Vivaldi\User Data\Default\History"
            $EpicHistoryPath = "\\$RemoteHost\$driveLetter$\Users\$RemoteAccount\AppData\Local\Epic Privacy Browser\User Data\Default\History"
            $BraveProfilesPath = "\\$RemoteHost\$driveLetter$\Users\$RemoteAccount\AppData\Local\BraveSoftware\Brave-Browser\User Data"
            $VivaldiProfilesPath = "\\$RemoteHost\$driveLetter$\Users\$RemoteAccount\AppData\Local\Vivaldi\User Data"
            $EpicProfilesPath = "\\$RemoteHost\$driveLetter$\Users\$RemoteAccount\AppData\Local\Epic Privacy Browser\User Data"


            $UserDestination = Join-Path -Path $LocalDestination -ChildPath $RemoteAccount
            if (-not (Test-Path $UserDestination)) {
                New-Item -ItemType Directory -Path $UserDestination | Out-Null
            }

	    $ChromeProfiles = Get-ChildItem -Path $ChromeProfilesPath -Filter "Profile*" -Directory -ErrorAction SilentlyContinue | out-null

		$ChromeProfiles | ForEach-Object {
    		$ProfilePath = $_.FullName
    		$ChromeHistoryPathalt = "$ProfilePath\History"
    		$ProfileName = $_.Name
    		$DestinationPath = "$UserDestination\$ProfileName-Chrome.sqlite"
    
    		Copy-Item -Path $ChromeHistoryPathalt -Destination $DestinationPath -Force -ErrorAction SilentlyContinue | out-null
		}

        $BraveProfiles = Get-ChildItem -Path $BraveProfilesPath -Filter "Profile*" -Directory -ErrorAction SilentlyContinue | out-null

        $BraveProfiles | ForEach-Object {
    		$BProfilePath = $_.FullName
    		$BraveHistoryPathalt = "$BProfilePath\History"
    		$BProfileName = $_.Name
    		$DestinationPath = "$UserDestination\$BProfileName-Brave.sqlite"
    
    		Copy-Item -Path $BraveHistoryPathalt -Destination $DestinationPath -Force -ErrorAction SilentlyContinue | out-null
		}

        $VivaldiProfiles = Get-ChildItem -Path $VivaldiProfilesPath -Filter "Profile*" -Directory -ErrorAction SilentlyContinue | out-null

        $VivaldiProfiles | ForEach-Object {
    		$VProfilePath = $_.FullName
    		$VivaldiHistoryPathalt = "$VProfilePath\History"
    		$VProfileName = $_.Name
    		$DestinationPath = "$UserDestination\$VProfileName-Vivaldi.sqlite"
    
    		Copy-Item -Path $VivaldiHistoryPathalt -Destination $DestinationPath -Force -ErrorAction SilentlyContinue | out-null
		}

        $EpicProfiles = Get-ChildItem -Path $EpicProfilesPath -Filter "Profile*" -Directory -ErrorAction SilentlyContinue | out-null

        $EpicProfiles | ForEach-Object {
    		$EProfilePath = $_.FullName
    		$EpicHistoryPathalt = "$EProfilePath\History"
    		$EProfileName = $_.Name
    		$DestinationPath = "$UserDestination\$EProfileName-Epic.sqlite"
    
    		Copy-Item -Path $EpicHistoryPathalt -Destination $DestinationPath -Force -ErrorAction SilentlyContinue | out-null
		}	
		
            Copy-Item -Path $EdgeHistoryPath -Destination $UserDestination\Edge.sqlite -Force -ErrorAction SilentlyContinue | out-null
            Copy-Item -Path $FirefoxHistoryPath -Destination $UserDestination\FireFox.sqlite -Force -ErrorAction SilentlyContinue | out-null
            Copy-Item -Path $ChromeHistoryPath -Destination $UserDestination\Chrome.sqlite -Force -ErrorAction SilentlyContinue | out-null
            Copy-Item -Path $PowerShellHistoryPath -Destination $UserDestination\PSHistory.txt -Force -ErrorAction SilentlyContinue | out-null
            Copy-Item -Path $OperaHistoryPath -Destination $UserDestination\Opera.sqlite -Force -ErrorAction SilentlyContinue | out-null
            Copy-Item -Path $BraveHistoryPath -Destination $UserDestination\Brave.sqlite -Force -ErrorAction SilentlyContinue | out-null
            Copy-Item -Path $VivaldiHistoryPath -Destination $UserDestination\Vivaldi.sqlite -Force -ErrorAction SilentlyContinue | out-null
            Copy-Item -Path $EpicHistoryPath -Destination $UserDestination\Epic.sqlite -Force -ErrorAction SilentlyContinue | out-null
            
        }
    $histstart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "Collected browser/ps histories for $RemoteAccount at $histstart" -ForegroundColor Cyan 
    Log_Message -logfile $logfile -Message "Collected browser/ps console histories for $RemoteAccount"
    $textboxResults.AppendText("Collected browser/ps histories for $RemoteAccount at $histstart `r`n")
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