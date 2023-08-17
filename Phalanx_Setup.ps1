<#
Run this script first. It downloads and installs a variety of security tools, including 
DeepBlueCLI, Hayabusa, Sysmon, and Eric Zimmerman's tools. You'll also get directories built 
to manage tools, their respective files, and outputs for the many functions of OpenPhalanX.
#>

Add-Type -AssemblyName System.Windows.Forms

# Check if running as Administrator
function Test-IsAdmin {
    return ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}

if (-not (Test-IsAdmin)) {
    Write-Host "This script needs to be run as an Administrator. Please run the script as Administrator." -ForegroundColor Red
    exit
}

$currentDirName = Split-Path -Path (Get-Location) -Leaf

if ($currentDirName -ne "OpenPhalanX") {
    if (-not (Test-Path ".\OpenPhalanX")) {
        New-Item -ItemType Directory -Path ".\OpenPhalanX" | Out-Null
    }
    Set-Location ".\OpenPhalanX"
}

$directories = @(
    ".\CopiedFiles",
    ".\Logs",
    ".\Logs\EVTX",
    ".\Logs\Reports",
    ".\Logs\Audit",
    ".\Tools\Hayabusa",
    ".\Tools\Sysmon",
    ".\Tools\EZTools",
    ".\Tools\Scripts\Integrations\Inuse",
    ".\Tools\Scripts\Integrations\Stored"
)

$errors = @()

################## Workspace ##################
foreach ($directory in $directories) {
    try {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
        Write-Host "Directory $directory created successfully."
    } catch {
        $errors += [PSCustomObject]@{
            Step  = "Creating $directory"
            Error = $_.Exception.Message
        }
    }
}

function DownloadFilesFromRepo {
    param (
        [string]$apiUrl,
        [string]$localPath,
        [string]$filterPath = $null
    )

    $files = Invoke-RestMethod -Uri $apiUrl

    foreach ($file in $files.tree) {
        if ($file.type -eq "blob" -and ($file.path.StartsWith($filterPath) -or ($filterPath -eq $null))) {
            $rawUrl = "https://raw.githubusercontent.com/mwhatter/OpenPhalanX2-staging/main/$($file.path)"
            
            # Maintain the relative directory structure
            $relativePath = if ($filterPath) { $file.path.Replace($filterPath, "") } else { $file.path }
            $destination = Join-Path $localPath $relativePath
            
            # Create directories if they don't exist
            $destinationDir = Split-Path -Path $destination -Parent
            if (-not (Test-Path $destinationDir)) {
                New-Item -ItemType Directory -Path $destinationDir | Out-Null
            }
            
            try {
                Invoke-WebRequest -Uri $rawUrl -OutFile $destination
            } catch {
                $errors += [PSCustomObject]@{
                    Step  = "Downloading file $rawUrl"
                    Error = $_.Exception.Message
                }
            }
        }
    }
}

$scriptsPath = ".\Tools\Scripts"
if ((Get-ChildItem -Path $scriptsPath -File).Count -gt 0) {
    Write-Host "Files found in the $scriptsPath directory. Skipping resource downloads."
} else {
    # Call the DownloadFilesFromRepo function for all desired directories
    DownloadFilesFromRepo -apiUrl "https://api.github.com/repos/mwhatter/OpenPhalanX2-staging/git/trees/main?recursive=1" -localPath ".\" -filterPath ""
    DownloadFilesFromRepo -apiUrl "https://api.github.com/repos/mwhatter/OpenPhalanX2-staging/git/trees/main?recursive=1" -localPath ".\Tools\Scripts" -filterPath "Tools/Scripts/"
    DownloadFilesFromRepo -apiUrl "https://api.github.com/repos/mwhatter/OpenPhalanX2-staging/git/trees/main?recursive=1" -localPath ".\Tools\Scripts\Integrations\Stored" -filterPath "Tools/Scripts/Integrations/Stored/"
}

unblock-file Defending_Off_the_land.ps1
unblock-file API_Setup.ps1

# Unblock all files in .\Tools\Scripts directory recursively
Get-ChildItem -Path ".\Tools\Scripts" -Recurse | ForEach-Object {
    if ($_.Attributes -ne "Directory") {
        Unblock-File -Path $_.FullName
    }
}

# Unblock all files in .\Tools\Scripts\Integrations\Stored directory recursively
Get-ChildItem -Path ".\Tools\Scripts\Integrations\Stored" -Recurse | ForEach-Object {
    if ($_.Attributes -ne "Directory") {
        Unblock-File -Path $_.FullName
    }
}

function PromptForInstallation {
    param (
        [string]$appName,
        [string]$checkPath = $null
    )

    $title = "Setup $appName"
    if ($checkPath -and (Test-Path $checkPath)) {
        $message = "$appName already exists! Overwrite?"
    } else {
        $message = "Do you want to download and setup $appName"
    }

    $messageBoxResult = [System.Windows.Forms.MessageBox]::Show($message, $title, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    
    return $messageBoxResult -eq 'Yes'
}

################## Custom Installation Prompt ##################
$title = "Installation Customization"
$message = "Do you want to customize your installation?"
$customizeResult = [System.Windows.Forms.MessageBox]::Show($message, $title, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

function ShouldInstall {
    param (
        [string]$appName,
        [string]$checkPath = $null
    )

    # If user opted for non-customized setup and the component isn't installed, install it
    if (-not $customizeSetup -and -not (Test-Path $checkPath)) {
        return $true
    }

    $title = "Setup $appName"
    if ($checkPath -and (Test-Path $checkPath)) {
        $message = "$appName already exists! Overwrite?"
        $messageBoxResult = [System.Windows.Forms.MessageBox]::Show($message, $title, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        return $messageBoxResult -eq 'Yes'
    } else {
        return $true
    }
}

################## Downloads ##################

# PowerShell 7
if (ShouldInstall -appName "PowerShell 7" -checkPath "$env:ProgramFiles\PowerShell\7\pwsh.exe") {
    $downloadUrl = "https://github.com/PowerShell/PowerShell/releases/download/v7.3.5/PowerShell-7.3.5-win-x64.msi"
    $fileName = $downloadUrl -split '/' | Select-Object -Last 1
    $savePath = Join-Path -Path $env:TEMP -ChildPath $fileName

    try {
        Invoke-WebRequest -Uri $downloadUrl -OutFile $savePath
        Write-Host "PowerShell 7 MSI downloaded successfully."
        Start-Process -FilePath msiexec.exe -ArgumentList "/i `"$savePath`" /qn" -Wait
        Remove-Item -Path $savePath
        Write-Host "PowerShell 7 installed successfully."
    } catch {
        $errors += [PSCustomObject]@{
            Step  = "Installing PowerShell 7"
            Error = $_.Exception.Message
        }
    }
}

# VSCode
if (ShouldInstall -appName "VSCode" -checkPath "C:\Users\*\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Visual Studio Code") {
    $vscodeDownloadUrl = "https://update.code.visualstudio.com/latest/win32-x64-user/stable"
    $vscodeInstallerPath = Join-Path -Path $env:TEMP -ChildPath "VSCodeInstaller.exe"
    
    try {
        # Download VSCode Installer
        Invoke-WebRequest -Uri $vscodeDownloadUrl -OutFile $vscodeInstallerPath
        Write-Host "VSCode Installer downloaded successfully."
        
        # Silent Install
        Start-Process -FilePath $vscodeInstallerPath -ArgumentList "/silent", "/mergetasks=!runcode" -Wait
        Remove-Item -Path $vscodeInstallerPath
        Write-Host "VSCode installed successfully."
    } catch {
        $errors += [PSCustomObject]@{
            Step  = "Installing VSCode"
            Error = $_.Exception.Message
        }
    }
}

# Python
if (ShouldInstall -appName "Python" -checkPath "C:\Python\python.exe") {
    $downloadUrl = "https://www.python.org/ftp/python/3.11.4/python-3.11.4-amd64.exe"
    $installerPath = Join-Path -Path $env:TEMP -ChildPath "python_installer.exe"

    try {
        Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath
        Start-Process -FilePath $installerPath -ArgumentList "/quiet", "TargetDir=C:\Python", "AddToPath=1", "AssociateFiles=1", "Shortcuts=1" -Wait
        Remove-Item -Path $installerPath
        Write-Host "Python installed successfully."
        
        # Install numscrypt Python module after successfully installing Python
        try {
            Start-Process -FilePath "C:\Python\python.exe" -ArgumentList "-m pip install numscrypt" -Wait
            Write-Host "numscrypt Python module installed successfully."
        } catch {
            $errors += [PSCustomObject]@{
                Step  = "Installing numscrypt Python module"
                Error = $_.Exception.Message
            }
        }

    } catch {
        $errors += [PSCustomObject]@{
            Step  = "Installing Python"
            Error = $_.Exception.Message
        }
    }
}

# DeepBlueCLI
if (ShouldInstall -appName "DeepBlueCLI" -checkPath ".\Tools\DeepBlueCLI\DeepBlue.ps1") {
    try {
        $deepBlueCLIZip = Invoke-WebRequest -URI "https://github.com/sans-blue-team/DeepBlueCLI/archive/refs/heads/master.zip" -OutFile "DeepBlueCLI.zip"
        Expand-Archive "DeepBlueCLI.zip" -DestinationPath ".\Tools\" -Force
        Remove-Item "DeepBlueCLI.zip"
        if (Test-Path .\Tools\DeepBlueCLI) {
            Remove-Item .\Tools\DeepBlueCLI -Recurse -Force
        }
        Rename-Item .\Tools\DeepBlueCLI-master -NewName DeepBlueCLI -Force | Out-Null
        Write-Host "DeepBlueCLI installed successfully."
    } catch {
        $errors += [PSCustomObject]@{
            Step  = "Installing DeepBlueCLI"
            Error = $deepBlueCLIZip.Exception.Message
        }
    }
}

# Hayabusa
if (ShouldInstall -appName "Hayabusa" -checkPath ".\Tools\Hayabusa\hayabusa.exe") {
    $hayabusaApiUrl = "https://api.github.com/repos/Yamato-Security/hayabusa/releases/latest"
    $hayabusaReleaseData = Invoke-RestMethod -Uri $hayabusaApiUrl
    $hayabusaDownloadUrl = $hayabusaReleaseData.assets | Where-Object { $_.browser_download_url -like "*-win-64-bit.zip" } | Select-Object -First 1 -ExpandProperty browser_download_url
    $hayabusaZipPath = Join-Path -Path ".\Tools" -ChildPath "hayabusa-win-x64.zip"
    $hayabusaExtractPath = ".\Tools\Hayabusa"

    # Download Hayabusa
    try {
        Invoke-WebRequest -Uri $hayabusaDownloadUrl -OutFile $hayabusaZipPath
        Write-Host "Hayabusa downloaded successfully."
    } catch {
        $errors += [PSCustomObject]@{
            Step  = "Downloading Hayabusa"
            Error = $_.Exception.Message
        }
    }

    # Extract Hayabusa
    try {
        Expand-Archive -Path $hayabusaZipPath -DestinationPath $hayabusaExtractPath -Force
        Write-Host "Hayabusa extracted successfully."
    } catch {
        $errors += [PSCustomObject]@{
            Step  = "Extracting Hayabusa"
            Error = $_.Exception.Message
        }
    }

    # Rename Hayabusa executable
    $hayabusaExecutable = Get-ChildItem -Path $hayabusaExtractPath -Filter "hayabusa-*" -Recurse -File | Select-Object -First 1
    remove-item $CWD\OpenPhalanX\Tools\Hayabusa\hayabusa.exe -ErrorAction SilentlyContinue
    $hayabusaExecutable | Rename-Item -NewName "hayabusa.exe" -Force | Out-Null
}

# Sysmon
if (ShouldInstall -appName "Sysmon" -checkPath ".\Tools\Sysmon\Sysmon64.exe") {
    $sysmonUrl = "https://download.sysinternals.com/files/Sysmon.zip"
    $sysmonconfigurl = "https://raw.githubusercontent.com/olafhartong/sysmon-modular/master/sysmonconfig.xml"
    $sysmonZip = ".\Tools\Sysmon.zip"
    $sysmonPath = ".\Tools\Sysmon"
    $sysmonConfigPath = ".\Tools\Sysmon\sysmonconfig.xml"

    try {
        Invoke-WebRequest -URI $sysmonUrl -OutFile $sysmonZip
        Expand-Archive $sysmonZip -DestinationPath $sysmonPath -Force
        Remove-Item $sysmonZip
        Invoke-WebRequest -URI $sysmonconfigurl -OutFile $sysmonConfigPath
        Write-Host "Sysmon and sysmonconfig downloaded and installed successfully."
    } catch {
        $errors += [PSCustomObject]@{
            Step  = "Installing Sysmon"
            Error = $_.Exception.Message
        }
    }
}

# Get-ZimmermanTools
if (ShouldInstall -appName "Get-ZimmermanTools" -checkPath ".\Tools\EZTools\Get-ZimmermanTools.ps1") {
    try {
        $zimmermanToolsZip = Invoke-WebRequest -URI "https://raw.githubusercontent.com/EricZimmerman/Get-ZimmermanTools/master/Get-ZimmermanTools.ps1" -OutFile ".\Tools\EZTools\Get-ZimmermanTools.ps1"
        Unblock-File ".\Tools\EZTools\Get-ZimmermanTools.ps1"
        .\Tools\EZTools\Get-ZimmermanTools.ps1 -Dest ".\Tools\EZTools" -NetVersion 4 -Verbose:$false *> $null
        Write-Host "Get-ZimmermanTools installed successfully."
    } catch {
        $errors += [PSCustomObject]@{
            Step  = "Installing Get-ZimmermanTools"
            Error = $zimmermanToolsZip.Exception.Message
        }
    }
}

################## Errors ##################
if ($errors.Count -gt 0) {
    Write-Host "Following errors occurred during the process:"
    $errors | ForEach-Object { Write-Host "$($_.Step): $($_.Error)" }
} else {
    Write-Host "All operations completed successfully."
}
