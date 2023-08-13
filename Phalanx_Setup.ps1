<#
Run this script first. It downloads and installs a variety of security tools, including 
DeepBlueCLI, Hayabusa, Sysmon, Eric Zimmerman's tools. You'll also get directories built 
to manage tools, their respective files, and outputs for the many funcitons of OpenPhalanX.
#>
Add-Type -AssemblyName System.Windows.Forms

$directories = @(
    ".\CopiedFiles",
    ".\Logs",
    ".\Logs\EVTX",
    ".\Logs\Reports",
    ".\Logs\Audit",
    ".\Tools\Hayabusa",
    ".\Tools\Sysmon",
    ".\Tools\EZTools"
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

################## Downloads ##################
# PowerShell 7
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

function PromptForInstallation($appName) {
    $messageBoxResult = [System.Windows.Forms.MessageBox]::Show("Do you want to download and install $appName?", "Install $appName", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    return $messageBoxResult -eq 'Yes'
}

# Python
if (PromptForInstallation("Python")) {
    $downloadUrl = "https://www.python.org/ftp/python/3.11.4/python-3.11.4-amd64.exe"
    $installerPath = Join-Path -Path c:\Test -ChildPath "python_installer.exe"

    try {
        Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath
        Start-Process -FilePath $installerPath -ArgumentList "/quiet", "TargetDir=C:\Python", "AddToPath=1", "AssociateFiles=1", "Shortcuts=1" -Wait
        Remove-Item -Path $installerPath
        Write-Host "Python installed successfully."
    } catch {
        $errors += [PSCustomObject]@{
            Step  = "Installing Python"
            Error = $_.Exception.Message
        }
    }
}

# numscrypt Python module
if (PromptForInstallation("numscrypt Python module")) {
    try {
        Start-Process -FilePath python.exe -ArgumentList "-m pip install numscrypt" -Wait
        Write-Host "numscrypt Python module installed successfully."
    } catch {
        $errors += [PSCustomObject]@{
            Step  = "Installing numscrypt Python module"
            Error = $_.Exception.Message
        }
    }
}

# DeepBlueCLI
if (PromptForInstallation("DeepBlueCLI")) {
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
            Error = $_.Exception.Message
        }
    }
}

# Hayabusa
if (PromptForInstallation("Hayabusa")) {
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
    $hayabusaExecutable | Rename-Item -NewName "hayabusa.exe" -Force | Out-Null

    # Update Hayabusa rules
    $rulesPath = ".Tools\Hayabusa\Rules"
    $hayabusaExecutablePath = Join-Path -Path $hayabusaExtractPath -ChildPath "hayabusa.exe"

    try {
        & $hayabusaExecutablePath update-rules -r $rulesPath | Out-Null
        Write-Host "Hayabusa rules updated successfully."
        Remove-Item -Path $hayabusaZipPath -Force
    } catch {
        $errors += [PSCustomObject]@{
            Step  = "Updating Hayabusa rules"
            Error = $_.Exception.Message
        }
    }
}

# Sysmon
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

# Get-ZimmermanTools
if (PromptForInstallation("Get-ZimmermanTools")) {
    $zimmermanToolsZip = Invoke-WebRequest -URI "https://raw.githubusercontent.com/EricZimmerman/Get-ZimmermanTools/master/Get-ZimmermanTools.ps1" -OutFile ".\Tools\EZTools\Get-ZimmermanTools.ps1"
    try {
        unblock-file ".\Tools\EZTools\Get-ZimmermanTools.ps1"
        .\Tools\EZTools\Get-ZimmermanTools.ps1 -Dest ".\Tools\EZTools" -NetVersion 4 -Verbose:$false *> $null
        Write-Host "Get-ZimmermanTools installed successfully."
    } catch {
        $errors += [PSCustomObject]@{
            Step  = "Installing Get-ZimmermanTools"
            Error = $_.Exception.Message
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
