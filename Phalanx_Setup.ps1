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

# Initialize an array to store errors
$errors = @()

################## Workspace ##################
foreach ($directory in $directories) {
    try {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
        Write-Host "Directory $directory created successfully."
    } catch {
        $errors += [PSCustomObject]@{
            'Step'   = "Creating $directory"
            'Error'  = $_.Exception.Message
        }
    }
}

################## Downloads ##################
# Define the URL for the PowerShell 7 MSI installer
$downloadUrl = "https://github.com/PowerShell/PowerShell/releases/download/v7.3.5/PowerShell-7.3.5-win-x64.msi"

# Extract the filename from the download URL
$fileName = $downloadUrl -split '/' | Select-Object -Last 1

# Define the path to save the MSI file
$savePath = Join-Path -Path $env:TEMP -ChildPath $fileName

try {
    # Download the MSI file
    Invoke-WebRequest -Uri $downloadUrl -OutFile $savePath

    Write-Host "PowerShell 7 MSI download successful."

    # Install PowerShell 7 using the downloaded MSI file
    Start-Process -FilePath msiexec.exe -ArgumentList "/i `"$savePath`" /qn" -Wait

    # Clean up the temporary MSI file
    Remove-Item -Path $savePath

    Write-Host "PowerShell 7 installation successful."
} catch {
    $installError = $_
    Write-Host "An error occurred during installation: $installError"
}

# Define the URL for the Python download page
$downloadUrl = "https://www.python.org/ftp/python/3.11.4/python-3.11.4-amd64.exe"

# Retrieve the download page
$response = Invoke-WebRequest -Uri $downloadUrl

# Extract the download URL for the latest stable release of Python 3
if ($response) {
    # Define the path to save the installer
    $installerPath = Join-Path -Path c:\Test -ChildPath "python_installer.exe"

    # Download the installer
    Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath

    # Install Python for all users
    Start-Process -FilePath $installerPath -ArgumentList "/quiet", "TargetDir=C:\Python", "AddToPath=1", "AssociateFiles=1", "Shortcuts=1" -Wait

    # Clean up the installer
    Remove-Item -Path $installerPath
    Write-Host "Python downloaded and installed successfully."
}
else {
    Write-Host "Unable to retrieve the download link for Python."
    $downloadError = $_
}
 
# Install the numscrypt Python module
try {
    Start-Process -FilePath python.exe -ArgumentList "-m pip install numscrypt" -Wait
    Write-Host "numscrypt Python module installation successful."
} catch {
    $installError = $_
}

# Install the vmray-rest-api Python module
try {
    Start-Process -FilePath python.exe -ArgumentList "-m pip install vmray-rest-api" -Wait
    Write-Host "vmray-rest-api Python module installation successful."
} catch {
    $installError = $_
}

################## Download and unzip DeepBlueCLI repository ##################
try {
    $deepBlueCLIZip = Invoke-WebRequest -URI https://github.com/sans-blue-team/DeepBlueCLI/archive/refs/heads/master.zip -OutFile DeepBlueCLI.zip
    Expand-Archive DeepBlueCLI.zip -DestinationPath ".\Tools\" -Force
    Remove-Item DeepBlueCLI.zip

    # Remove existing DeepBlueCLI directory if it exists
    if (Test-Path .\Tools\DeepBlueCLI) {
        Remove-Item .\Tools\DeepBlueCLI -Recurse -Force
    }

    Rename-Item .\Tools\DeepBlueCLI-master -NewName DeepBlueCLI -Force | Out-Null

    Write-Host "DeepBlueCLI download and unzip successful."
} catch {
    Write-Host "DeepBlueCLI download and unzip failed:`n$_"
    $downloadError = $_
}

################## Get the latest Hayabusa release page ##################
$hayabusaApiUrl = "https://api.github.com/repos/Yamato-Security/hayabusa/releases/latest"
$hayabusaReleaseData = Invoke-RestMethod -Uri $hayabusaApiUrl

# Extract the download URL for the latest Hayabusa release
$hayabusaDownloadUrl = $hayabusaReleaseData.assets |
                       Where-Object { $_.browser_download_url -like "*-win-64-bit.zip" } |
                       Select-Object -First 1 -ExpandProperty browser_download_url

# Prepare the local paths for downloading and extracting the Hayabusa zip file
$hayabusaZipPath = Join-Path -Path ".\Tools" -ChildPath "hayabusa-win-x64.zip"
$hayabusaExtractPath = ".\Tools\Hayabusa"

################## Download and extract Hayabusa ##################
try {
    Invoke-WebRequest -Uri $hayabusaDownloadUrl -OutFile $hayabusaZipPath
    Write-Host "Hayabusa download successful."
} catch {
    Write-Host "Hayabusa download failed:`n$_"
    $downloadError = $_
}

################## Extract the Hayabusa files ##################
try {
    Expand-Archive -Path $hayabusaZipPath -DestinationPath $hayabusaExtractPath -Force
    Write-Host "Hayabusa files extracted successfully."
} catch {
    Write-Host "Hayabusa files extraction failed:`n$_"
}

################## Rename the executable to hayabusa.exe ##################
$hayabusaExecutable = Get-ChildItem -Path $hayabusaExtractPath -Filter "hayabusa-*" -Recurse -File | Select-Object -First 1
$hayabusaExecutable | Rename-Item -NewName "hayabusa.exe" -Force | Out-Null

################## Update the Hayabusa rules ##################
$rulesPath = ".Tools\Hayabusa\Rules"
$hayabusaExecutablePath = Join-Path -Path $hayabusaExtractPath -ChildPath "hayabusa.exe"

try {
    & $hayabusaExecutablePath update-rules -r $rulesPath | Out-Null
    Write-Host "Hayabusa rules updated successfully."

    # Clean up the downloaded zip file
    Remove-Item -Path $hayabusaZipPath -Force
} catch {
    Write-Host "Hayabusa rules update failed:`n$_"
}

################## Download and unzip Sysmon ##################
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

    Write-Host "Sysmon and sysmonconfig download and unzip successful."
} catch {
    Write-Host "Sysmon and sysmonconfig download and unzip failed:`n$_"
    $downloadError = $_
}

################## Download and execute Get-ZimmermanTools to retrieve Eric Zimmerman's toolset ##################
$zimmermanToolsZip = Invoke-WebRequest -URI "https://raw.githubusercontent.com/EricZimmerman/Get-ZimmermanTools/master/Get-ZimmermanTools.ps1" -OutFile ".\Tools\EZTools\Get-ZimmermanTools.ps1"

try {
    unblock-file ".\Tools\EZTools\Get-ZimmermanTools.ps1"
    .\Tools\EZTools\Get-ZimmermanTools.ps1 -Dest ".\Tools\EZTools" -NetVersion 4 -Verbose:$false *> $null
    Write-Host "Get-ZimmermanTools download and execution successful."
} catch {
    Write-Host "Get-ZimmermanTools download and execution failed:`n$_"
    $downloadError = $_
}

################## Errors ##################
if ($errors.Count -gt 0) {
    Write-Host "The following errors occurred:"
    foreach ($error in $errors) {
        Write-Host ("Step: " + $error.Step)
        Write-Host ("Error: " + $error.Error)
        Write-Host "------------------------------------"
    }
} else {
    Write-Host "All downloads and installations successful."
}
