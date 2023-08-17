# Function to download files from GitHub
function DownloadFiles($url, $destination) {
    $files = Invoke-RestMethod -Uri $url
    foreach ($file in $files) {
        $downloadUrl = $file.download_url
        $fileName = $file.name
        Invoke-WebRequest -Uri $downloadUrl -OutFile "$destination\$fileName"
    }
}

# Prompt user for OpenPhalanX main scripts update
$openPhalanxUpdate = Read-Host "Do you want to update the OpenPhalanX main scripts? (yes/no)"
if ($openPhalanxUpdate -eq "yes") {
    $openPhalanxUrl = "https://api.github.com/repos/mwhatter/OpenPhalanX2-staging/contents?ref=main"
    DownloadFiles -url $openPhalanxUrl -destination ".\"
    Write-Host "OpenPhalanX main scripts updated successfully!"
}

# Prompt user for button click scripts update
$buttonClickUpdate = Read-Host "Do you want to update the button click scripts? (yes/no)"
if ($buttonClickUpdate -eq "yes") {
    $buttonClickUrl = "https://api.github.com/repos/mwhatter/OpenPhalanX2-staging/contents/Tools/Scripts?ref=main"
    DownloadFiles -url $buttonClickUrl -destination ".\Tools\Scripts"
    Write-Host "Button click scripts updated successfully!"
}

# Prompt user for Python API helper scripts update
$pythonApiUpdate = Read-Host "Do you want to update the Python API helper scripts? (yes/no)"
if ($pythonApiUpdate -eq "yes") {
    $pythonApiUrl = "https://api.github.com/repos/mwhatter/OpenPhalanX2-staging/contents/Tools/Scripts/Integrations/Stored?ref=main"
    DownloadFiles -url $pythonApiUrl -destination ".\Tools\Scripts\Integrations\Stored"
    Write-Host "Python API helper scripts updated successfully!"
}
