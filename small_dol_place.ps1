$computerName = if ($comboBoxComputerName.SelectedItem) {
    $comboBoxComputerName.SelectedItem.ToString()
} else {
    $comboBoxComputerName.Text
}
$selectedFile = if ($comboboxlocalFilePath.SelectedItem) {
    $comboboxlocalFilePath.SelectedItem.ToString()
} else {
    $comboboxlocalFilePath.Text
}
$remoteFilePath = $textboxRemoteFilePath.Text

if (![string]::IsNullOrEmpty($computerName) -and ![string]::IsNullOrEmpty($selectedFile) -and ![string]::IsNullOrEmpty($remoteFilePath)) {
    $filename = [System.IO.Path]::GetFileName($selectedFile)
    $remoteDriveLetter = $remoteFilePath.Substring(0, 1)
    $remoteUncPath = "\\$computerName\$remoteDriveLetter$" + $remoteFilePath.Substring(2)

    try {
        Copy-Item -Path $selectedFile -Destination $remoteUncPath -Force
        [System.Windows.Forms.MessageBox]::Show("File '$filename' copied to remote directory.", "Place File Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $textboxResults.AppendText("File '$filename' copied to remote directory.`r`n")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error copying file '$filename' to remote directory: $($_.Exception.Message)", "Place File Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $textboxResults.AppendText("Error copying file '$filename' to remote directory: $($_.Exception.Message)`r`n")
        $fileput = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Write-Host "Copied $selectedFile from localhost to $remoteFilePath on $computerName at $fileput" -ForegroundColor Cyan 
        Log_Message -logfile $logfile -Message "Copied $selectedFile from localhost to $remoteFilePath on $computerName"
    }
}