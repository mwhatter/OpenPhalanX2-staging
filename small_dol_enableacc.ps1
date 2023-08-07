$selectedUser = if ($comboboxUsername.SelectedItem) {
    $comboboxUsername.SelectedItem.ToString()
} else {
    $comboboxUsername.Text
}
Enable-ADAccount -Identity $selectedUser
$gotenabusr = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "$Username enabled $selectedUser at $gotenabusr" -ForegroundColor Cyan 
Log_Message -logfile $logfile -Message "$Username enabled $selectedUser"
$textboxResults.AppendText("$Username enabled $selectedUser at $gotenabusr")