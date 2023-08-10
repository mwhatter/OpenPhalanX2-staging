$selectedUser = if ($comboboxUsername.SelectedItem) {
    $comboboxUsername.SelectedItem.ToString()
} else {
    $comboboxUsername.Text
}
Disable-ADAccount -Identity $selectedUser  
$gotdisabusr = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "$Username disabled $selectedUser at $gotdisabusr" -ForegroundColor Cyan 
Log_Message -logfile $logfile -Message "$Username disabled $selectedUser"
$textboxResults.AppendText("$Username disabled $selectedUser at $gotdisabusr")