$selectedUser = if ($comboboxUsername.SelectedItem) {
    $comboboxUsername.SelectedItem.ToString()
} else {
    $comboboxUsername.Text
}
Set-ADUser -Identity $selectedUser -ChangePasswordAtLogon $true
$gotpwchng = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "$Username forced a password change for $selectedUser at $gotpwchng" -ForegroundColor Cyan 
Log_Message -logfile $logfile -Message "$Username forced a password change for $selectedUser"
$textboxResults.AppendText("$Username forced a password change for $selectedUser at $gotpwchng")