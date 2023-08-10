$selectedUser = if ($comboboxUsername.SelectedItem) {
    $comboboxUsername.SelectedItem.ToString()
} else {
    $comboboxUsername.Text
}
$computerName = if ($comboBoxComputerName.SelectedItem) {
    $comboBoxComputerName.SelectedItem.ToString()
} else {
    $comboBoxComputerName.Text
}
Invoke-Command -ComputerName $computerName -ScriptBlock {
    $username = $args[0]
    $userSessions = (Get-CimInstance -ClassName Win32_ComputerSystem).UserName
    foreach ($userSession in $userSessions) {
        if ($userSession -eq $username) {
            (Get-CimInstance -ClassName Win32_OperatingSystem -EnableAllPrivileges).Win32Shutdown(4)
        }
    }
} -ArgumentList $selectedUser
$gotlogoffusr = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "$Username logged $selectedUser off from $computerName at $gotlogoffusr" -ForegroundColor Cyan 
Log_Message -logfile $logfile -Message "$Username logged $selectedUser off from $computerName"
$textboxResults.AppendText("$Username logged $selectedUser off from $computerName at $gotlogoffusr")