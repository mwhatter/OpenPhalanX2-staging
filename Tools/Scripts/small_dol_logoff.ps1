$selectedUser = if ($comboboxUsername.SelectedItem) {
    $comboboxUsername.SelectedItem.ToString().split('\')[-1]  # Extract username without domain
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
    $sessions = query user 2>&1
    foreach ($session in $sessions) {
        if ($session -match "^\s*(\d+)\s+($username)\s+") {
            $sessionId = $matches[1]
            logoff $sessionId
        }
    }
} -ArgumentList $selectedUser

$gotlogoffusr = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "$Username logged $selectedUser off from $computerName at $gotlogoffusr" -ForegroundColor Cyan 
Log_Message -logfile $logfile -Message "$Username logged $selectedUser off from $computerName"
$textboxResults.AppendText("$Username logged $selectedUser off from $computerName at $gotlogoffusr`r`n")
