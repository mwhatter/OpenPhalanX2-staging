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
    $sessions = query user | Where-Object { $_ -match "\s+(\d+)\s+($username)" } | ForEach-Object { if ($matches[1]) { $matches[1] } }
    foreach ($sessionId in $sessions) {
        logoff $sessionId
    }
} -ArgumentList $selectedUser

$gotlogoffusr = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "$Username logged $selectedUser off from $computerName at $gotlogoffusr" -ForegroundColor Cyan 
Log_Message -logfile $logfile -Message "$Username logged $selectedUser off from $computerName"
$textboxResults.AppendText("$Username logged $selectedUser off from $computerName at $gotlogoffusr`r`n")
