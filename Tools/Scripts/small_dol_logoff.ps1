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

# Get the session ID for the selected user
$sessionData = Invoke-Command -ComputerName $computerName -ScriptBlock {
    $output = & query user $args[0]
    $sessionLine = $output | Where-Object { $_ -like "*$args[0]*" }
    $sessionId = ($sessionLine -split "\s+")[2]
    return $sessionId
} -ArgumentList $selectedUser

# Log off the user with the retrieved session ID
if ($sessionData) {
    Invoke-Command -ComputerName $computerName -ScriptBlock {
        & logoff $args[0]
    } -ArgumentList $sessionData
    Write-Host "$selectedUser has been logged off from $computerName"
} else {
    Write-Host "Could not find session ID for user $selectedUser on $computerName"
}
