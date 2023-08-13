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

$result = Invoke-Command -ComputerName $computerName -ScriptBlock {
    & query user
}
$result
