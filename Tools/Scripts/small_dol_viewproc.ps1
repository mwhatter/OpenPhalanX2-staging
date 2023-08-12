$computerName = if ($comboBoxComputerName.SelectedItem) {
    $comboBoxComputerName.SelectedItem.ToString()
} else {$comboBoxComputerName.Text}

try {
    $processes = Get-CimInstance -ClassName Win32_Process -ComputerName $computerName -ErrorAction Stop | Select-Object ProcessId, Name, CommandLine

    $dropdownProcessId.Items.Clear()
    foreach ($process in $processes) {
        $dropdownProcessId.Items.Add("$($process.ProcessId) - $($process.Name)")
    }
    $processesString = $processes | Out-String
    $textboxResults.AppendText($processesString)

    $gotproc = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "Processes listed from $computerName at $gotproc" -ForegroundColor Cyan 
} catch {
    [System.Windows.Forms.MessageBox]::Show("Error while trying to retrieve processes: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}
