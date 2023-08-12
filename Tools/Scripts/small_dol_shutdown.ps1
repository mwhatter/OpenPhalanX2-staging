$computerName = if ($comboBoxComputerName.SelectedItem) {$comboBoxComputerName.SelectedItem.ToString()} else {$comboBoxComputerName.Text}
try {
    $osInstance = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $computerName -ErrorAction Stop
    $result = Invoke-CimMethod -CimInstance $osInstance -MethodName "Win32Shutdown" -Arguments @{Flags = 12} -ErrorAction Stop

    if ($result.ReturnValue -eq 0) {
        $shutdown = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $textboxResults.AppendText("Shutdown $computerName at $shutdown`r`n")
    } else {
        [System.Windows.Forms.MessageBox]::Show("Failed to send shutdown command. Return value: $($result.ReturnValue)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
} catch {
    [System.Windows.Forms.MessageBox]::Show("Error while trying to send shutdown command: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}
