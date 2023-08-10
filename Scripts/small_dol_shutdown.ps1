$computerName = if ($comboBoxComputerName.SelectedItem) {
    $comboBoxComputerName.SelectedItem.ToString()
} else {
    $comboBoxComputerName.Text
}

if (![string]::IsNullOrEmpty($computerName)) {
    try {
        $osInstance = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $computerName -ErrorAction Stop
        $result = Invoke-CimMethod -CimInstance $osInstance -MethodName "Win32Shutdown" -Arguments @{Flags = 12} -ErrorAction Stop
        if ($result.ReturnValue -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Shutdown command sent successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $shutdown = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Write-Host "Shutdown $computerName at $shutdown" -ForegroundColor Cyan 
            Log_Message -logfile $logfile -Message "Shutdown $computerName"
            $textboxResults.AppendText("Shutdown $computerName at $shutdown")
        } else {
            [System.Windows.Forms.MessageBox]::Show("Failed to send shutdown command. Return value: $($result.ReturnValue)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error while trying to send shutdown command: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}