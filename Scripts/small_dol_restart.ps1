$computerName = if ($comboBoxComputerName.SelectedItem) {
    $comboBoxComputerName.SelectedItem.ToString()
} else {
    $comboBoxComputerName.Text
}

if (![string]::IsNullOrEmpty($computerName)) {
    try {
        $osInstance = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $computerName -ErrorAction Stop
        $result = Invoke-CimMethod -CimInstance $osInstance -MethodName "Win32Shutdown" -Arguments @{Flags = 6} -ErrorAction Stop
        if ($result.ReturnValue -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Restart command sent successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $restart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Write-Host "Restarted $computerName at $restart" -ForegroundColor Cyan 
            Log_Message -logfile $logfile -Message "Restarted $computerName"
            $textboxResults.AppendText("Restarted $computerName at $restart")
        } else {
            [System.Windows.Forms.MessageBox]::Show("Failed to send restart command. Return value: $($result.ReturnValue)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error while trying to send restart command: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}