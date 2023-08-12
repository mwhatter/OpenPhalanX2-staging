$computerName = if ($comboBoxComputerName.SelectedItem) {$comboBoxComputerName.SelectedItem.ToString()} else {$comboBoxComputerName.Text}
If (![string]::IsNullOrEmpty($computerName)) {
    Try {
        $osInstance = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $computerName -ErrorAction Stop
        $result = Invoke-CimMethod -CimInstance $osInstance -MethodName "Win32Shutdown" -Arguments @{Flags = 6} -ErrorAction Stop
        If ($result.ReturnValue -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Restart command sent successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $restart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Write-Host "Restarted $computerName at $restart" -ForegroundColor Cyan
            $textboxResults.AppendText("Restarted $computerName at $restart`r`n")
        } else {
            [System.Windows.Forms.MessageBox]::Show("Failed to send restart command. Return value: $($result.ReturnValue)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } Catch {
        [System.Windows.Forms.MessageBox]::Show("Error while trying to send restart command: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
} else {
    [System.Windows.Forms.MessageBox]::Show("Please enter a computer name.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}
