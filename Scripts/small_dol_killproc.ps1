$computerName = if ($comboBoxComputerName.SelectedItem) {
    $comboBoxComputerName.SelectedItem.ToString()
} else {
    $comboBoxComputerName.Text
}

$selectedProcess = if ($dropdownProcessId.SelectedItem) {
    $dropdownProcessId.SelectedItem.ToString()
} else {
    $dropdownProcessId.Text
}

if (![string]::IsNullOrEmpty($computerName) -and ![string]::IsNullOrEmpty($selectedProcess)) {
    $processId = $selectedProcess.split()[0]
    $process = Get-CimInstance -ClassName Win32_Process -Filter "ProcessId = $processId" -ComputerName $computerName -ErrorAction SilentlyContinue
    if ($process) {
        try {
            $result = Invoke-CimMethod -CimInstance $process -MethodName "Terminate" -ErrorAction Stop
            if ($result.ReturnValue -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Process terminated successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                $prockill = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Write-Host "Terminated $selectedProcess from $computerName at $prockill" -ForegroundColor Cyan 
                Log_Message -logfile $logfile -Message "Terminated $selectedProcess from $computerName"
            } else {
                [System.Windows.Forms.MessageBox]::Show("Failed to terminate the process. Return value: $($result.ReturnValue)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error while trying to terminate the process: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Process not found on the specified computer.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}