$computerName = if ($comboBoxComputerName.SelectedItem) {
    $comboBoxComputerName.SelectedItem.ToString()
} else {
    $comboBoxComputerName.Text
}

if (![string]::IsNullOrEmpty($computerName)) {
    try {
        # Set the local paths for Sysmon and the configuration file
        $sysmonPath = "$CWD\Tools\Sysmon\Sysmon64.exe"
        $configPath = "$CWD\Tools\Sysmon\sysmonconfig.xml"

        # Copy Sysmon and the config to the remote computer
        $remoteSysmonPath = "\\$computerName\C$\Windows\Temp\Sysmon64.exe"
        $remoteConfigPath = "\\$computerName\C$\Windows\Temp\sysmonconfig-export.xml"

        Copy-Item -Path $sysmonPath -Destination $remoteSysmonPath -Force
        Copy-Item -Path $configPath -Destination $remoteConfigPath -Force

        # Deploy and run Sysmon on the remote computer
        Invoke-Command -ComputerName $computerName -ScriptBlock {
            $sysmonPath = "C:\Windows\Temp\Sysmon64.exe"
            $configPath = "C:\Windows\Temp\sysmonconfig-export.xml"

            Start-Process -FilePath $sysmonPath -ArgumentList "-accepteula -i $configPath" -NoNewWindow -Wait -PassThru
        } -ErrorAction Stop

        [System.Windows.Forms.MessageBox]::Show("Sysmon installed and running with the configuration on the remote computer.", "Installation Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $textboxResults.AppendText("Sysmon installed and running with the configuration on the remote computer.`r`n")
        $sysmonstart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Write-Host "Sysmon started with the configuration on $computerName at $sysmonstart" -ForegroundColor Cyan
        Log_Message -logfile $logfile -Message "Sysmon started with the configuration on $computerName"
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error installing and running Sysmon on the remote computer: $_", "Installation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $textboxResults.AppendText("Error installing and running Sysmon on the remote computer: $_`r`n")
    }
}