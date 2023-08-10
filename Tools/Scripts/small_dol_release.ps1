$computerName = if ($comboBoxComputerName.SelectedItem) {
    $comboBoxComputerName.SelectedItem.ToString()
} else {
    $comboBoxComputerName.Text
}

if (![string]::IsNullOrEmpty($computerName)) {
    Invoke-Command -ComputerName $computerName -ScriptBlock {
        # Remove the ISOLATION: Allowed Hosts firewall rule
        $isolationRule = Get-NetFirewallRule -DisplayName "ISOLATION: Allowed Hosts" -ErrorAction SilentlyContinue
        if ($isolationRule) {
            Remove-NetFirewallRule -DisplayName "ISOLATION: Allowed Hosts"
        }

        # Set the default outbound action to allow for all profiles
        $firewallProfiles = Get-NetFirewallProfile
        $firewallProfiles | Set-NetFirewallProfile -DefaultOutboundAction Allow

        Write-Host "$computerName isolation changes have been undone at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
        Log_Message -logfile $logfile -Message "$computerName isolated"
    }
} else {
    [System.Windows.Forms.MessageBox]::Show("Please enter a valid computer name.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}