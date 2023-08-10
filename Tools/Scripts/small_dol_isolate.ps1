$computerName = if ($comboBoxComputerName.SelectedItem) {
    $comboBoxComputerName.SelectedItem.ToString()
} else {
    $comboBoxComputerName.Text
}

if (![string]::IsNullOrEmpty($computerName)) {
    Invoke-Command -ComputerName $computerName -ScriptBlock {
        # Get the local host's IP addresses
        $localHostIPs = (Get-NetIPAddress -AddressFamily IPv4 -CimSession localhost).IPAddress | Where-Object { $_ -notlike "127.0.0.*" }
        
        # Check if Windows firewall is enabled
        $firewallProfiles = Get-NetFirewallProfile
        if ($firewallProfiles.Enabled -contains $false) {
            # Enable Windows Firewall for all profiles
            $firewallProfiles | Set-NetFirewallProfile -Enabled:True
        }
        
        # Create the isolation firewall rule if it doesn't exist
        $isolationRule = Get-NetFirewallRule -DisplayName "ISOLATION: Allowed Hosts" -ErrorAction SilentlyContinue
        if (!$isolationRule) {
            New-NetFirewallRule -DisplayName "ISOLATION: Allowed Hosts" -Direction Outbound -RemoteAddress $localHostIPs -Action Allow -Enabled:True
        } else {
            Set-NetFirewallRule -DisplayName "ISOLATION: Allowed Hosts" -RemoteAddress $localHostIPs
        }
        
        # Set the default outbound action to block for all profiles
        $firewallProfiles | Set-NetFirewallProfile -DefaultOutboundAction Block

        Write-Host "$computerName isolated at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
        Log_Message -logfile $logfile -Message "$computerName isolated"
    }
} else {
    [System.Windows.Forms.MessageBox]::Show("Please enter a valid computer name.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}