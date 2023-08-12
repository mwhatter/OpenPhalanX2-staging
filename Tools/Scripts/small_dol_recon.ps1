$computerName = if ($comboBoxComputerName.SelectedItem) {
    $comboBoxComputerName.SelectedItem.ToString()
} else {$comboBoxComputerName.Text}
if (![string]::IsNullOrEmpty($computerName)) {
    try {
        # Gather system info
        $systemInfo = Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $computerName | Select-Object -Property Name,PrimaryOwnerName,Domain,Model,Manufacturer 
        $textboxResults.AppendText("System Information: `r`n$( $systemInfo | Out-String)")

        # Gather specified AD host info
        $adEntry = Get-ADComputer -Filter "Name -eq '$computerName'" -Properties Name, DNSHostName, IPv4Address, Created, Description, DistinguishedName, Enabled, OperatingSystem, OperatingSystemVersion, SID
        $textboxResults.AppendText("Host AD Information: `r`n$( $adEntry | Out-String)")

        # Enumerate all users known to the system
        $allUsers = Get-CimAssociatedInstance -CimInstance (Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $computerName) -ResultClassName Win32_UserAccount

        # Clear existing users from combobox
        $comboboxUsername.Items.Clear()

        # Gather specified user AD info for each user
        $adUserEntries = @()
        foreach ($user in $allUsers) {
            try {
                $adUserEntry = Get-ADUser -Identity $user.Name -Properties SamAccountName, whenCreated, Description, Enabled, GivenName, Surname, DisplayName, EmailAddress
                $textboxResults.AppendText("User AD Information for $($user.Name) `r`n$( $adUserEntry | Out-String)")

                # Add user to combobox
                $comboboxUsername.Items.Add($adUserEntry.SamAccountName)

                # Append user info to array for later export
                $adUserEntries += $adUserEntry
            } catch {
                $textboxResults.AppendText("")
            }
        }

        New-Item -ItemType Directory -Path ".\Logs\Reports\$computerName\ADRecon" -Force | Out-Null

        # Export to Excel
        $systemInfo | Export-Excel -Path ".\Logs\Reports\$computerName\ADRecon\SysUserADInfo.xlsx"-WorksheetName 'System Info' -AutoSize -AutoFilter -TableStyle Medium6
        $adEntry | Export-Excel -Path ".\Logs\Reports\$computerName\ADRecon\SysUserADInfo.xlsx" -WorksheetName 'Host AD Info' -AutoSize -AutoFilter -TableStyle Medium6 -Append
        $adUserEntries | Export-Excel -Path ".\Logs\Reports\$computerName\ADRecon\SysUserADInfo.xlsx" -WorksheetName 'User AD Info' -AutoSize -AutoFilter -TableStyle Medium6 -Append

        # Print completion message
        $reconstart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $textboxResults.AppendText("System and User AD Information exported to Excel workbook at $reconstart`r`n")
        Write-Host "System and User AD Information exported to Excel workbook at $reconstart" -ForegroundColor Green 
    } catch {
        $textboxResults.AppendText("Error obtaining system information: $_`r`n")
    }
}    