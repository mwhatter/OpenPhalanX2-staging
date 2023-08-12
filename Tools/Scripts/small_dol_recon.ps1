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

        # Fetch AD groups the computer is a member of
        $computerGroups = Get-ADPrincipalGroupMembership -Identity $adEntry.DistinguishedName

        # Format the group names as a bulleted list
        $computerGroupNames = "- " + ($computerGroups.Name -join "`r`n- ")

        # Count the number of groups
        $groupCount = $computerGroups.Count

        # Append the formatted output to the results with separators and a count for computer groups
        $textboxResults.AppendText("Host $computerName is a member of $groupCount groups:`r`n$computerGroupNames`r`n")
        $textboxResults.AppendText("======================================`r`n`r`n`r`n`r`n")

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

                # Fetch AD groups the user is a member of
                $userGroups = Get-ADPrincipalGroupMembership -Identity $user.Name
                $userGroupNames = "- " + ($userGroups.Name -join "`r`n- ")

                # Count the number of groups the user is a member of
                $userGroupCount = $userGroups.Count

                # Append the formatted output to the results for user groups
                $textboxResults.AppendText("User $($user.Name) is a member of $userGroupCount groups:`r`n$userGroupNames`r`n")
                $textboxResults.AppendText("`r`n======================================`r`n`r`n`r`n`r`n`r`n")

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
