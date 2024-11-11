# Import necessary assemblies and modules
Import-Module ImportExcel
Import-Module ActiveDirectory
Import-Module PSSQLite
Import-Module Microsoft.PowerShell.Utility

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Web
Add-Type -AssemblyName System.Net.Http
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# Define all functions from Defending_Off_the_Land.ps1

function Log_Message {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $true)]
        [string]$LogFile
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogFile -Value "Time=[$timestamp] User=[$Username] Message=[$Message]"
}

# Export all functions using Export-ModuleMember
Export-ModuleMember -Function Log_Message
