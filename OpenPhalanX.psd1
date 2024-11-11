@{
    # Module manifest for module 'OpenPhalanX'
    RootModule        = 'OpenPhalanX.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = '12345678-1234-1234-1234-123456789012'
    Author            = 'Your Name'
    CompanyName       = 'Your Company'
    Copyright         = '(c) 2023 Your Company. All rights reserved.'
    Description       = 'A module for remote system management and security.'
    PowerShellVersion = '5.1'
    RequiredModules   = @('ImportExcel', 'ActiveDirectory', 'PSSQLite', 'Microsoft.PowerShell.Utility')
    RequiredAssemblies = @('System.Windows.Forms', 'System.Web', 'System.Net.Http', 'System.Drawing')
    FunctionsToExport = @('Log_Message')
    AliasesToExport   = @()
    CmdletsToExport   = @()
    VariablesToExport = @()
    PrivateData       = @{
        PSData = @{
            Tags = @('Remote Management', 'Security', 'Sysmon', 'Threat Hunting')
            LicenseUri = 'https://opensource.org/licenses/MIT'
            ProjectUri = 'https://github.com/mwhatter/OpenPhalanX'
            IconUri    = 'https://example.com/icon.png'
        }
    }
}
