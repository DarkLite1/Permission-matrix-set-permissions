@{
    RootModule           = 'PermissionMatrix.psm1'
    ModuleVersion        = '1.0.0'
    GUID                 = 'bbdb2c2a-f17f-4ef6-9b86-d2d6b7b762d3'
    CompatiblePSEditions = @('Desktop', 'Core')

    Author               = 'DarkLite1'
    CompanyName          = ''
    Copyright            = '(c) No warranty. Provided as-is.'

    Description          = @'
A robust PowerShell tool designed to apply, verify, and manage NTFS and SMB permissions at scale. 
By reading a centralized Excel-based matrix and a JSON configuration file, the script automates 
the complex task of ensuring folder security across multiple remote computers.
'@

    PowerShellVersion    = '7.0'

    RequiredModules      = @(
        @{
            ModuleName    = 'ImportExcel'
            ModuleVersion = '7.8.5'
        }
    )

    FunctionsToExport    = @(
        'Invoke-PermissionMatrix',
        'Invoke-PermissionMatrixAuditReport'
    )

    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @()

    PrivateData          = @{
        PSData = @{
            Tags         = @('PermissionMatrix', 'Excel', 'NTFS', 'SMB', 'Security', 'ServiceNow', 'AD')
            ProjectUri   = 'https://github.com/DarkLite1/Permission-matrix-set-permissions'
            ReleaseNotes = 'Initial release.'
        }
    }
}