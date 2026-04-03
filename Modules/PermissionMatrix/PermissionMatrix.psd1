@{
    RootModule           = 'PermissionMatrix.psm1'
    ModuleVersion        = '1.0.0'
    GUID                 = 'bbdb2c2a-f17f-4ef6-9b86-d2d6b7b762d3'
    CompatiblePSEditions = @('Desktop', 'Core')

    Author               = 'DarkLite1'
    CompanyName          = ''
    Copyright            = '(c) No warranty.'

    Description          = @'
PermissionMatrix provides all private and public functions needed
to process, validate, export, and analyze Permission Matrix Excel files.
'@

    PowerShellVersion    = '7.0'

    RequiredModules      = @(
        @{
            ModuleName    = 'ImportExcel'
            ModuleVersion = '7.8.5'
        }
    )

    FunctionsToExport    = @(
        'Invoke-PermissionMatrix'
    )

    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @()

    PrivateData          = @{
        PSData = @{
            Tags         = @('PermissionMatrix', 'Excel', 'Export', 'AD', 'Logging')
            LicenseUri   = 'https://your-company/license'
            ProjectUri   = 'https://your-repo/project'
            ReleaseNotes = 'Initial release.'
        }
    }
}
