@{
    # ---------------------------------------------------------------------
    # Module metadata
    # ---------------------------------------------------------------------
    RootModule             = 'Toolbox.PermissionMatrixHC.psm1'
    ModuleVersion          = '1.0.0'
    GUID                   = 'bbdb2c2a-f17f-4ef6-9b86-d2d6b7b762d3'
    CompatiblePSEditions   = @('Desktop', 'Core')

    Author                 = 'DarkLite1'
    CompanyName            = ''
    Copyright              = '(c) No warranty.'

    # ---------------------------------------------------------------------
    # General information
    # ---------------------------------------------------------------------
    Description            = @'
Toolbox.PermissionMatrixHC provides all private and public functions needed
to process, validate, export, and analyze Permission Matrix Excel files.
It contains functions for Settings processing, Permissions processing,
HTML generation, AD object mapping, logging, export building, email
composition, and more.
'@

    # ---------------------------------------------------------------------
    # PowerShell Version Requirements
    # ---------------------------------------------------------------------
    PowerShellVersion      = '7.0'
    PowerShellHostName     = ''
    PowerShellHostVersion  = ''
    DotNetFrameworkVersion = ''
    CLRVersion             = ''
    ProcessorArchitecture  = ''

    # ---------------------------------------------------------------------
    # Required modules (if any of your code depends on external modules)
    # ---------------------------------------------------------------------
    RequiredModules        = @(
        @{ ModuleName = 'ImportExcel'; ModuleVersion = '7.8.5' }
    )

    # ---------------------------------------------------------------------
    # NestedModules NOT used (psm1 imports individual files internally)
    # ---------------------------------------------------------------------
    NestedModules          = @()

    # ---------------------------------------------------------------------
    # Public / private functions
    #
    # All Public/*.ps1 functions will be exported.
    # All Private/*.ps1 functions will NOT be exported.
    # The psm1 will dot-source these correctly.
    # ---------------------------------------------------------------------
    FunctionsToExport      = @(
        # Add ONLY your public functions here
        'Invoke-PermissionMatrix',
        'Export-PermissionMatrix'
    )

    CmdletsToExport        = @()
    VariablesToExport      = @()
    AliasesToExport        = @()

    # ---------------------------------------------------------------------
    # Private data
    #
    # This is where you store metadata, formatting, etc.
    # ---------------------------------------------------------------------
    PrivateData            = @{
        PSData = @{
            Tags         = @('PermissionMatrix', 'HTML', 'Excel', 'Export', 'AD', 'Logging')
            LicenseUri   = 'https://your-company/license'
            ProjectUri   = 'https://your-repo/project'
            ReleaseNotes = 'Initial release of the refactored PermissionMatrix module.'
        }
    }
}