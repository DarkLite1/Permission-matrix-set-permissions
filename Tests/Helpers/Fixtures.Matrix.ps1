function _FakeMatrixSettingsRows {
    param([string]$Scenario)

    switch ($Scenario) {

        'MissingColumn' {
            return @(
                # Missing Permission column
                [pscustomobject]@{
                    MailTo       = 'x@x.com'
                    ADObjectName = 'GroupA'
                    # Permission missing
                }
            )
        }

        'InvalidPermission' {
            return @(
                [pscustomobject]@{
                    MailTo       = 'x@x.com'
                    ADObjectName = 'GroupA'
                    Permission   = 'Z'   # invalid permission character
                }
            )
        }

        'MissingMailTo' {
            return @(
                [pscustomobject]@{
                    MailTo       = $null
                    ADObjectName = 'GroupA'
                    Permission   = 'R'
                }
            )
        }

        default {
            throw "Unknown fake Settings scenario '$Scenario'"
        }
    }
}
function _FakeMatrixPermissionsRows {
    param([string]$Scenario)

    switch ($Scenario) {

        'MissingADObjectName' {
            return @(
                [pscustomobject]@{
                    ADObjectName = $null
                    Path         = 'Test:\Folder'
                    Permission   = 'R'
                }
            )
        }

        'InvalidPermissionChar' {
            return @(
                [pscustomobject]@{
                    ADObjectName = 'GroupA'
                    Path         = 'Test:\Folder'
                    Permission   = 'XYZ' # invalid pattern for your permission logic
                }
            )
        }

        default {
            throw "Unknown fake Permissions scenario '$Scenario'"
        }
    }
}
function Get-MatrixSettingsFixtures {

    return @(
        @{
            Issue           = 'Missing mandatory ApplyDefaultPermissions column'
            SheetMutation   = "New-MatrixExcelFixture -Path 'TestDrive:\Matrix\MutatedSettings.xlsx' -SettingsRows (New-MatrixSettingsFixtureRows -Scenario 'MissingApplyDefaultPermissions')"
            ExpectedMessage = "The column 'ApplyDefaultPermissions' cannot be empty"
        }

        @{
            Issue           = 'Missing mandatory Settings column'
            SheetMutation   = "New-MatrixExcelFixture -Path 'TestDrive:\Matrix\MutatedSettings.xlsx' -SettingsRows (New-MatrixSettingsFixtureRows -Scenario 'MissingAction')"
            ExpectedMessage = "The column 'Action' cannot be empty"
        }

        @{
            Issue           = 'Invalid Action value'
            SheetMutation   = "New-MatrixExcelFixture -Path 'TestDrive:\Matrix\MutatedSettings.xlsx' -SettingsRows (New-MatrixSettingsFixtureRows -Scenario 'InvalidAction')"
            ExpectedMessage = 'Invalid Action'
        }

        @{
            Issue           = 'Missing ComputerName'
            SheetMutation   = "New-MatrixExcelFixture -Path 'TestDrive:\Matrix\MutatedSettings.xlsx' -SettingsRows (New-MatrixSettingsFixtureRows -Scenario 'MissingComputerName')"
            ExpectedMessage = "The column 'ComputerName' cannot be empty"
        }

        @{
            Issue           = 'Missing Path'
            SheetMutation   = "New-MatrixExcelFixture -Path 'TestDrive:\Matrix\MutatedSettings.xlsx' -SettingsRows (New-MatrixSettingsFixtureRows -Scenario 'MissingPath')"
            ExpectedMessage = 'Path'
        }

        @{
            Issue           = 'Missing GroupName (Required by Permissions Sheet)'
            SheetMutation   = "New-MatrixExcelFixture -Path 'TestDrive:\Matrix\MutatedSettings.xlsx' -SettingsRows (New-MatrixSettingsFixtureRows -Scenario 'MissingGroupName') -PermissionsRows (New-MatrixPermissionsFixtureRows -Scenario 'WithGroupNamePlaceholder')"
            ExpectedMessage = "The column 'GroupName' cannot be empty because it is used as a placeholder in the Permissions sheet."
        }

        @{
            Issue           = 'Missing SiteCode (Required by Permissions Sheet)'
            SheetMutation   = "New-MatrixExcelFixture -Path 'TestDrive:\Matrix\MutatedSettings.xlsx' -SettingsRows (New-MatrixSettingsFixtureRows -Scenario 'MissingSiteCode') -PermissionsRows (New-MatrixPermissionsFixtureRows -Scenario 'WithSiteCodePlaceholder')"
            ExpectedMessage = "The column 'SiteCode' cannot be empty because it is used as a placeholder in the Permissions sheet."
        }
    )
}
function Get-MatrixPermissionsFixtures {

    return @(

        # ---------------------------------------------------------------
        # 1. Missing AD group name (column header)
        # ---------------------------------------------------------------
        @{
            Issue    = 'MissingADObjectName'
            Mutation = @'
New-MatrixExcelFixture -Path 'TestDrive:\Matrix\MutatedPermissions.xlsx' -SettingsRows (New-MatrixSettingsFixtureRows -Scenario 'Valid') -PermissionsRows (New-MatrixPermissionsFixtureRows -Scenario 'MissingADObjectName')
'@
            Expected = 'Missing AD object name'
        }


        # ---------------------------------------------------------------
        # 2. Invalid permission characters
        # ---------------------------------------------------------------
        @{
            Issue    = 'InvalidPermissionChar'
            Mutation = @'
New-MatrixExcelFixture -Path 'TestDrive:\Matrix\MutatedPermissions.xlsx' -SettingsRows (New-MatrixSettingsFixtureRows -Scenario 'Valid') -PermissionsRows (New-MatrixPermissionsFixtureRows -Scenario 'InvalidPermissionChar')
'@
            Expected = 'Invalid permission character'
        }

        # ---------------------------------------------------------------
        # 3. Missing rows (Less than 4 rows in sheet)
        # ---------------------------------------------------------------
        @{
            Issue    = 'MissingRows'
            Mutation = @'
New-MatrixExcelFixture -Path 'TestDrive:\Matrix\MutatedPermissions.xlsx' -SettingsRows (New-MatrixSettingsFixtureRows -Scenario 'Valid') -PermissionsRows (New-MatrixPermissionsFixtureRows -Scenario 'MissingRows')
'@
            Expected = 'Missing rows'
        }

        # ---------------------------------------------------------------
        # 4. Missing columns (Less than 2 columns in sheet)
        # ---------------------------------------------------------------
        @{
            Issue    = 'MissingColumns'
            Mutation = @'
New-MatrixExcelFixture -Path 'TestDrive:\Matrix\MutatedPermissions.xlsx' -SettingsRows (New-MatrixSettingsFixtureRows -Scenario 'Valid') -PermissionsRows (New-MatrixPermissionsFixtureRows -Scenario 'MissingColumns')
'@
            Expected = 'Missing columns'
        }

        # ---------------------------------------------------------------
        # 5. Folder name missing (Blank Path)
        # ---------------------------------------------------------------
        @{
            Issue    = 'MissingFolderName'
            Mutation = @'
New-MatrixExcelFixture -Path 'TestDrive:\Matrix\MutatedPermissions.xlsx' -SettingsRows (New-MatrixSettingsFixtureRows -Scenario 'Valid') -PermissionsRows (New-MatrixPermissionsFixtureRows -Scenario 'MissingFolderName')
'@
            Expected = 'Missing folder name'
        }

        # ---------------------------------------------------------------
        # 6. Duplicate folder name
        # ---------------------------------------------------------------
        @{
            Issue    = 'DuplicateFolderName'
            Mutation = @'
New-MatrixExcelFixture -Path 'TestDrive:\Matrix\MutatedPermissions.xlsx' -SettingsRows (New-MatrixSettingsFixtureRows -Scenario 'Valid') -PermissionsRows (New-MatrixPermissionsFixtureRows -Scenario 'DuplicateFolderName')
'@
            Expected = 'Duplicate folder name'
        }

        # ---------------------------------------------------------------
        # 7. Matrix design flaw (Warning - Inaccessible deepest folder)
        # ---------------------------------------------------------------
        @{
            Issue    = 'InaccessibleFolders'
            Mutation = @'
New-MatrixExcelFixture -Path 'TestDrive:\Matrix\MutatedPermissions.xlsx' -SettingsRows (New-MatrixSettingsFixtureRows -Scenario 'Valid') -PermissionsRows (New-MatrixPermissionsFixtureRows -Scenario 'InaccessibleFolders')
'@
            Expected = 'Inaccessible folders'
        }

    )
}
function Get-DisabledMatrixFixtures {
    return @(
        @{
            Description    = 'All matrices disabled'
            FixtureBuilder = {
                New-MatrixExcelFixture -Path 'TestDrive:\Matrix\Matrix1.xlsx' -Disabled
                New-MatrixExcelFixture -Path 'TestDrive:\Matrix\Matrix2.xlsx' -Disabled
            }
            Assertions     = @(
                @{ 
                    Pattern   = '*This matrix file does not contain any enabled matrix settings row and is skipped*' 
                    FileMatch = 'Matrix1' 
                    Not       = $false 
                }
                @{ 
                    Pattern   = '*This matrix file does not contain any enabled matrix settings row and is skipped*' 
                    FileMatch = 'Matrix2' 
                    Not       = $false 
                }
            )
        }
        @{
            Description    = 'One disabled, one enabled'
            FixtureBuilder = {
                New-MatrixExcelFixture -Path 'TestDrive:\Matrix\Matrix1.xlsx' -Disabled
                New-MatrixExcelFixture -Path 'TestDrive:\Matrix\Matrix2.xlsx'
            }
            Assertions     = @(
                @{
                    Pattern   = '*This matrix file does not contain any enabled matrix settings row and is skipped*' 
                    FileMatch = 'Matrix1' 
                    Not       = $false 
                }
                @{ 
                    Pattern   = '*This matrix file does not contain any enabled matrix settings row and is skipped*' 
                    FileMatch = 'Matrix2' 
                    Not       = $true 
                }
            )
        }
    )
}
function Get-DuplicateMatrixFixtures {
    return @(
        @{
            Description    = 'Duplicate ComputerName + Path combination'
            FixtureBuilder = {

                $path = 'TestDrive:\Matrix\DUP.xlsx'

                $settings = @(
                    [pscustomobject]@{
                        Status       = 'Enabled'
                        SiteName     = 'HQ South'
                        SiteCode     = 'CS&L'
                        ComputerName = 'BEL$FFRAN0001'
                        Path         = 'E:\DEPARTMENTS\Sagrev\GROUPS\C&S&L'
                        GroupName    = 'BEL ROL-AGS-SAGREV'
                        Action       = 'Fix'
                    }
                    [pscustomobject]@{
                        Status       = 'Enabled'
                        SiteName     = 'HQ South'
                        SiteCode     = 'CS&L'
                        ComputerName = 'BEL$FFRAN0001'   # DUPLICATE
                        Path         = 'E:\DEPARTMENTS\Sagrev\GROUPS\C&S&L' # DUPLICATE
                        GroupName    = 'BEL ROL-AGS-SAGREV'
                        Action       = 'Check'
                    }
                )

                New-MatrixExcelFixture -Path $path -SettingsRows $settings
            }
            ExpectedError  = 'Duplicate ComputerName/Path'
        }
    )
}
function Get-DefaultPermissionsMergeFixtures {
    return @(
        @{
            Description             = 'ApplyDefaultPermissions=$false: Only Matrix is returned'
            ApplyDefaultPermissions = $false
            DefaultsRows            = @( [pscustomobject]@{ ADObject = 'IT_Staff' ; Permission = 'R' } )
            MatrixRows              = @( [pscustomobject]@{ ADObject = 'HR_Team'  ; Permission = 'M' } )
            ExpectedMerged          = @( [pscustomobject]@{ ADObject = 'HR_Team'  ; Permission = 'M' } )
            ExpectedError           = $null
        }

        @{
            Description             = 'ApplyDefaultPermissions=$true (No Conflict): Defaults are appended'
            ApplyDefaultPermissions = $true
            DefaultsRows            = @( [pscustomobject]@{ ADObject = 'IT_Staff' ; Permission = 'R' } )
            MatrixRows              = @( [pscustomobject]@{ ADObject = 'HR_Team'  ; Permission = 'M' } )
            ExpectedMerged          = @( 
                [pscustomobject]@{ ADObject = 'HR_Team'  ; Permission = 'M' },
                [pscustomobject]@{ ADObject = 'IT_Staff' ; Permission = 'R' } 
            )
            ExpectedError           = $null
        }

        @{
            Description             = 'ApplyDefaultPermissions=$true (Conflict): Throws terminating error'
            ApplyDefaultPermissions = $true
            DefaultsRows            = @( [pscustomobject]@{ ADObject = 'IT_Staff' ; Permission = 'R' } )
            MatrixRows              = @( [pscustomobject]@{ ADObject = 'IT_Staff' ; Permission = 'F' } )
            ExpectedMerged          = $null
            ExpectedError           = 'Defaults conflict detected.*IT_Staff'
        }
    )
}
function Get-AdObjectBuildFixtures {
    return @(
        @{
            Description    = 'Two AD objects'
            FixtureBuilder = {
                return @{
                    'GroupA' = @{
                        adObject      = @{ Name = 'GroupA'; ObjectClass = 'group' }
                        adGroupMember = @()
                    }
                    'UserB'  = @{
                        adObject      = @{ Name = 'UserB' ; ObjectClass = 'user' }
                        adGroupMember = @()
                    }
                }
            }
            Expected       = 2
        }
    )
}
function Get-MatrixBuildFixtures {
    return @(
        @{
            Description    = 'Two matrix files → two HTML tables'

            FixtureBuilder = {
                New-MatrixExcelFixture -Path 'TestDrive:\Matrix\M1.xlsx'
                New-MatrixExcelFixture -Path 'TestDrive:\Matrix\M2.xlsx'
            }

            ExpectedFiles  = 2
        }

        @{
            Description    = 'One matrix file → one HTML table'

            FixtureBuilder = {
                New-MatrixExcelFixture -Path 'TestDrive:\Matrix\M1.xlsx'
            }

            ExpectedFiles  = 1
        }
    )
}