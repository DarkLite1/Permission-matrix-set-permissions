#Requires -Version 7
#Requires -Modules Pester, ImportExcel

BeforeAll {
    $testParams = @{
        ScriptName    = 'Test (Brecht)'
        MaxConcurrent = @{
            Computers             = 1
            JobsPerRemoteComputer = 1
            FoldersPerMatrix      = 2
        }
        Matrix        = @{
            FolderPath             = New-Item 'TestDrive:/Matrix' -ItemType Directory
            DefaultsFile           = New-Item 'TestDrive:/Default.xlsx' -ItemType File
            Archive                = $false
            ExcludedSamAccountName = @()
        }
        LogFolder     = 'TestDrive:\log\File and folder\Test (Brecht)'
        ScriptPath    = @{
            TestRequirementsFile = New-Item 'TestDrive:/TestRequirements.ps1' -ItemType File
            SetPermissionFile    = New-Item 'TestDrive:/SetPermissions.ps1' -ItemType File
        }
        ScriptAdmin   = 'admin@contoso.com'
        
    }
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')

    $testCherwellFolder = New-Item 'TestDrive:/Cherwell' -ItemType Directory

    #region Valid Excel files
    $testMatrix = @(
        [PSCustomObject]@{
            Path   = 'Path'
            ACL    = @{'Bob' = 'L' }
            Parent = $true
            Ignore = $false
        }
    )
    $testPermissions = @(
        [PSCustomObject]@{P1 = $null      ; P2 = 'bob' }
        [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
        [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
        [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
        [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
    )
    $testSettings = @(
        [PSCustomObject]@{
            Status       = 'Enabled'
            ComputerName = 'S1'
            Path         = 'E:\Department'
            Action       = 'Check'
        }
    )
    $testDefaultSettings = @(
        [PSCustomObject]@{
            ADObjectName = 'Bob'
            Permission   = 'L'
            MailTo       = 'Bob@contoso.com'
        }
        [PSCustomObject]@{
            ADObjectName = 'Mike'
            Permission   = 'R'
        }
    )
    #endregion

    $testDefaultSettings |
    Export-Excel -Path $testParams.Matrix.DefaultsFile -WorksheetName 'Settings'

    $SettingsParams = @{
        Path          = Join-Path $testParams.Matrix.FolderPath 'Matrix.xlsx'
        WorkSheetName = 'Settings'
    }
    $PermissionsParams = @{
        Path          = $SettingsParams.Path
        WorkSheetName = 'Permissions'
        NoHeader      = $true
    }

    function Compare-HashTableHC {
        param (
            [Parameter(Mandatory)]
            [hashtable]$ReferenceObject,
            [Parameter(Mandatory)]
            [hashtable]$DifferenceObject
        )

        (
            $ReferenceObject.GetEnumerator() |
            Sort-Object { $_.Key } | ConvertTo-Json
        ) |
        Should -BeExactly (
            $DifferenceObject.GetEnumerator() |
            Sort-Object { $_.Key } | ConvertTo-Json
        )
    }

    Mock Invoke-Command
    Mock New-PSSession
    Mock Send-MailHC
    Mock Test-MatrixPermissionsHC
    Mock Test-MatrixSettingHC
    Mock Wait-MaxRunningJobsHC
    Mock Write-EventLog
    Mock Write-Warning
    Mock Test-FormDataHC
    Mock Get-AdUserPrincipalNameHC
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @(
        'ConfigurationJsonFile'
    ) {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'stop the script and send an e-mail to the admin when' {
    BeforeAll {
        $MailAdminParams = {
            ($To -eq $testParams.ScriptAdmin) -and
            ($Priority -eq 'High') -and
            ($Subject -eq 'FAILURE')
        }
    }
    Context 'a file or folder is not found' {
        It 'Script SetPermissionFile' {
            $testParams = $testParams.Clone()
            $testParams.ScriptPath.SetPermissionFile = 'NonExisting.ps1'

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and
                ($Message -like '*ScriptPath.NonExisting.ps1*not found*')
            }
        }
        It 'Script TestRequirements' {
            $testParams = $testParams.Clone()
            $testParams.ScriptPath.TestRequirementsFile = 'ShareConfigNotExisting.ps1'

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and
                ($Message -like '*ScriptPath.ShareConfigNotExisting.ps1*not found*')
            }
        }
        It 'LogFolder' {
            $testParams = $testParams.Clone()
            $testParams.LogFolder = 'x:\NonExistingLog'

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and
                ($Message -like "*Failed to create log folder 'x:\NonExistingLog'*")
            }
        }
        It 'CherwellFolder' {
            $testParams = $testParams.Clone()
            $testParams.CherwellFolder = 'NonExistingFolder'

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and
                ($Message -like '*NonExistingFolder*not found*')
            }
        }
    }
    Context 'the default settings file' {
        It 'is not found' {
            $clonedParams = $testParams.Clone()
            $clonedParams.DefaultsFile = 'notExisting'

            .$testScript @clonedParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and
                ($Message -like "*$($clonedParams.DefaultsFile)*Cannot find*")
            }
        }
        It "does not have the worksheet 'Settings'" {
            $clonedParams = $testParams.Clone()
            $clonedParams.DefaultsFile = New-Item 'TestDrive:/Folder/Default.xlsx' -ItemType File -Force

            '1' | Export-Excel -Path $clonedParams.DefaultsFile -WorksheetName Sheet1

            .$testScript @clonedParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and
                ($Message -like "*'$($clonedParams.DefaultsFile)'* worksheet 'Settings' not found*")
            }
        }

        $TestCases = @(
            @{
                Name         = "column header 'MailTo'"
                DefaultsFile = @(
                    [PSCustomObject]@{
                        ADObjectName = 'Bob'
                        Permission   = 'L'
                    }
                    [PSCustomObject]@{
                        ADObjectName = 'Mike'
                        Permission   = 'R'
                    }
                )
                errorMessage = "Column header 'MailTo' not found"
            }
            @{
                Name         = "column header 'ADObjectName'"
                DefaultsFile = @(
                    [PSCustomObject]@{
                        Permission = 'L'
                        MailTo     = 'Bob@mail.com'
                    }
                    [PSCustomObject]@{
                        Permission = 'R'
                    }
                )
                errorMessage = "Column header 'ADObjectName' not found"
            }
            @{
                Name         = "column header 'Permission'"
                DefaultsFile = @(
                    [PSCustomObject]@{
                        ADObjectName = 'Bob'
                        MailTo       = 'Bob@mail.com'
                    }
                    [PSCustomObject]@{
                        ADObjectName = 'Mike'
                    }
                )
                errorMessage = "Column header 'Permission' not found"
            }
            @{
                Name         = "'MailTo' addresses"
                DefaultsFile = @(
                    [PSCustomObject]@{
                        ADObjectName = 'Bob'
                        Permission   = 'L'
                        MailTo       = $null
                    }
                    [PSCustomObject]@{
                        ADObjectName = 'Mike'
                        Permission   = 'R'
                        MailTo       = ' '
                    }
                )
                errorMessage = 'No mail addresses found'
            }
        )

        It 'is missing <Name>' -ForEach $TestCases {
            $clonedParams = $testParams.Clone()
            $clonedParams.DefaultsFile = New-Item 'TestDrive:/Folder/Default.xlsx' -ItemType File -Force

            $DefaultsFile | Export-Excel -Path $clonedParams.DefaultsFile -WorksheetName Settings

            .$testScript @clonedParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and
                ($Message -like "*$($clonedParams.DefaultsFile)*$errorMessage*")
            }
        }
    }
    Context 'the argument CherwellFolder is used but' {
        BeforeAll {
            $testCherwellParams = @{
                CherwellFolder                = $testCherwellFolder.FullName
                CherwellAdObjectsFileName     = 'BNL Matrix AD object names.csv'
                CherwellFormDataFileName      = 'BNL Matrix form data.csv'
                CherwellExcelOverviewFileName = 'Overview.xlsx'
            }
        }
        It '<Name> is missing' -ForEach @(
            @{ Name = 'CherwellAdObjectsFileName' }
            @{ Name = 'CherwellFormDataFileName' }
            @{ Name = 'CherwellGroupManagersFileName' }
            @{ Name = 'CherwellAccessListFileName' }
            @{ Name = 'CherwellExcelOverviewFileName' }
        ) {
            $clonedCherwellParams = $testCherwellParams.Clone()
            $clonedCherwellParams.$Name = ''

            .$testScript @testParams @clonedCherwellParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and
                ($Message -like "*Parameter '$Name' is mandatory when the parameter CherwellFolder is used*")
            }
        }
    }
}
Describe 'a sub folder in the log folder' {
    BeforeAll {
        @(
            [PSCustomObject]@{
                Status       = $null
                ComputerName = 'S1'
                Path         = 'E:\Test'
                Action       = 'Check'
            }
        ) | Export-Excel @SettingsParams
        @(
            [PSCustomObject]@{P1 = $null      ; P2 = 'Manager' }
            [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
            [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
            [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
            [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
        ) | Export-Excel @PermissionsParams

        .$testScript @testParams
    }
    It "is created for each specific Excel file regardless of its 'Status'" {
        @(Get-ChildItem -Path $testParams.LogFolder -Directory).Count |
        Should -BeExactly 1
    }
    It 'the Excel file is copied to the log folder' {
        $testMatrixLogFolder = Get-ChildItem -Path $testParams.LogFolder -Directory

        @(Get-ChildItem -Path $testMatrixLogFolder.FullName -File -Filter '*.xlsx').Count | Should -BeExactly 1
    }
}
Describe "when 'Matrix.Archive' is true then" {
    BeforeAll {
        @(
            [PSCustomObject]@{
                ComputerName = 'S1'
                Path         = 'E:\Department'
                Action       = 'Check'
            }
        ) | Export-Excel @SettingsParams
        @(
            [PSCustomObject]@{P1 = $null      ; P2 = 'Manager' }
            [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
            [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
            [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
            [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
        ) | Export-Excel @PermissionsParams

        $testParams = $testParams.Clone()
        $testParams.Matrix.Archive = $true
    }
    It "a sub folder in the 'Matrix.FolderPath' named 'Archive' is created" {
        "$($testParams.Matrix.FolderPath)\Archive" | Should -Exist
    }
    It 'all matrix files are moved to the archive folder, even disabled ones' {
        $SettingsParams.Path | Should -Not -Exist
        "$($testParams.Matrix.FolderPath)\Archive\Matrix.xlsx" | Should -Exist
    }
    It 'a matrix with the same name is overwritten in the archive folder' {
        $testFile = "$($testParams.Matrix.FolderPath)\Archive\Matrix.xlsx"
        $testFile | Remove-Item -Force -EA Ignore

        @(
            [PSCustomObject]@{
                ComputerName = 'S1'
                Path         = 'E:\Department'
                Action       = 'Check'
            }
        ) | Export-Excel -Path $testFile -WorksheetName $SettingsParams.WorkSheetName

        $testFile | Should -Exist

        @(
            [PSCustomObject]@{
                ComputerName = 'S2'
                Path         = 'E:\Department'
                Action       = 'Check'
            }
        ) | Export-Excel @SettingsParams
        @(
            [PSCustomObject]@{P1 = $null      ; P2 = 'Manager' }
            [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
            [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
            [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
            [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
        ) | Export-Excel @PermissionsParams

        .$testScript @testParams -Archive

        $testFile | Should -Exist
        $SettingsParams.Path | Should -Not -Exist
        (Import-Excel -Path $testFile -WorksheetName Settings).ComputerName |
        Should -Be 'S2'
    }
    It 'multiple matrix files are moved to the archive folder' {
        Remove-Item -Path "$($testParams.Matrix.FolderPath)\Archive" -Recurse -EA Ignore
        1..5 | ForEach-Object {
            $FileName = "$($testParams.Matrix.FolderPath)\Matrix $_.xlsx"
            @(
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = 'S1'
                    Path         = 'E:\Department'
                    GroupName    = 'G1'
                    SiteName     = 'S1'
                    SiteCode     = 'C1'
                    Action       = 'Check'
                }
            ) | Export-Excel -Path $FileName -WorksheetName Settings
            @(
                [PSCustomObject]@{P1 = $null      ; P2 = 'Manager' }
                [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
                [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
                [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
                [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
            ) | Export-Excel -Path $FileName -WorksheetName Permissions -NoHeader
        }

        .$testScript @testParams -Archive

        (Get-ChildItem "$($testParams.Matrix.FolderPath)\Matrix*" -File).Count | Should -BeExactly 0
        (Get-ChildItem "$($testParams.Matrix.FolderPath)\Archive" -File).Count |
        Should -BeExactly 5
    }
}
Describe 'do not invoke the script to set permissions when' {
    It "there's only a default settings file in the 'Matrix.FolderPath' folder" {
        .$testScript @testParams

        Should -Not -Invoke Invoke-Command
    }
    It "there are only other file types than '.xlsx' in the 'Matrix.FolderPath' folder" {
        1 | Out-File "$($testParams.Matrix.FolderPath)\Wrong.txt"
        1 | Out-File "$($testParams.Matrix.FolderPath)\Wrong.csv"

        .$testScript @testParams

        Should -Not -Invoke Invoke-Command
    }
    It "there are only valid matrixes in subfolders of the 'Matrix.FolderPath' folder" {
        $Folder = (New-Item "$($testParams.Matrix.FolderPath)\Archive" -ItemType Directory -Force -EA Ignore).FullName
        @(
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'S1'
                Path         = 'E:\Department'
                Action       = 'Check'
            }
        ) | Export-Excel -Path "$Folder/Matrix.xlsx" -WorksheetName Settings
        @(
            [PSCustomObject]@{P1 = $null      ; P2 = 'Manager' }
            [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
            [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
            [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
            [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
        ) | Export-Excel -Path "$Folder/Matrix.xlsx" -WorksheetName Permissions -NoHeader

        .$testScript @testParams

        Should -Not -Invoke Invoke-Command
    }
    It "the 'Status' in the worksheet 'Settings' of the matrix file is not set to 'Enabled'" {
        @(
            [PSCustomObject]@{
                Status       = 'NOTEnabled'
                ComputerName = 'S1'
                Path         = 'E:\Department'
                Action       = 'Check'
            }
        ) | Export-Excel @SettingsParams
        @(
            [PSCustomObject]@{P1 = $null      ; P2 = 'Manager' }
            [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
            [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
            [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
            [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
        ) | Export-Excel @PermissionsParams

        .$testScript @testParams

        Should -Not -Invoke Invoke-Command
    }
}
Describe 'a FatalError object is registered' {
    AfterEach {
        $Error.Clear()
        Remove-Item -Path "$($testParams.LogFolder)\*" -Recurse -Force -EA Ignore
        Remove-Item -Path "$($testParams.Matrix.FolderPath)\*" -Exclude $TestDefaultsFileName -Recurse -Force -EA Ignore
    }
    Context "for the Excel 'File' when" {
        It "building the matrix with 'ConvertTo-MatrixAclHC' fails" {
            Mock ConvertTo-MatrixAclHC {
                throw 'Failed building the matrix'
            }

            @(
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = 'S1'
                    Path         = 'E:\Department'
                    Action       = 'Check'
                }
            ) | Export-Excel @SettingsParams
            @(
                [PSCustomObject]@{P1 = $null      ; P2 = 'Manager' }
                [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
                [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
                [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
                [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
            ) | Export-Excel @PermissionsParams

            .$testScript @testParams

            @{
                Type        = 'FatalError'
                Name        = 'Unknown error'
                Description = 'While checking the input and generating the matrix an error was reported.'
                Value       = 'Failed building the matrix'
            }.GetEnumerator().ForEach( {
                    $ImportedMatrix.File.Check.($_.Key) | Should -Be $_.Value
                })
        }
        It 'the worksheet Settings is not found' {
            @(
                [PSCustomObject]@{P1 = $null      ; P2 = 'Manager' }
                [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
                [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
                [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
                [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
            ) | Export-Excel @PermissionsParams

            .$testScript @testParams

            @{
                Type        = 'FatalError'
                Name        = 'Excel file incorrect'
                Description = "The worksheets 'Settings' and 'Permissions' are mandatory."
                Value       = "Failed importing the Excel workbook '$($PermissionsParams.Path)' with worksheet 'Settings'*"
            }.GetEnumerator().ForEach( {
                    $ImportedMatrix.File.Check.($_.Key) |
                    Should -BeLike $_.Value
                })
        } -Skip
        It 'the worksheet Settings is empty' {
            @(
                [PSCustomObject]@{P1 = $null      ; P2 = 'Manager' }
                [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
                [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
                [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
                [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
            ) | Export-Excel @PermissionsParams

            #region Add empty worksheet
            $pkg = New-Object OfficeOpenXml.ExcelPackage (Get-Item -Path $SettingsParams.Path)
            $null = $pkg.Workbook.Worksheets.Add('Settings')
            $pkg.Save()
            $pkg.Dispose()
            #endregion

            .$testScript @testParams

            @{
                Type        = 'FatalError'
                Name        = 'Excel file incorrect'
                Description = "The worksheets 'Settings' and 'Permissions' are mandatory."
                Value       = "Failed importing the Excel workbook '$($SettingsParams.Path)' with worksheet 'Settings'*"
            }.GetEnumerator().ForEach( {
                    $ImportedMatrix.File.Check.($_.Key) |
                    Should -BeLike $_.Value
                })
        } -Skip
        It "the worksheet Permissions is not found when the 'Settings' sheet has 'Status' set to 'Enabled'" {
            $testSettings | Export-Excel @SettingsParams

            .$testScript @testParams

            @{
                Type        = 'FatalError'
                Name        = 'Excel file incorrect'
                Description = "The worksheets 'Settings' and 'Permissions' are mandatory."
                Value       = "Worksheet 'Permissions' not found"
            }.GetEnumerator().ForEach( {
                    $ImportedMatrix.File.Check.($_.Key) | Should -Be $_.Value
                })
        } -Skip
        It "the worksheet Permissions is empty when the 'Settings' sheet has 'Status' set to 'Enabled'" {
            $testSettings | Export-Excel @SettingsParams

            #region Add empty worksheet
            $pkg = New-Object OfficeOpenXml.ExcelPackage (Get-Item -Path $SettingsParams.Path)
            $null = $pkg.Workbook.Worksheets.Add('Permissions')
            $pkg.Save()
            $pkg.Dispose()
            #endregion

            .$testScript @testParams

            @{
                Type        = 'FatalError'
                Name        = 'Excel file incorrect'
                Description = "The worksheets 'Settings' and 'Permissions' are mandatory."
                Value       = "Worksheet 'Permissions' is empty"
            }.GetEnumerator().ForEach( {
                    $ImportedMatrix.File.Check.($_.Key) | Should -Be $_.Value
                })
        }
    }
    Context "for the worksheet 'Permissions' when" {
        AfterAll {
            Mock Test-MatrixPermissionsHC
        }
        It "'Test-MatrixPermissionsHC' detects an input problem" {
            Mock Test-MatrixPermissionsHC {
                @{
                    Type = 'Warning'
                    Name = 'Matrix permission incorrect'
                }
            }

            $testSettings | Export-Excel @SettingsParams
            $testPermissions | Export-Excel @PermissionsParams

            .$testScript @testParams

            $ImportedMatrix.Permissions.Check.Name |
            Should -Contain 'Matrix permission incorrect'
        }
    }
    Context "for the worksheet 'Settings' when" {
        AfterAll {
            Mock Test-MatrixSettingHC
            Mock Test-ExpandedMatrixHC
        }
        It 'a duplicate ComputerName/Path combination is found' {
            $testProblem = @(
                [PSCustomObject]@{
                    Type        = 'FatalError'
                    Name        = 'Duplicate ComputerName/Path combination'
                    Description = "Every 'ComputerName' combined with a 'Path' needs to be unique over all the 'Settings' worksheets found in all the active matrix files."
                    Value       = @{
                        ('S1.' + $env:USERDNSDOMAIN) = 'E:\DUPLICATE'
                    }
                }
            )

            @(
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = $($testProblem.Value.Keys)
                    Path         = 'E:\Reports'
                    Action       = 'Check'
                }
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = $($testProblem.Value.Keys)
                    Path         = $($testProblem.Value.Values)
                    Action       = 'Check'
                }
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = $($testProblem.Value.Keys)
                    Path         = $($testProblem.Value.Values)
                    Action       = 'Fix'
                }
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = 'S3'
                    Path         = 'E:\Department'
                    Action       = 'Check'
                }
            ) | Export-Excel @SettingsParams

            $testPermissions | Export-Excel @PermissionsParams

            .$testScript @testParams

            $toTest = @($ImportedMatrix.Settings.Where(
                    { $_.Import.Path -eq $testProblem.Value.Values }
                ))

            $toTest.Count | Should -BeExactly 2

            foreach ($testMatrix in $toTest) {
                $testCheck = $testMatrix.Check | Where-Object {
                    $_.Name -eq $testProblem.Name
                }
                $testCheck.Type | Should -Be $testProblem.Type
                $testCheck.Name | Should -Be $testProblem.Name
                $testCheck.Description | Should -Be $testProblem.Description
                $testCheck.Value.Name | Should -Be $testProblem.Value.Name
                $testCheck.Value.Value | Should -Be $testProblem.Value.Value
            }

        }
        It "'Test-MatrixSettingHC' detects an input problem" {
            $testProblem = @{
                Name = 'Matrix setting incorrect'
            }
            Mock Test-MatrixSettingHC {
                $testProblem
            }

            $testSettings | Export-Excel @SettingsParams
            $testPermissions | Export-Excel @PermissionsParams

            .$testScript @testParams

            $testProblem.Name |
            Should -Be ($ImportedMatrix.Settings.Check | Where-Object Name -EQ $testProblem.Name).Name
        }
        It "'Test-ExpandedMatrixHC' detects a problem" {
            Mock Test-MatrixSettingHC
            $testProblem = @{
                Name = 'Expansion incorrect'
            }
            Mock Test-ExpandedMatrixHC {
                $testProblem
            }
            Mock ConvertTo-MatrixAclHC {
                $testMatrix
            }

            Mock Get-ADObjectDetailHC { $true }

            $testSettings | Export-Excel @SettingsParams
            $testPermissions | Export-Excel @PermissionsParams

            .$testScript @testParams

            $testProblem.Name |
            Should -Be ($ImportedMatrix.Settings.Check | Where-Object Name -EQ $testProblem.Name).Name
        }
    }
}
Describe 'a Warning object is registered' {
    AfterEach {
        $Error.Clear()
        Remove-Item -Path "$($testParams.LogFolder)\*" -Recurse -Force -EA Ignore
        Remove-Item -Path "$($testParams.Matrix.FolderPath)\*" -Exclude $TestDefaultsFileName -Recurse -Force -EA Ignore
    }
    Context "for the Excel 'File' when" {
        It "the worksheet 'Settings' has no row with status 'Enabled'" {
            @(
                [PSCustomObject]@{
                    Status       = $null
                    ComputerName = 'A'
                    Path         = 'E:\Reports'
                    Action       = 'Check'
                }
            ) | Export-Excel @SettingsParams
            $testPermissions | Export-Excel @PermissionsParams

            .$testScript @testParams

            @{
                Type        = 'Warning'
                Name        = 'Matrix disabled'
                Description = 'Every Excel file needs at least one enabled matrix.'
                Value       = "The worksheet 'Settings' does not contain a row with 'Status' set to 'Enabled'."
            }.GetEnumerator().ForEach( {
                    $ImportedMatrix.File.Check.($_.Key) | Should -Be $_.Value
                })
        }
    }
}
Describe "each row in the worksheet 'settings'" {
    BeforeAll {
        Mock ConvertTo-MatrixADNamesHC { @{} }
        Mock ConvertTo-MatrixAclHC
        Mock Test-AdObjectsHC
        Mock Test-MatrixSettingHC

        @(
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'pc1'
                Path         = 'E:\Department'
                Action       = 'Check'
                GroupName    = 'A'
                SiteCode     = 'B'
            }
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'pc2'
                Path         = 'E:\Reports'
                Action       = 'Check'
                GroupName    = 'C'
                SiteCode     = 'D'
            }
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'pc3'
                Path         = 'E:\Finance'
                Action       = 'Check'
                GroupName    = 'E'
                SiteCode     = 'F'
            }
        ) | Export-Excel @SettingsParams

        @(
            [PSCustomObject]@{P1 = $null      ; P2 = 'bob' }
            [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
            [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
            [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
            [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
        ) | Export-Excel @PermissionsParams

        .$testScript @testParams -EA ignore
    }
    It 'is tested for incorrect input' {
        Should -Invoke Test-MatrixSettingHC -Exactly 3 -Scope Describe
        @('pc1', 'pc2', 'pc3') | ForEach-Object {
            Should -Invoke Test-MatrixSettingHC -Exactly 1 -Scope Describe -ParameterFilter {
                $Setting.ComputerName -eq $_
            }
        }
    }
    Context 'creates a unique matrix with' {
        It 'complete SamAccountNames constructed from the header rows' {
            function testColumnHeaders {
                ($null -eq $ColumnHeaders[0].P1) -and
                ($ColumnHeaders[0].P2 -eq 'bob') -and
                ($ColumnHeaders[1].P1 -eq 'SiteCode') -and
                ($ColumnHeaders[1].P2 -eq 'SiteCode') -and
                ($ColumnHeaders[2].P1 -eq 'GroupName') -and
                ($ColumnHeaders[2].P2 -eq 'GroupName')
            }

            Should -Invoke ConvertTo-MatrixADNamesHC -Exactly 3 -Scope Describe
            Should -Invoke ConvertTo-MatrixADNamesHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($Begin -eq 'A') -and ($Middle -eq 'B') -and
                (testColumnHeaders)
            }
            Should -Invoke ConvertTo-MatrixADNamesHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($Begin -eq 'C') -and ($Middle -eq 'D') -and
                (testColumnHeaders)
            }
            Should -Invoke ConvertTo-MatrixADNamesHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($Begin -eq 'E') -and ($Middle -eq 'F') -and
                (testColumnHeaders)
            }
        }
        It 'path and Acl' {
            Should -Invoke ConvertTo-MatrixAclHC -Exactly 3 -Scope Describe
        }
    }
}
Describe "the worksheet 'Permissions' is" {
    BeforeAll {
        Mock Test-MatrixPermissionsHC

        @(
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'PC1'
                Path         = 'E:\Department'
                Action       = 'Check'
                GroupName    = 'A'
                SiteCode     = 'B'
            }
        ) | Export-Excel @SettingsParams

        @(
            [PSCustomObject]@{P1 = $null      ; P2 = 'bob' }
            [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
            [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
            [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
            [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
        ) | Export-Excel @PermissionsParams

        .$testScript @testParams
    }
    It 'tested for incorrect input' {
        Should -Invoke Test-MatrixPermissionsHC -Exactly 1 -Scope Describe
        Should -Invoke Test-MatrixPermissionsHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($null -eq $Permissions[0].P1) -and
            ($Permissions[0].P2 -eq 'bob') -and
            ($Permissions[1].P1 -eq 'SiteCode') -and
            ($Permissions[1].P2 -eq 'SiteCode') -and
            ($Permissions[2].P1 -eq 'GroupName') -and
            ($Permissions[2].P2 -eq 'GroupName') -and
            ($Permissions[3].P1 -eq 'Path') -and
            ($Permissions[3].P2 -eq 'L') -and
            ($Permissions[4].P1 -eq 'Folder') -and
            ($Permissions[4].P2 -eq 'W')
        }
    }
}
Describe 'the script that tests the remote computers for compliance' {
    BeforeAll {
        Mock Test-ExpandedMatrixHC
        Mock Invoke-Command {
            'A'
        } -ParameterFilter {
            ($ComputerName -eq 'PC1') -and
            ($ConfigurationName) -and
            ($FilePath -eq $testParams.ScriptTestRequirements)
        }
        Mock Invoke-Command {
            'B'
        } -ParameterFilter {
            ($ComputerName -eq 'PC2') -and
            ($ConfigurationName) -and
            ($FilePath -eq $testParams.ScriptTestRequirements)
        }

        @(
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'PC1'
                Path         = 'E:\Department'
                Action       = 'Check'
            }
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'PC1'
                Path         = 'E:\Reports'
                Action       = 'Check'
            }
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'PC2'
                Path         = 'E:\Finance'
                Action       = 'Check'
            }
            [PSCustomObject]@{
                Status       = $null
                ComputerName = 'ignoredPc'
                Path         = 'E:\Finance'
                Action       = 'Check'
            }
        ) | Export-Excel @SettingsParams

        $testPermissions | Export-Excel @PermissionsParams

        .$testScript @testParams
    }
    It "is not called for rows in the 'Settings' worksheets where Status is not Enabled" {
        Should -Not -Invoke Invoke-Command -Scope Describe -ParameterFilter {
            ($FilePath -eq $testParams.ScriptTestRequirements) -and
            ($ComputerName -eq 'ignoredPc')
        }
    }
    It "is only called for unique ComputerNames in the 'Settings' worksheets" {
        Should -Invoke Invoke-Command -Times 2 -Exactly -Scope Describe -ParameterFilter {
            ($ConfigurationName) -and
            ($FilePath -eq $testParams.ScriptTestRequirements)
        }
        @('PC1', 'PC2') | ForEach-Object {
            Should -Invoke Invoke-Command -Times 1 -Exactly -Scope Describe -ParameterFilter {
                ($ConfigurationName) -and
                ($FilePath -eq $testParams.ScriptTestRequirements) -and
                ($ComputerName -eq $_)
            }
        }
    }
    It 'saves the job result in Settings for each matrix' {
        @($ImportedMatrix.Settings.Where(
                {
                    ($_.Import.ComputerName -eq 'PC1') -and
                    ($_.Check -eq 'A') }
            )
        ).Count |
        Should -BeExactly 2

        @($ImportedMatrix.Settings.Where(
                {
                    ($_.Import.ComputerName -eq 'PC2') -and
                    ($_.Check -eq 'B') }
            )
        ).Count |
        Should -BeExactly 1
    }
}
Describe 'the script that sets the permissions on the remote computers' {
    BeforeAll {
        Mock Test-ExpandedMatrixHC
        Mock Invoke-Command { 1 } -ParameterFilter {
            ($ComputerName -eq 'PC1') -and
            ($ArgumentList[0] -eq 'E:\Department') -and
            ($ArgumentList[1] -eq 'New') -and
            ($ArgumentList[2]) -and
            ($ArgumentList[3] -eq $testParams.MaxConcurrentFoldersPerMatrix) -and
            ($ConfigurationName) -and
            ($FilePath -eq $testParams.ScriptSetPermissionFile)
        }
        Mock Invoke-Command { 2 } -ParameterFilter {
            ($ComputerName -eq 'PC1') -and
            ($ArgumentList[0] -eq 'E:\Reports') -and
            ($ArgumentList[1] -eq 'Fix') -and
            ($ArgumentList[2]) -and
            ($ArgumentList[3] -eq $testParams.MaxConcurrentFoldersPerMatrix) -and
            ($ConfigurationName) -and
            ($FilePath -eq $testParams.ScriptSetPermissionFile)
        }
        Mock Invoke-Command { 3 } -ParameterFilter {
            ($ComputerName -eq 'PC2') -and
            ($ArgumentList[0] -eq 'E:\Finance') -and
            ($ArgumentList[1] -eq 'Check') -and
            ($ArgumentList[2]) -and
            ($ArgumentList[3] -eq $testParams.MaxConcurrentFoldersPerMatrix) -and
            ($ConfigurationName) -and
            ($FilePath -eq $testParams.ScriptSetPermissionFile)
        }

        @(
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'PC1'
                Path         = 'E:\Department'
                Action       = 'New'
            }
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'PC1'
                Path         = 'E:\Reports'
                Action       = 'Fix'
            }
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'PC2'
                Path         = 'E:\Finance'
                Action       = 'Check'
            }
            [PSCustomObject]@{
                Status       = $null
                ComputerName = 'ignoredPc'
                Path         = 'E:\Finance'
                Action       = 'Check'
            }
        ) | Export-Excel @SettingsParams

        $testPermissions | Export-Excel @PermissionsParams

        .$testScript @testParams
    }
    It "is not called for rows in the 'Settings' worksheets where Status is not Enabled" {
        Should -Not -Invoke Invoke-Command -Scope Describe -ParameterFilter {
            ($ComputerName -eq 'ignoredPc')
        }
    }
    It "is called for each row in the 'Settings' worksheets with Status Enabled" {
        Should -Invoke Invoke-Command -Times 3 -Exactly -Scope Describe -ParameterFilter {
            ($FilePath -eq $testParams.ScriptSetPermissionFile)
        }
        Should -Invoke Invoke-Command -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($FilePath -eq $testParams.ScriptSetPermissionFile) -and
            ($ComputerName -eq 'PC1') -and
            ($ArgumentList[0] -eq 'E:\Department') -and
            ($ArgumentList[1] -eq 'New') -and
            ($ArgumentList[2] -ne $null) -and
            ($ArgumentList[3] -ne $null)
        }
        Should -Invoke Invoke-Command -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($FilePath -eq $testParams.ScriptSetPermissionFile) -and
            ($ComputerName -eq 'PC1') -and
            ($ArgumentList[0] -eq 'E:\Reports') -and
            ($ArgumentList[1] -eq 'Fix') -and
            ($ArgumentList[2] -ne $null) -and
            ($ArgumentList[3] -ne $null)
        }
        Should -Invoke Invoke-Command -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($FilePath -eq $testParams.ScriptSetPermissionFile) -and
            ($ComputerName -eq 'PC2') -and
            ($ArgumentList[0] -eq 'E:\Finance') -and
            ($ArgumentList[1] -eq 'Check') -and
            ($ArgumentList[2] -ne $null) -and
            ($ArgumentList[3] -ne $null)
        }
    }
    It 'saves the start/end/duration times for each job in the settings' {
        $ImportedMatrix.Settings.JobTime.Start | Should -HaveCount 3
        $ImportedMatrix.Settings.JobTime.End | Should -HaveCount 3
        $ImportedMatrix.Settings.JobTime.Duration | Should -HaveCount 3
    }
    It 'saves the job result in Settings for each matrix' {
        ($ImportedMatrix.Settings.Where( { ($_.ID -eq 1) })).Check |
        Should -Contain 1
        ($ImportedMatrix.Settings.Where( { ($_.ID -eq 2) })).Check |
        Should -Contain 2
        ($ImportedMatrix.Settings.Where( { ($_.ID -eq 3) })).Check |
        Should -Contain 3
    }
}
Describe 'an email is sent to the user in the default settings file' {
    BeforeAll {
        Mock Test-ExpandedMatrixHC
        @(
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'PC1'
                Path         = 'E:\Reports'
                Action       = 'Check'
                GroupName    = 'C'
                SiteCode     = 'D'
            }
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'PC2'
                Path         = 'E:\Finance'
                Action       = 'New'
                GroupName    = 'x'
                SiteCode     = 'x'
            }
        ) | Export-Excel @SettingsParams

        $testPermissions | Export-Excel @PermissionsParams

        .$testScript @testParams
    }
    It 'containing a summary per Settings row for executed matrixes' {
        Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($To -eq 'Bob@contoso.com') -and
            ($Subject -eq '1 matrix file') -and
            ($Save -like "$((Get-Item $testParams.LogFolder).FullName)* - Mail - 1 matrix file.html") -and
            ($Priority -eq 'Normal') -and
            ($Message -notlike '*Cherwell*') -and
            ($Message -like '*Matrix results per file*') -and
            ($Message -like '*Matrix.xlsx*') -and
            ($Message -like '*Settings*') -and
            ($Message -like '*ID*ComputerName*Path*Action*Duration*') -and
            ($Message -like '*1*PC1*E:\Reports*Check*') -and
            ($Message -like '*2*PC2*E:\Finance*New*') -and
            ($Message -like '*Error*Warning*Information*')
        }
    }
}
Describe 'export an Excel file with' {
    BeforeAll {
        Mock Get-ADObjectDetailHC {
            [PSCustomObject]@{
                samAccountName = 'A B bob'
                adObject       = @{
                    ObjectClass    = 'user'
                    Name           = 'A B Bob'
                    SamAccountName = 'A B bob'
                    ManagedBy      = $null
                }
                adGroupMember  = $null
            }
            [PSCustomObject]@{
                samAccountName = 'movieStars'
                adObject       = @{
                    ObjectClass    = 'group'
                    Name           = 'Movie Stars'
                    SamAccountName = 'movieStars'
                    ManagedBy      = $null
                }
                adGroupMember  = $null
            }
            [PSCustomObject]@{
                samAccountName = 'starTrekCaptains'
                adObject       = @{
                    ObjectClass    = 'group'
                    SamAccountName = 'starTrekCaptains'
                    Name           = 'Star Trek Captains'
                    ManagedBy      = 'CN=CaptainManagers,DC=contoso,DC=net'
                }
                adGroupMember  = @(
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Jean Luc Picard'
                        SamAccountName = 'picard'
                    }
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Ignored account'
                        SamAccountName = 'ignoreMe'
                    }
                )
            }
            [PSCustomObject]@{
                samAccountName = 'singers'
                adObject       = @{
                    ObjectClass    = 'group'
                    SamAccountName = 'singers'
                    Name           = 'Singers'
                    ManagedBy      = 'CN=SingerManagers,DC=contoso,DC=net'
                }
                adGroupMember  = @(
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Beyonce'
                        SamAccountName = 'queenb'
                    }
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Ignored account'
                        SamAccountName = 'ignoreMe'
                    }
                )
            }
        } -ParameterFilter { $Type -eq 'SamAccountName' }
        Mock Get-ADObjectDetailHC {
            [PSCustomObject]@{
                DistinguishedName = 'CN=CaptainManagers,DC=contoso,DC=net'
                adObject          = @{
                    ObjectClass = 'group'
                    Name        = 'Captain Managers'
                }
                adGroupMember     = @(
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Admiral Pike'
                        SamAccountName = 'pike'
                    }
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Excluded user'
                        SamAccountName = 'ignoreMe'
                    }
                )
            }
            [PSCustomObject]@{
                DistinguishedName = 'CN=SingerManagers,DC=contoso,DC=net'
                adObject          = @{
                    ObjectClass = 'group'
                    Name        = 'Singer Managers'
                }
                adGroupMember     = $null
            }
        } -ParameterFilter { $Type -eq 'DistinguishedName' }

        @(
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'PC1'
                Path         = 'E:\Reports'
                Action       = 'Check'
                GroupName    = 'A'
                SiteCode     = 'B'
            }
        ) | Export-Excel @SettingsParams

        @(
            [PSCustomObject]@{
                P1 = $null      ; P2 = 'bob'       ; P3 = 'movieStars'; P4 = '' ; P5 = ''
            }
            [PSCustomObject]@{
                P1 = 'SiteCode' ; P2 = 'SiteCode'  ; P3 = ''; P4 = 'starTrekCaptains' ; P5 = ''
            }
            [PSCustomObject]@{
                P1 = 'GroupName'; P2 = 'GroupName' ; P3 = ''; P4 = '' ; P5 = 'Singers'
            }
            [PSCustomObject]@{
                P1 = 'Path'     ; P2 = 'L'         ; P3 = ''; P4 = '' ; P5 = ''
            }
            [PSCustomObject]@{
                P1 = 'Folder'   ; P2 = 'W'         ; P3 = ''; P4 = '' ; P5 = ''
            }
        ) | Export-Excel @PermissionsParams

        .$testScript @testParams -ExcludedSamAccountName 'IgnoreMe'

        $testMatrixFile = Get-ChildItem $testParams.logFolder -Filter '*Matrix.xlsx' -Recurse -File
    }
    Context "the worksheet 'AccessList'" {
        BeforeAll {
            $testExportedExcelRows = @(
                @{
                    SamAccountName       = 'starTrekCaptains'
                    Name                 = 'Star Trek Captains'
                    Type                 = 'group'
                    MemberName           = 'Jean Luc Picard'
                    MemberSamAccountName = 'picard'
                }
                @{
                    SamAccountName       = 'A B bob'
                    Name                 = 'A B Bob'
                    Type                 = 'user'
                    MemberName           = $null
                    MemberSamAccountName = $null
                }
                @{
                    SamAccountName       = 'Singers'
                    Name                 = 'Singers'
                    Type                 = 'group'
                    MemberName           = 'Beyonce'
                    MemberSamAccountName = 'queenb'
                }
                @{
                    SamAccountName       = 'movieStars'
                    Name                 = 'Movie Stars'
                    Type                 = 'group'
                    MemberName           = $null
                    MemberSamAccountName = $null
                }
            )

            $actual = Import-Excel -Path $testMatrixFile.FullName -WorksheetName 'AccessList'
        }
        It 'added to the matrix log file' {
            $actual | Should -Not -BeNullOrEmpty
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    $_.SamAccountName -eq $testRow.SamAccountName
                }
                $actualRow.Name | Should -Be $testRow.Name
                $actualRow.Type | Should -BeLike $testRow.Type
                $actualRow.MemberName | Should -Be $testRow.MemberName
                $actualRow.MemberSamAccountName | Should -BeLike $testRow.MemberSamAccountName
            }
        }
    }
    Context "the worksheet 'GroupManagers'" {
        BeforeAll {
            $testExportedExcelRows = @(
                @{
                    GroupName         = 'Star Trek Captains'
                    ManagerName       = 'Captain Managers'
                    ManagerType       = 'group'
                    ManagerMemberName = 'Admiral Pike'
                }
                @{
                    GroupName         = 'Singers'
                    ManagerName       = 'Singer Managers'
                    ManagerType       = 'group'
                    ManagerMemberName = $null
                }
                @{
                    GroupName         = 'Movie Stars'
                    ManagerName       = $null
                    ManagerType       = $null
                    ManagerMemberName = $null
                }
            )

            $actual = Import-Excel -Path $testMatrixFile.FullName -WorksheetName 'GroupManagers'
        }
        It 'added to the matrix log file' {
            $actual | Should -Not -BeNullOrEmpty
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    $_.GroupName -eq $testRow.GroupName
                }
                $actualRow.ManagerName | Should -Be $testRow.ManagerName
                $actualRow.ManagerType | Should -BeLike $testRow.ManagerType
                $actualRow.ManagerMemberName | Should -Be $testRow.ManagerMemberName
            }
        }
    }
}
Describe 'when a job fails' {
    Context 'the test requirements script' {
        BeforeAll {
            Mock Test-ExpandedMatrixHC
            Mock Invoke-Command { throw 'failure' } -ParameterFilter {
                ($ComputerName -eq 'PC1') -and
                ($ConfigurationName) -and
                ($FilePath -eq $testParams.ScriptTestRequirements)
            }
            Mock Invoke-Command { 'B' } -ParameterFilter {
                ($ComputerName -eq 'PC2') -and
                ($ConfigurationName) -and
                ($FilePath -eq $testParams.ScriptTestRequirements)
            }

            @(
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = 'PC1'
                    Path         = 'E:\Department'
                    Action       = 'Check'
                }
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = 'PC2'
                    Path         = 'E:\Reports'
                    Action       = 'Check'
                }
            ) | Export-Excel @SettingsParams

            $testPermissions | Export-Excel @PermissionsParams

            .$testScript @testParams
        }
        It 'the job error is saved in Settings for each matrix' {
            $actual = $ImportedMatrix.Settings.Where( { ($_.ID -eq 1) })
            $actual.Check.Type | Should -Be 'FatalError'
            $actual.Check.Value | Should -Be 'failure'

            $actual = $ImportedMatrix.Settings.Where( { ($_.ID -eq 2) })
            $actual.Check.Type | Should -Not -Be 'FatalError'
            $actual.Check.Value | Should -Not -Be 'failure'
        }
    }
    Context 'the set permissions script' {
        BeforeAll {
            Mock Test-ExpandedMatrixHC
            Mock Invoke-Command { 1 } -ParameterFilter {
                ($ConfigurationName) -and
                ($ArgumentList[0] -eq 'E:\Department') -and
                ($FilePath -eq $testParams.ScriptSetPermissionFile)
            }
            Mock Invoke-Command { throw 'failure' } -ParameterFilter {
                ($ConfigurationName) -and
                ($ArgumentList[0] -eq 'E:\Reports') -and
                ($FilePath -eq $testParams.ScriptSetPermissionFile)
            }

            @(
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = 'PC1'
                    Path         = 'E:\Department'
                    Action       = 'Check'
                }
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = 'PC1'
                    Path         = 'E:\Reports'
                    Action       = 'Check'
                }
            ) | Export-Excel @SettingsParams

            $testPermissions | Export-Excel @PermissionsParams

            .$testScript @testParams
        }
        It 'the job error is saved in Settings for each matrix' {
            $actual = $ImportedMatrix.Settings.Where( { ($_.ID -eq 1) })
            $actual.Check.Type | Should -Not -Be 'FatalError'
            $actual.Check.Value | Should -Not -Be 'failure'

            $actual = $ImportedMatrix.Settings.Where( { ($_.ID -eq 2) })
            $actual.Check.Type | Should -Be 'FatalError'
            $actual.Check.Value | Should -Be 'failure'
        }
    }
}
Describe 'internal functions' {
    Context 'default permissions vs matrix permissions' {
        It 'add default permissions to the matrix' {
            Mock Test-ExpandedMatrixHC
            Mock ConvertTo-MatrixAclHC {
                @(
                    [PSCustomObject]@{
                        Path   = 'Path'
                        ACL    = @{'Mike' = 'L' }
                        Parent = $true
                        Ignore = $false
                    }
                )
            }
            Mock Get-DefaultAclHC {
                @{
                    'Bob' = 'R'
                }
            }
            @(
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = 'test'
                    Path         = 'E:\Department'
                    Action       = 'Check'
                }
            ) | Export-Excel @SettingsParams
            @(
                [PSCustomObject]@{P1 = $null      ; P2 = 'Mike' }
                [PSCustomObject]@{P1 = 'SiteCode' ; P2 = '' }
                [PSCustomObject]@{P1 = 'GroupName'; P2 = '' }
                [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
                [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
            ) | Export-Excel @PermissionsParams

            .$testScript @testParams

            $actual = ($ImportedMatrix.Settings.Matrix.Where(
                    { $_.Path -eq 'Path' })
            ).ACL

            $expected = @{
                'Bob'  = 'R'
                'Mike' = 'L'
            }

            Compare-HashTableHC $actual $expected
        }
        It 'do not add default permissions to the matrix ACL when the folder has no ACL' {
            Mock Test-ExpandedMatrixHC
            Mock ConvertTo-MatrixAclHC {
                @(
                    [PSCustomObject]@{
                        Path   = 'Path'
                        ACL    = @{}
                        Parent = $true
                        Ignore = $false
                    }
                )
            }
            Mock Get-DefaultAclHC {
                @{
                    'Bob' = 'R'
                }
            }

            @(
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = 'test'
                    Path         = 'E:\Department'
                    Action       = 'Check'
                }
            ) | Export-Excel @SettingsParams
            @(
                [PSCustomObject]@{P1 = $null      ; P2 = 'Mike' }
                [PSCustomObject]@{P1 = 'SiteCode' ; P2 = '' }
                [PSCustomObject]@{P1 = 'GroupName'; P2 = '' }
                [PSCustomObject]@{P1 = 'Path'     ; P2 = '' }
                [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'L' }
            ) | Export-Excel @PermissionsParams

            .$testScript @testParams

            $actual = ($ImportedMatrix.Settings.Matrix.Where( {
                        $_.Path -eq 'Path' })).ACL

            $actual | Should -BeNullOrEmpty
        }
        It 'do not overwrite permissions to the matrix ACL when they are also in the default ACL' {
            Mock Test-ExpandedMatrixHC
            Mock ConvertTo-MatrixAclHC {
                @(
                    [PSCustomObject]@{
                        Path   = 'Path'
                        ACL    = @{
                            'Mike' = 'L'
                            'Bob'  = 'L'
                        }
                        Parent = $true
                        Ignore = $false
                    }
                )
            }
            Mock Get-DefaultAclHC {
                @{
                    'Bob' = 'R'
                }
            }
            @(
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = 'test'
                    Path         = 'E:\Department'
                    Action       = 'Check'
                }
            ) | Export-Excel @SettingsParams
            @(
                [PSCustomObject]@{P1 = $null      ; P2 = 'Mike' }
                [PSCustomObject]@{P1 = 'SiteCode' ; P2 = '' }
                [PSCustomObject]@{P1 = 'GroupName'; P2 = '' }
                [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
                [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
            ) | Export-Excel @PermissionsParams

            .$testScript @testParams

            $actual = ($ImportedMatrix.Settings.Matrix.Where( { $_.Path -eq 'Path' })).ACL

            $expected = @{
                'Bob'  = 'L'
                'Mike' = 'L'
            }

            Compare-HashTableHC $actual $expected
        }
    }
}
Describe 'when a FatalError occurs while executing the matrix' {
    AfterEach {
        $Error.Clear()
        Remove-Item -Path "$($testParams.LogFolder)\*" -Recurse -Force -EA Ignore
        Remove-Item -Path "$($testParams.Matrix.FolderPath)\*" -Exclude $TestDefaultsFileName -Recurse -Force -EA Ignore
    }
    It 'a detailed HTML log file is created for each settings row' {
        $testProblem = @{
            Type = 'FatalError'
            Name = 'Matrix setting incorrect'
        }
        Mock Test-MatrixSettingHC {
            $testProblem
        }

        @(
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'S1'
                Path         = 'E:\Department'
                Action       = 'Check'
            }
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'S2'
                Path         = 'E:\Department'
                Action       = 'Check'
            }
        ) | Export-Excel @SettingsParams
        $testPermissions | Export-Excel @PermissionsParams

        .$testScript @testParams

        $testMatrixLogFolder = Get-ChildItem -Path $testParams.LogFolder -Directory
        @(Get-ChildItem -Path $testMatrixLogFolder.FullName -File | Where-Object Extension -NE '.xlsx').Count | Should -BeExactly 2
    }
    It 'a TXT log file is created for each settings row when there are more than 5 elements in the value array' {
        $testProblem = @{
            Type        = 'FatalError'
            Name        = 'Matrix setting incorrect'
            Description = 'When things go south we need to report it. because it needs to get fixed'
            Value       = @('C:\Folder1', 'C:\Folder1', 'C:\Folder1', 'C:\Folder1', 'C:\Folder1', 'C:\Folder1', 'C:\Folder1', 'C:\Folder1')
        }
        Mock Test-MatrixSettingHC {
            $testProblem
        }

        @(
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'S1'
                Path         = 'E:\Department'
                Action       = 'Check'
                GroupName    = 'Group'
                SiteCode     = 'Site'
            }
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'S2'
                Path         = 'E:\Department'
                Action       = 'Check'
            }
        ) | Export-Excel @SettingsParams
        $testPermissions | Export-Excel @PermissionsParams

        .$testScript @testParams

        $testMatrixLogFolder = Get-ChildItem -Path $testParams.LogFolder -Directory
        @(Get-ChildItem -Path $testMatrixLogFolder.FullName -File | Where-Object Extension -EQ '.txt').Count | Should -BeExactly 2
    }
    It 'an e-mail is send' {
        $testProblem = @{
            Type        = 'FatalError'
            Name        = 'Matrix setting incorrect'
            Description = 'When things go south we need to report it. because it needs to get fixed'
            Value       = @('C:\Folder1', 'C:\Folder1', 'C:\Folder1', 'C:\Folder1', 'C:\Folder1', 'C:\Folder1', 'C:\Folder1', 'C:\Folder1')
        }
        Mock Test-MatrixSettingHC {
            $testProblem
        }

        @(
            [PSCustomObject]@{Status = 'Enabled'
                ComputerName         = 'S1'
                Path                 = 'E:\Department'
                Action               = 'Check'
                GroupName            = 'Group'
                SiteCode             = 'Site'
            }
            [PSCustomObject]@{Status = 'Enabled'
                ComputerName         = 'S2'
                Path                 = 'E:\Department'
                Action               = 'Check'
            }
        ) | Export-Excel @SettingsParams
        $testPermissions | Export-Excel @PermissionsParams

        .$testScript @testParams

        Should -Invoke Send-MailHC -Scope it -Times 1 -Exactly
    }
}
Describe 'when the argument CherwellFolder is used' {
    Context 'but the Excel file is missing the sheet FormData' {
        BeforeAll {
            @(
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = 'PC1'
                    Path         = 'E:\Department'
                    Action       = 'Check'
                }
            ) | Export-Excel @SettingsParams

            $testPermissions | Export-Excel @PermissionsParams

            .$testScript @testParams -CherwellFolder $testCherwellFolder.FullName
        }
        It 'a FatalError is registered for the file' {
            $actual = $ImportedMatrix.File.Check
            $actual.Type | Should -Contain 'FatalError'
            $actual.Name | Should -Contain "Worksheet 'FormData' not found"
        }
        It 'the permissions script is not executed' {
            Should -Not -Invoke Invoke-Command
        }
        It 'an email is sent to the user with the error' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
                ($To -eq 'Bob@contoso.com') -and
                ($Save -like "$((Get-Item $testParams.LogFolder).FullName)* - Mail - 1 matrix file, 1 error.html") -and
                ($Subject -eq '1 matrix file, 1 error') -and
                ($Priority -eq 'High') -and
                ($Message -like "*Worksheet 'FormData' not found*") -and
                ($Message -notlike '*Check the*overview*for details*')
            }
        }
    }
    Context 'but the worksheet FormData contains incorrect data' {
        AfterAll {
            Mock Test-FormDataHC
            Mock Get-AdUserPrincipalNameHC
        }
        BeforeAll {
            Mock Test-ExpandedMatrixHC
            Mock Test-FormDataHC {
                @{
                    Type = 'FatalError'
                    Name = 'incorrect data'
                }
            }

            @(
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = 'PC1'
                    Path         = 'E:\Department'
                    Action       = 'Check'
                }
            ) | Export-Excel @SettingsParams

            @(
                [PSCustomObject]@{
                    MatrixFormStatus = 'Enabled'
                }
            ) |
            Export-Excel -Path $SettingsParams.Path -WorksheetName 'FormData'

            $testPermissions | Export-Excel @PermissionsParams

            .$testScript @testParams -CherwellFolder $testCherwellFolder.FullName
        }
        It 'a FatalError is registered for the FormData sheet' {
            $actual = $ImportedMatrix.FormData.Check
            $actual.Type | Should -Contain 'FatalError'
            $actual.Name | Should -Contain 'incorrect data'
        }
        It 'the permissions script is not executed' {
            Should -Not -Invoke Invoke-Command
        }
        It 'an email is sent to the user with the error' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
                ($To -eq 'Bob@contoso.com') -and
                ($Save -like "$((Get-Item $testParams.LogFolder).FullName)* - Mail - 1 matrix file, 1 error.html") -and
                ($Subject -eq '1 matrix file, 1 error') -and
                ($Priority -eq 'High') -and
                ($Message -like '*Errors*Warnings*FormData*') -and
                ($Message -like '*FormData*incorrect data*') -and
                ($Message -notlike '*Check the*overview*for details*')
            }
        }
    }
    Context 'but the worksheet FormData has a non existing MatrixResponsible' {
        BeforeAll {
            Mock Test-ExpandedMatrixHC
            Mock Test-FormDataHC
            Mock Get-AdUserPrincipalNameHC {
                @{
                    UserPrincipalName = 'mike@contoso.com'
                    notFound          = 'bob@contoso.com'
                }
            }

            @(
                [PSCustomObject]@{
                    Status       = 'Enabled'
                    ComputerName = 'PC1'
                    Path         = 'E:\Department'
                    Action       = 'Check'
                }
            ) | Export-Excel @SettingsParams

            @(
                [PSCustomObject]@{
                    MatrixFormStatus  = 'Enabled'
                    MatrixResponsible = 'mike@contoso.com, bob@contoso.com'
                }
            ) |
            Export-Excel -Path $SettingsParams.Path -WorksheetName 'FormData'

            $testPermissions | Export-Excel @PermissionsParams

            .$testScript @testParams -CherwellFolder $testCherwellFolder.FullName
        }
        It 'a Warning is registered for the FormData sheet' {
            $actual = $ImportedMatrix.FormData.Check
            $actual.Type | Should -Contain 'Warning'
            $actual.Name | Should -Contain 'AD object not found'
            $actual.Value | Should -Contain 'bob@contoso.com'
        }
        It 'the permissions script is not executed' {
            Should -Not -Invoke Invoke-Command
        }
        It 'an email is sent to the user with the warning message' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
                ($To -eq 'Bob@contoso.com') -and
                ($Save -like "$((Get-Item $testParams.LogFolder).FullName)* - Mail - 1 matrix file, 1 warning.html") -and
                ($Subject -eq '1 matrix file, 1 warning') -and
                ($Priority -eq 'High') -and
                ($Message -like '*Errors*Warnings*FormData*') -and
                ($Message -like '*FormData*AD object not found*') -and
                ($Message -like '*Check the*overview*for details*')
            }
        }
    }
}
Describe 'when the argument CherwellFolder is used on a successful run' {
    BeforeAll {
        Mock Test-ExpandedMatrixHC
        Mock Get-AdUserPrincipalNameHC {
            @{
                UserPrincipalName = @('bob@contoso.com', 'mike@contoso.com')
                notFound          = $null
            }
        }
        Mock Test-FormDataHC
        Mock Get-ADObjectDetailHC {
            [PSCustomObject]@{
                samAccountName = 'A B C'
                adObject       = @{
                    ObjectClass    = 'group'
                    Name           = 'A B C'
                    SamAccountName = 'A B c'
                    ManagedBy      = 'CN=CaptainManagers,DC=contoso,DC=net'
                }
                adGroupMember  = @(
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Jean Luc Picard'
                        SamAccountName = 'picard'
                    }
                )
            }
        } -ParameterFilter { $Type -eq 'SamAccountName' }
        Mock Get-ADObjectDetailHC {
            [PSCustomObject]@{
                DistinguishedName = 'CN=CaptainManagers,DC=contoso,DC=net'
                adObject          = @{
                    ObjectClass = 'group'
                    Name        = 'Captain Managers'
                }
                adGroupMember     = @(
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Admiral Pike'
                        SamAccountName = 'pike'
                    }
                )
            }
        } -ParameterFilter { $Type -eq 'DistinguishedName' }

        @(
            [PSCustomObject]@{P1 = $null      ; P2 = 'C' }
            [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
            [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
            [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
            [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
        ) | Export-Excel @PermissionsParams

        @(
            [PSCustomObject]@{
                Status       = 'Enabled'
                ComputerName = 'SERVER1'
                GroupName    = 'A'
                SiteCode     = 'B'
                Path         = 'E:\Department'
                Action       = 'Check'
            }
        ) | Export-Excel @SettingsParams

        @(
            [PSCustomObject]@{
                MatrixFormStatus        = 'Enabled'
                MatrixCategoryName      = 'a'
                MatrixSubCategoryName   = 'b'
                MatrixResponsible       = 'c'
                MatrixFolderDisplayName = 'd'
                MatrixFolderPath        = 'e'
            }
        ) |
        Export-Excel -Path $SettingsParams.Path -WorksheetName 'FormData'

        $testCherwellParams = @{
            CherwellFolder                = $testCherwellFolder.FullName
            CherwellAdObjectsFileName     = 'AD object names.csv'
            CherwellFormDataFileName      = 'Form data.csv'
            CherwellGroupManagersFileName = 'GroupManagers.csv'
            CherwellAccessListFileName    = 'AccessList.csv'
            CherwellExcelOverviewFileName = 'Overview.xlsx'
        }
        .$testScript @testParams @testCherwellParams

        $testCherwellExport = Get-ChildItem $testCherwellFolder.FullName

        $testLogFolderExport = Get-ChildItem $testParams.LogFolder -Recurse -File |
        Where-Object Extension -Match '.xlsx$|.csv$'

        $testLogFolder = @{
            ExcelFile            = $testLogFolderExport |
            Where-Object Name -Like '*Overview.xlsx'
            FormDataCsvFile      = $testLogFolderExport |
            Where-Object Name -Like '*Form data.csv'
            AdObjectsCsvFile     = $testLogFolderExport |
            Where-Object Name -Like '*AD object names.csv'
            GroupManagersCsvFile = $testLogFolderExport |
            Where-Object Name -Like '*GroupManagers.csv'
            AccessListCsvFile    = $testLogFolderExport |
            Where-Object Name -Like '*AccessList.csv'
        }

        $testCherwellFolder = @{
            ExcelFile            = $testCherwellExport |
            Where-Object Name -Like '*Overview.xlsx'
            FormDataCsvFile      = $testCherwellExport |
            Where-Object Name -EQ 'Form data.csv'
            AdObjectsCsvFile     = $testCherwellExport |
            Where-Object Name -EQ 'AD object names.csv'
            GroupManagersCsvFile = $testCherwellExport |
            Where-Object Name -Like '*GroupManagers.csv'
            AccessListCsvFile    = $testCherwellExport |
            Where-Object Name -Like '*AccessList.csv'
        }
    }
    Context 'the data in worksheet FormData' {
        It 'is verified to be correct' {
            Should -Invoke Test-FormDataHC -Exactly 1 -Scope Describe
        }
        It 'is retrieved regardless the MatrixFormStatus' {
            $actual = $ImportedMatrix.FormData.Import
            $actual | Should -Not -BeNullOrEmpty
        }
        It 'the MatrixResponsible is converted to UserPrincipalName' {
            Should -Invoke Get-AdUserPrincipalNameHC -Exactly 1 -Scope Describe
        }
    }
    Context 'the FormData is exported' {
        It 'to a CSV file in the Cherwell folder' {
            $testCherwellFolder.FormDataCsvFile.FullName |
            Should -Not -BeNullOrEmpty
        }
        It 'to a CSV file in the log folder' {
            $testLogFolder.FormDataCsvFile.FullName |
            Should -Not -BeNullOrEmpty
        }
        It 'to an Excel file in the Cherwell folder' {
            $testCherwellFolder.ExcelFile.FullName | Should -Not -BeNullOrEmpty
        }
        It 'to an HTML file in the Cherwell folder' {
            $testCherwellFolder.ExcelFile.FullName.Replace('.xlsx', '.html') | Should -Exist
        }
        It 'to an Excel file in the log folder' {
            $testLogFolder.ExcelFile.FullName | Should -Not -BeNullOrEmpty
        }
        Context 'with the property' {
            BeforeAll {
                $actual = @{
                    logFolder      = @{
                        Excel    = Import-Excel -Path $testLogFolder.ExcelFile.FullName -WorksheetName 'FormData'
                        FormData = Import-Csv -Path $testLogFolder.FormDataCsvFile.FullName
                    }
                    cherwellFolder = @{
                        Excel    = Import-Excel -Path $testCherwellFolder.ExcelFile.FullName -WorksheetName 'FormData'
                        FormData = Import-Csv -Path $testCherwellFolder.FormDataCsvFile.FullName
                    }
                }
            }
            It '<Name>' -ForEach @(
                @{ Name = 'MatrixFormStatus'; Value = 'Enabled' }
                @{ Name = 'MatrixFileName'; Value = 'Matrix' }
                # @{ Name = 'MatrixFilePath'; Value = $SettingsParams.Path }
                @{ Name = 'MatrixCategoryName'; Value = 'a' }
                @{ Name = 'MatrixSubCategoryName'; Value = 'b' }
                @{ Name = 'MatrixResponsible'; Value = 'bob@contoso.com,mike@contoso.com' }
                @{ Name = 'MatrixFolderDisplayName'; Value = 'd' }
                @{ Name = 'MatrixFolderPath'; Value = 'e' }
            ) {
                $actual.cherwellFolder.FormData.$Name | Should -Be $Value
                $actual.cherwellFolder.Excel.$Name | Should -Be $Value
                $actual.logFolder.FormData.$Name | Should -Be $Value
                $actual.logFolder.Excel.$Name | Should -Be $Value
            }
            It 'MatrixFilePath' {
                # scoping issue in Pester
                $actual.cherwellFolder.FormData.MatrixFilePath |
                Should -Be $SettingsParams.Path
                $actual.cherwellFolder.Excel.MatrixFilePath |
                Should -Be $SettingsParams.Path
                $actual.logFolder.FormData.MatrixFilePath |
                Should -Be $SettingsParams.Path
                $actual.logFolder.Excel.MatrixFilePath |
                Should -Be $SettingsParams.Path
            }
        }
    }
    Context 'the AD object names are exported' {
        It 'to a CSV file in the Cherwell folder' {
            $testCherwellFolder.AdObjectsCsvFile.FullName |
            Should -Not -BeNullOrEmpty
        }
        It 'to a CSV file in the log folder' {
            $testLogFolder.AdObjectsCsvFile.FullName |
            Should -Not -BeNullOrEmpty
        }
        It 'to an Excel file in the Cherwell folder' {
            $testCherwellFolder.ExcelFile.FullName | Should -Not -BeNullOrEmpty
        }
        It 'to an Excel file in the log folder' {
            $testLogFolder.ExcelFile.FullName | Should -Not -BeNullOrEmpty
        }
        Context 'with the property' {
            BeforeAll {
                $actual = @{
                    logFolder      = @{
                        Excel         = Import-Excel -Path $testLogFolder.ExcelFile.FullName -WorksheetName 'AdObjectNames'
                        AdObjectNames = Import-Csv -Path $testLogFolder.AdObjectsCsvFile.FullName
                    }
                    cherwellFolder = @{
                        Excel         = Import-Excel -Path $testCherwellFolder.ExcelFile.FullName -WorksheetName 'AdObjectNames'
                        AdObjectNames = Import-Csv -Path $testCherwellFolder.AdObjectsCsvFile.FullName
                    }
                }
            }
            It '<Name>' -ForEach @(
                @{ Name = 'MatrixFileName'; Value = 'Matrix' }
                @{ Name = 'SamAccountName'; Value = 'A B C' }
                @{ Name = 'GroupName'; Value = 'A' }
                @{ Name = 'SiteCode'; Value = 'B' }
                @{ Name = 'Name'; Value = 'C' }
            ) {
                $actual.cherwellFolder.AdObjectNames.$Name | Should -Be $Value
                $actual.cherwellFolder.Excel.$Name | Should -Be $Value
                $actual.logFolder.AdObjectNames.$Name | Should -Be $Value
                $actual.logFolder.Excel.$Name | Should -Be $Value
            }
        }
    }
    Context 'the GroupManagers are exported' {
        It 'to a CSV file in the Cherwell folder' {
            $testCherwellFolder.GroupManagersCsvFile.FullName |
            Should -Not -BeNullOrEmpty
        }
        It 'to a CSV file in the log folder' {
            $testLogFolder.GroupManagersCsvFile.FullName |
            Should -Not -BeNullOrEmpty
        }
        It 'to an Excel file in the Cherwell folder' {
            $testCherwellFolder.ExcelFile.FullName | Should -Not -BeNullOrEmpty
        }
        It 'to an Excel file in the log folder' {
            $testLogFolder.ExcelFile.FullName | Should -Not -BeNullOrEmpty
        }
        Context 'with the property' {
            BeforeAll {
                $actual = @{
                    logFolder      = @{
                        Excel         = Import-Excel -Path $testLogFolder.ExcelFile.FullName -WorksheetName 'GroupManagers'
                        GroupManagers = Import-Csv -Path $testLogFolder.GroupManagersCsvFile.FullName
                    }
                    cherwellFolder = @{
                        Excel         = Import-Excel -Path $testCherwellFolder.ExcelFile.FullName -WorksheetName 'GroupManagers'
                        GroupManagers = Import-Csv -Path $testCherwellFolder.GroupManagersCsvFile.FullName
                    }
                }
            }
            It '<Name>' -ForEach @(
                @{ Name = 'MatrixFileName'; Value = 'Matrix' }
                @{ Name = 'GroupName'; Value = 'A B C' }
                @{ Name = 'ManagerName'; Value = 'Captain Managers' }
                @{ Name = 'ManagerType'; Value = 'group' }
                @{ Name = 'ManagerMemberName'; Value = 'Admiral Pike' }
            ) {
                $actual.cherwellFolder.GroupManagers.$Name | Should -Be $Value
                $actual.cherwellFolder.Excel.$Name | Should -Be $Value
                $actual.logFolder.GroupManagers.$Name | Should -Be $Value
                $actual.logFolder.Excel.$Name | Should -Be $Value
            }
        }
    }
    Context 'the AccessList are exported' {
        It 'to a CSV file in the Cherwell folder' {
            $testCherwellFolder.AccessListCsvFile.FullName |
            Should -Not -BeNullOrEmpty
        }
        It 'to a CSV file in the log folder' {
            $testLogFolder.AccessListCsvFile.FullName |
            Should -Not -BeNullOrEmpty
        }
        It 'to an Excel file in the Cherwell folder' {
            $testCherwellFolder.ExcelFile.FullName | Should -Not -BeNullOrEmpty
        }
        It 'to an Excel file in the log folder' {
            $testLogFolder.ExcelFile.FullName | Should -Not -BeNullOrEmpty
        }
        Context 'with the property' {
            BeforeAll {
                $actual = @{
                    logFolder      = @{
                        Excel      = Import-Excel -Path $testLogFolder.ExcelFile.FullName -WorksheetName 'AccessList'
                        AccessList = Import-Csv -Path $testLogFolder.AccessListCsvFile.FullName
                    }
                    cherwellFolder = @{
                        Excel      = Import-Excel -Path $testCherwellFolder.ExcelFile.FullName -WorksheetName 'AccessList'
                        AccessList = Import-Csv -Path $testCherwellFolder.AccessListCsvFile.FullName
                    }
                }
            }
            It '<Name>' -ForEach @(
                @{ Name = 'MatrixFileName'; Value = 'Matrix' }
                @{ Name = 'SamAccountName'; Value = 'A B C' }
                @{ Name = 'Name'; Value = 'A B C' }
                @{ Name = 'Type'; Value = 'group' }
                @{ Name = 'MemberName'; Value = 'Jean Luc Picard' }
                @{ Name = 'MemberSamAccountName'; Value = 'picard' }
            ) {
                $actual.cherwellFolder.AccessList.$Name | Should -Be $Value
                $actual.cherwellFolder.Excel.$Name | Should -Be $Value
                $actual.logFolder.AccessList.$Name | Should -Be $Value
                $actual.logFolder.Excel.$Name | Should -Be $Value
            }
        }
    }
    It 'an email is sent to the user in the default settings file' {
        Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($To -eq 'Bob@contoso.com') -and
            ($Save -like "$((Get-Item $testParams.LogFolder).FullName)* - Mail - 1 matrix file.html") -and
            ($Subject -eq '1 matrix file') -and
            ($Priority -eq 'Normal') -and
            ($Message -like '*Export to*Cherwell*') -and
            ($Message -like '*Check the*overview*for details*') -and
            ($Message -like '*Access list*1*') -and
            ($Message -like '*AD objects*2*') -and
            ($Message -like '*Group managers*1*') -and
            ($Message -like '*Form data*1*') -and
            ($Message -like '*Matrix results per file*') -and
            ($Message -like '*Matrix.xlsx*') -and
            ($Message -like '*Settings*') -and
            ($Message -like '*ID*ComputerName*Path*Action*Duration*') -and
            ($Message -like '*1*SERVER1*E:\Department*Check*') -and
            ($Message -like '*Error*Warning*Information*')
        }
    }
}