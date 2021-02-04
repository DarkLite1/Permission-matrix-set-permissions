#Requires -Version 5.1
#Requires -Modules Assert, Pester, ImportExcel

BeforeAll {
    Import-Module 'T:\Test\Brecht\PowerShell\Toolbox.PermissionMatrix\Toolbox.PermissionMatrix.psm1' -Verbose

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')

    $TestInvokeCommand = Get-Command -Name Invoke-Command

    $ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN

    $MailAdminParams = {
        ($To -eq $ScriptAdmin) -and 
        ($Priority -eq 'High') -and 
        ($Subject -eq 'FAILURE')
    }

    $testParams = @{
        ScriptName              = 'Test (Brecht)'
        ImportDir               = New-Item "TestDrive:/TestMatrixFolder" -ItemType Directory
        LogFolder               = New-Item "TestDrive:/TestLogFolder" -ItemType Directory
        ScriptSetPermissionFile = New-Item "TestDrive:/TestSetPermissions.ps1" -ItemType File
        ScriptTestRequirements  = New-Item "TestDrive:/TestScriptTestRequirements.ps1" -ItemType File
        DefaultsFile            = New-Item "TestDrive:/Default.xlsx" -ItemType File
    }

    $testCherwellFolder = New-Item "TestDrive:/TestCherwellFolder" -ItemType Directory

    #region Valid Excel files
    $testMatrix = @(
        [PSCustomObject]@{Path = 'Path'; ACL = @{'Bob' = 'L' }; Parent = $true; Ignore = $false }
    )
    $testPermissions = @(
        [PSCustomObject]@{P1 = $null      ; P2 = 'bob' }
        [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
        [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
        [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
        [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
    )
    $testSettings = @(
        [PSCustomObject]@{Status = 'Enabled'; ComputerName = 'S1'; Path = 'E:\Department'; Action = 'Check' }
    )
    $testDefaultSettings = @(
        [PSCustomObject]@{ADObjectName = 'Bob' ; Permission = 'L'; MailTo = 'Bob@contoso.com' }
        [PSCustomObject]@{ADObjectName = 'Mike'; Permission = 'R' }
    )
    #endregion

    $testDefaultSettings | Export-Excel -Path $testParams.DefaultsFile -WorksheetName Settings

    $testComputerName = $env:COMPUTERNAME
    $testComputerName2 = 'DEUSFFRAN0031'

    if (-not (Test-Connection -ComputerName $testComputerName2 -Count 1 -Quiet)) {
        throw "Test computer '$testComputerName2' is not online"
    }

    $SettingsParams = @{
        Path          = Join-Path $testParams.ImportDir 'Matrix.xlsx'
        WorkSheetName = 'Settings'
    }
    $PermissionsParams = @{
        Path          = $SettingsParams.Path
        WorkSheetName = 'Permissions'
    }

    # Mock ConvertTo-MatrixAclHC
    Mock ConvertTo-MatrixAclHC { $true }
    Mock ConvertTo-MatrixADNamesHC { @{
            'P2' = @{
                SamAccountName = 'bob'
                Original       = @{
                    Begin  = ''
                    Middle = ''
                    End    = 'bob'
                }
                Converted      = @{
                    Begin  = ''
                    Middle = ''
                    End    = 'bob'
                }
            }
        } }
    Mock Invoke-Command
    Mock New-PSSession
    Mock Optimize-ExecutionOrderHC { $Name }
    Mock Send-MailHC
    Mock Test-MatrixPermissionsHC
    Mock Test-MatrixSettingHC
    Mock Wait-MaxRunningJobsHC
    Mock Write-EventLog
    Mock Test-FormDataHC
    Mock Get-AdUserPrincipalNameHC
}
Describe 'the mandatory parameters are' {
    It "<Name>" -TestCases @(
        @{Name = 'ScriptName' }
        @{Name = 'ImportDir' }
    ) {
        (Get-Command $testScript).Parameters[$Name].Attributes.Mandatory | Should -Be $true
    }
}
Describe 'stop the script and send an e-mail to the admin when' {
    Context 'a file or folder is not found' {
        It 'ScriptSetPermissionFile' {
            $testParams = $testParams.Clone()
            $testParams.ScriptSetPermissionFile = 'NonExisting.ps1'

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*NonExisting.ps1*not found*")
            }

            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter { $EntryType -eq 'Error' }
        }
        It 'ScriptTestRequirements' {
            $testParams = $testParams.Clone()
            $testParams.ScriptTestRequirements = 'ShareConfigNotExisting.ps1'

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*ShareConfigNotExisting.ps1*not found*")
            }

            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter { $EntryType -eq 'Error' }
        }
        It 'LogFolder' {
            $testParams = $testParams.Clone()
            $testParams.LogFolder = 'NonExistingLog'

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*NonExistingLog*not found*")
            }

            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter { $EntryType -eq 'Error' }
        }
        It 'CherwellFolder' {
            $testParams = $testParams.Clone()
            $testParams.CherwellFolder = 'NonExistingFolder'

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*NonExistingFolder*not found*")
            }

            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter { $EntryType -eq 'Error' }
        }
    }
    Context 'the default settings file' {
        It 'is not found' {
            $clonedParams = $testParams.Clone()
            $clonedParams.DefaultsFile = 'notExisting'

            .$testScript @clonedParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*$($clonedParams.DefaultsFile)*Cannot find*")
            }

            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter { $EntryType -eq 'Error' }
        }
        It "does not have the worksheet 'Settings'" {
            $clonedParams = $testParams.Clone()
            $clonedParams.DefaultsFile = New-Item "TestDrive:/Folder/Default.xlsx" -ItemType File -Force

            '1' | Export-Excel -Path $clonedParams.DefaultsFile -WorksheetName Sheet1

            .$testScript @clonedParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*'$($clonedParams.DefaultsFile)'* worksheet 'Settings' not found*")
            }

            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter { $EntryType -eq 'Error' }
        }

        $TestCases = @(
            @{
                Name         = "column header 'MailTo'"
                DefaultsFile = @(
                    [PSCustomObject]@{ADObjectName = 'Bob' ; Permission = 'L' }
                    [PSCustomObject]@{ADObjectName = 'Mike'; Permission = 'R' }
                )
                errorMessage = "Column header 'MailTo' not found"
            }
            @{
                Name         = "column header 'ADObjectName'"
                DefaultsFile = @(
                    [PSCustomObject]@{Permission = 'L'; MailTo = 'Bob@mail.com' }
                    [PSCustomObject]@{Permission = 'R' }
                )
                errorMessage = "Column header 'ADObjectName' not found"
            }
            @{
                Name         = "column header 'Permission'"
                DefaultsFile = @(
                    [PSCustomObject]@{ADObjectName = 'Bob' ; MailTo = 'Bob@mail.com' }
                    [PSCustomObject]@{ADObjectName = 'Mike' }
                )
                errorMessage = "Column header 'Permission' not found"
            }
            @{
                Name         = "'MailTo' addresses"
                DefaultsFile = @(
                    [PSCustomObject]@{ADObjectName = 'Bob' ; Permission = 'L'; MailTo = $null }
                    [PSCustomObject]@{ADObjectName = 'Mike'; Permission = 'R'; MailTo = ' ' }
                )
                errorMessage = "No mail addresses found"
            }
        )

        It "is missing <Name>" -TestCases $TestCases {
            $clonedParams = $testParams.Clone()
            $clonedParams.DefaultsFile = New-Item "TestDrive:/Folder/Default.xlsx" -ItemType File -Force

            $DefaultsFile | Export-Excel -Path $clonedParams.DefaultsFile -WorksheetName Settings

            .$testScript @clonedParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*$($clonedParams.DefaultsFile)*$errorMessage*")
            }

            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter { $EntryType -eq 'Error' }
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
        It '<Name> is missing' -TestCases @(
            @{ Name = 'CherwellAdObjectsFileName' }
            @{ Name = 'CherwellFormDataFileName' }
            @{ Name = 'CherwellExcelOverviewFileName' }
        ) {
            $clonedCherwellParams = $testCherwellParams.Clone()
            $clonedCherwellParams.$Name = ''

            .$testScript @testParams @clonedCherwellParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Parameter '$Name' is mandatory when the parameter CherwellFolder is used*")
            }

            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter { $EntryType -eq 'Error' }
        }
    }
}
Describe 'a sub folder in the log folder' {
    BeforeAll {
        @(
            [PSCustomObject]@{Status = $null; ComputerName = 'S1'; Path = 'E:\Test'; Action = 'Check' }
        ) | Export-Excel @SettingsParams
        @(
            [PSCustomObject]@{P1 = $null      ; P2 = 'Manager' }
            [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
            [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
            [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
            [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
        ) | Export-Excel @PermissionsParams

        .$testScript @testParams

        $testLogFolder = "$($testParams.LogFolder)\Permission matrix\$($testParams.ScriptName)"
    }
    It "is created for each specific Excel file regardless of its 'Status'" {
        @(Get-ChildItem -Path $testLogFolder -Directory).Count |
        Should -BeExactly 1
    }
    It 'the Excel file is copied to the log folder' {
        $testMatrixLogFolder = Get-ChildItem -Path $testLogFolder -Directory

        @(Get-ChildItem -Path $testMatrixLogFolder.FullName -File -Filter '*.xlsx').Count | Should -BeExactly 1
    }
}
Describe "when the 'Archive' switch is used then" {
    BeforeAll {
        @(
            [PSCustomObject]@{ComputerName = 'S1'; Path = 'E:\Department'; Action = 'Check' }
        ) | Export-Excel @SettingsParams
        @(
            [PSCustomObject]@{P1 = $null      ; P2 = 'Manager' }
            [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
            [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
            [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
            [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
        ) | Export-Excel @PermissionsParams

        .$testScript @testParams -Archive
    }
    It "a sub folder in the 'ImportDir' named 'Archive' is created" {
        "$($testParams.ImportDir)\Archive" | Should -Exist
    }
    It 'all matrix files are moved to the archive folder, even disabled ones' {
        $SettingsParams.Path | Should -Not -Exist
        "$($testParams.ImportDir)\Archive\Matrix.xlsx" | Should -Exist
    }
    It 'a matrix with the same name is overwritten in the archive folder' {
        $testFile = "$($testParams.ImportDir)\Archive\Matrix.xlsx"
        $testFile | Remove-Item -Force -EA Ignore

        @(
            [PSCustomObject]@{
                ComputerName = 'S1'; Path = 'E:\Department'; Action = 'Check' 
            }
        ) | Export-Excel -Path $testFile -WorksheetName $SettingsParams.WorkSheetName

        $testFile | Should -Exist

        @(
            [PSCustomObject]@{
                ComputerName = 'S2'; Path = 'E:\Department'; Action = 'Check' 
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
        Remove-Item -Path "$($testParams.ImportDir)\Archive" -Recurse -EA Ignore
        1..5 | ForEach-Object {
            $FileName = "$($testParams.ImportDir)\Matrix $_.xlsx"
            @(
                [PSCustomObject]@{Status = 'Enabled'; ComputerName = 'S1'; Path = 'E:\Department'; GroupName = 'G1'; SiteName = 'S1'; SiteCode = 'C1'; Action = 'Check' }
            ) | Export-Excel -Path $FileName -WorksheetName Settings
            @(
                [PSCustomObject]@{P1 = $null      ; P2 = 'Manager' }
                [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
                [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
                [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
                [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
            ) | Export-Excel -Path $FileName -WorksheetName Permissions
        }

        .$testScript @testParams -Archive

        (Get-ChildItem "$($testParams.ImportDir)\Matrix*" -File).Count | Should -BeExactly 0
        (Get-ChildItem "$($testParams.ImportDir)\Archive" -File).Count | 
        Should -BeExactly 5
    }
}
Describe "do not invoke the script to set permissions when" {
    It "there's only a default settings file in the 'ImportDir' folder" {
        .$testScript @testParams

        Should -Not -Invoke Invoke-Command
    }
    It "there are only other file types than '.xlsx' in the 'ImportDir' folder" {
        1 | Out-File "$($testParams.ImportDir)\Wrong.txt"
        1 | Out-File "$($testParams.ImportDir)\Wrong.csv"

        .$testScript @testParams

        Should -Not -Invoke Invoke-Command
    }
    It "there are only valid matrixes in subfolders of the 'ImportDir' folder" {
        $Folder = (New-Item  "$($testParams.ImportDir)\Archive" -ItemType Directory -Force -EA Ignore).FullName
        @(
            [PSCustomObject]@{Status = 'Enabled'; ComputerName = 'S1'; Path = 'E:\Department'; Action = 'Check' }
        ) | Export-Excel -Path "$Folder/Matrix.xlsx" -WorksheetName Settings
        @(
            [PSCustomObject]@{P1 = $null      ; P2 = 'Manager' }
            [PSCustomObject]@{P1 = 'SiteCode' ; P2 = 'SiteCode' }
            [PSCustomObject]@{P1 = 'GroupName'; P2 = 'GroupName' }
            [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
            [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
        ) | Export-Excel -Path "$Folder/Matrix.xlsx" -WorksheetName Permissions

        .$testScript @testParams

        Should -Not -Invoke Invoke-Command
    }
    It "the 'Status' in the worksheet 'Settings' of the matrix file is not set to 'Enabled'" {
        @(
            [PSCustomObject]@{Status = 'NOTEnabled'; ComputerName = 'S1'; Path = 'E:\Department'; Action = 'Check' }
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
        Remove-Item -Path "$($testParams.ImportDir)\*" -Exclude $TestDefaultsFileName -Recurse -Force -EA Ignore
    }
    Context "for the Excel 'File' when" {
        It "building the matrix with 'ConvertTo-MatrixAclHC' fails" {
            Mock ConvertTo-MatrixAclHC {
                throw 'Failed building the matrix'
            }

            @(
                [PSCustomObject]@{Status = 'Enabled'; ComputerName = 'S1'; Path = 'E:\Department'; Action = 'Check' }
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
                Value       = "Worksheet 'Settings' not found"
            }.GetEnumerator().ForEach( {
                    $ImportedMatrix.File.Check.($_.Key) | Should -Be $_.Value
                })
        }
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
                Value       = "Worksheet 'Settings' is empty"
            }.GetEnumerator().ForEach( {
                    $ImportedMatrix.File.Check.($_.Key) | Should -Be $_.Value
                })
        }
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
        }
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
                    Value       = @{('S1.' + $env:USERDNSDOMAIN) = 'E:\DUPLICATE' }
                }
            )

            @(
                [PSCustomObject]@{Status = 'Enabled'; ComputerName = $($testProblem.Value.Keys); Path = 'E:\Reports'; Action = 'Check' }
                [PSCustomObject]@{Status = 'Enabled'; ComputerName = $($testProblem.Value.Keys); Path = $($testProblem.Value.Values); Action = 'Check' }
                [PSCustomObject]@{Status = 'Enabled'; ComputerName = $($testProblem.Value.Keys); Path = $($testProblem.Value.Values); Action = 'Fix' }
                [PSCustomObject]@{Status = 'Enabled'; ComputerName = 'S3'; Path = 'E:\Department'; Action = 'Check' }
            ) | Export-Excel @SettingsParams
            $testPermissions | Export-Excel @PermissionsParams

            .$testScript @testParams

            $toTest = @($ImportedMatrix.Settings.Where( { $_.Import.Path -eq $testProblem.Value.Values }))
            $toTest.Count | Should -BeExactly 2

            $toTest.ForEach( {
                    # Assert-Equivalent $_.Check $testProblem
                    $_.Check.Type | Should -Be  $testProblem.Type
                    $_.Check.Name | Should -Be  $testProblem.Name
                    $_.Check.Description | Should -Be  $testProblem.Description
                    $_.Check.Value.Name | Should -Be  $testProblem.Value.Name
                    $_.Check.Value.Value | Should -Be  $testProblem.Value.Value
                })
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
Describe 'on a successful run' {
    BeforeAll {
        #region Set permissions script
        Mock Invoke-Command {
            & $TestInvokeCommand -Scriptblock { 1 } -ComputerName localhost -AsJob -JobName 'SetPermissions_1'
        } -ParameterFilter {
            ($AsJob -eq $true) -and
            ($ComputerName -eq $testComputerName) -and
            ($JobName -eq 'SetPermissions_1')
        }
        Mock Invoke-Command {
            & $TestInvokeCommand -Scriptblock { 2 } -ComputerName localhost -AsJob -JobName 'SetPermissions_2'
        } -ParameterFilter {
            ($AsJob -eq $true) -and
            ($ComputerName -eq $testComputerName) -and
            ($JobName -eq 'SetPermissions_2')
        }
        Mock Invoke-Command {
            & $TestInvokeCommand -Scriptblock { 3 } -ComputerName localhost -AsJob -JobName 'SetPermissions_3'
        } -ParameterFilter {
            ($AsJob -eq $true) -and
            ($ComputerName -eq $testComputerName2) -and
            ($JobName -eq 'SetPermissions_3')
        }
        #endregion

        #region Test requirements script
        Mock Invoke-Command {
            & $TestInvokeCommand -Scriptblock { 'A' } -ComputerName $testComputerName -AsJob -JobName 'TestRequirements'
        } -ParameterFilter {
            ($AsJob -eq $true) -and
            ($ComputerName -eq $testComputerName) -and
            ($JobName -eq 'TestRequirements') 
        }
        Mock Invoke-Command {
            & $TestInvokeCommand -Scriptblock { 'B' } -ComputerName $testComputerName2 -AsJob -JobName 'TestRequirements'
        } -ParameterFilter {
            ($AsJob -eq $true) -and
            ($ComputerName -eq $testComputerName2) -and
            ($JobName -eq 'TestRequirements')
        }
        #endregion

        Mock Get-ADObjectDetailHC {
            [PSCustomObject]@{
                samAccountName = 'A B bob'
                adObject       = @{ 
                    ObjectClass    = 'user'
                    Name           = 'A B Bob'
                    SamAccountName = 'A B bob'
                }
                adGroupMember  = $null
            }
            [PSCustomObject]@{
                samAccountName = 'C D bob'
                adObject       = @{ 
                    ObjectClass    = 'user'
                    Name           = 'C D Bob'
                    SamAccountName = 'C D bob'
                }
                adGroupMember  = $null
            }
            [PSCustomObject]@{
                samAccountName = 'E F bob'
                adObject       = @{ 
                    ObjectClass    = 'user'
                    Name           = 'E F Bob'
                    SamAccountName = 'E F bob'
                }
                adGroupMember  = $null
            }
        }

        Mock ConvertTo-MatrixAclHC { 
            @(
                [PSCustomObject]@{
                    Path   = 'E:\Department'
                    Parent = $true
                    ACL    = @{'A B bob' = 'L' }
                    Ignore = $false
                }
                [PSCustomObject]@{
                    Path   = 'folder'
                    Parent = $false
                    ACL    = @{'A B bob' = 'W' }
                    Ignore = $false
                }
            )
        } -ParameterFilter { $adObjects.Values.SamAccountName -eq 'A B bob' }
        Mock ConvertTo-MatrixAclHC { 
            @(
                [PSCustomObject]@{
                    Path   = 'E:\Reports'
                    Parent = $true
                    ACL    = @{'C D bob' = 'L' }
                    Ignore = $false
                }
                [PSCustomObject]@{
                    Path   = 'folder'
                    Parent = $false
                    ACL    = @{'C D bob' = 'W' }
                    Ignore = $false
                }
            )
        } -ParameterFilter { $adObjects.Values.SamAccountName -eq 'C D bob' }
        Mock ConvertTo-MatrixAclHC { 
            @(
                [PSCustomObject]@{
                    Path   = 'E:\Finance'
                    Parent = $true
                    ACL    = @{'E F bob' = 'L' }
                    Ignore = $false
                }
                [PSCustomObject]@{
                    Path   = 'folder'
                    Parent = $false
                    ACL    = @{'E F bob' = 'W' }
                    Ignore = $false
                }
            )
        } -ParameterFilter { $adObjects.Values.SamAccountName -eq 'E F bob' }

        Mock ConvertTo-MatrixADNamesHC {  
            @{
                'P2' = @{
                    SamAccountName = 'A B bob'
                    Original       = @{
                        Begin  = 'GroupName'
                        Middle = 'SiteCode'
                        End    = 'bob'
                    }
                    Converted      = @{
                        Begin  = 'A'
                        Middle = 'B'
                        End    = 'bob'
                    }
                }
            } 
        } -ParameterFilter { $Begin -eq 'A' }
        Mock ConvertTo-MatrixADNamesHC {  
            @{
                'P2' = @{
                    SamAccountName = 'C D bob'
                    Original       = @{
                        Begin  = 'GroupName'
                        Middle = 'SiteCode'
                        End    = 'bob'
                    }
                    Converted      = @{
                        Begin  = 'C'
                        Middle = 'D'
                        End    = 'bob'
                    }
                }
            } 
        } -ParameterFilter { $Begin -eq 'C' }
        Mock ConvertTo-MatrixADNamesHC {  
            @{
                'P2' = @{
                    SamAccountName = 'E F bob'
                    Original       = @{
                        Begin  = 'GroupName'
                        Middle = 'SiteCode'
                        End    = 'bob'
                    }
                    Converted      = @{
                        Begin  = 'E'
                        Middle = 'F'
                        End    = 'bob'
                    }
                }
            } 
        } -ParameterFilter { $Begin -eq 'E' }

        @(
            [PSCustomObject]@{
                Status = 'Enabled'; ComputerName = $testComputerName; 
                Path = 'E:\Department'; Action = 'Check' 
                GroupName = 'A'; SiteCode = 'B'
            }
            [PSCustomObject]@{
                Status = 'Enabled'; ComputerName = $testComputerName; 
                Path = 'E:\Reports'; Action = 'Check' 
                GroupName = 'C'; SiteCode = 'D'
            }
            [PSCustomObject]@{
                Status = 'Enabled'; ComputerName = $testComputerName2; 
                Path = 'E:\Finance'; Action = 'Check' 
                GroupName = 'E'; SiteCode = 'F'
            }
        ) | Export-Excel @SettingsParams

        $testPermissions | Export-Excel @PermissionsParams

        .$testScript @testParams
    }
    Context 'the worksheet Permission is' {
        It 'tested for incorrect input' {
            Should -Invoke Test-MatrixPermissionsHC -Exactly 1 -Scope Describe
        }
    }
    Context 'each row in the worksheet Settings' {
        It 'is tested for incorrect input' {
            Should -Invoke Test-MatrixSettingHC -Exactly 3 -Scope Describe
        }
        It 'creates AD object names' {
            Should -Invoke ConvertTo-MatrixADNamesHC -Exactly 3 -Scope Describe
            Should -Invoke ConvertTo-MatrixADNamesHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($Begin -eq 'A') -and ($Middle -eq 'B') -and 
                ($ColumnHeaders -ne $null)
            }
            Should -Invoke ConvertTo-MatrixADNamesHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($Begin -eq 'C') -and ($Middle -eq 'D') -and
                ($ColumnHeaders -ne $null)
            }
            Should -Invoke ConvertTo-MatrixADNamesHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($Begin -eq 'E') -and ($Middle -eq 'F') -and
                ($ColumnHeaders -ne $null)
            }
        }
        It 'creates a matrix with path and Acl' {
            Should -Invoke ConvertTo-MatrixAclHC -Exactly 3 -Scope Describe
        }
    }
    Context 'the test requirements script' {
        It 'is called once for each unique ComputerName' {
            Should -Invoke Invoke-Command -Times 2 -Exactly -Scope Describe -ParameterFilter {
                $JobName -eq 'TestRequirements' 
            }
        }
        It 'saves the job result in Settings for each matrix' {
            @($ImportedMatrix.Settings.Where( {
                        ($_.Import.ComputerName -eq $testComputerName) -and 
                        ($_.Check -eq 'A') })).Count |
            Should -BeExactly 2
    
            @($ImportedMatrix.Settings.Where( {
                        ($_.Import.ComputerName -eq $testComputerName2) -and 
                        ($_.Check -eq 'B') })).Count |
            Should -BeExactly 1
        }
    }
    Context 'the set permissions script' {
        It 'is called for each Enabled row in Settings' {
            Should -Invoke Invoke-Command -Times 3 -Exactly -Scope Describe -ParameterFilter {
                $JobName -like 'SetPermissions*' 
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
    It 'an email is sent to the user in the default settings file' {
        Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($To -eq 'Bob@contoso.com') -and
            ($Subject -eq '1 matrix file') -and
            ($Save -like "$($testParams.LogFolder.FullName)* - Mail - 1 matrix file.html") -and
            ($Priority -eq 'Normal') -and
            ($Message -notLike '*Cherwell*') -and
            ($Message -like '*Matrix results per file*') -and
            ($Message -like '*Matrix.xlsx*') -and
            ($Message -like '*Settings*') -and
            ($Message -like '*ID*ComputerName*Path*Action*Duration*') -and
            ($Message -like "*1*$testComputerName*E:\Department*Check*") -and
            ($Message -like "*2*$testComputerName*E:\Reports*Check*") -and
            ($Message -like "*3*$testComputerName2*E:\Finance*Check*") -and
            ($Message -like '*Error*Warning*Information*')
        }
    }
    It "the worksheet 'adObjects' is added to the matrix log file" {
        $testFiles = Get-ChildItem $testParams.logFolder -Filter '*.xlsx' -Recurse -File
        Import-Excel -Path $testFiles.FullName -WorksheetName 'adObjects' |
        Should -Not -BeNullOrEmpty
    } -Tag test
} 
Describe 'when a job fails' {
    Context 'the test requirements script' {
        BeforeAll {
            Mock Invoke-Command {
                & $TestInvokeCommand -Scriptblock { throw 'failure' } -ComputerName $testComputerName -AsJob -JobName 'TestRequirements'
            } -ParameterFilter {
                ($AsJob -eq $true) -and
                ($ComputerName -eq $testComputerName) -and
                ($JobName -eq 'TestRequirements') 
            }
            Mock Invoke-Command {
                & $TestInvokeCommand -Scriptblock { 'B' } -ComputerName $testComputerName2 -AsJob -JobName 'TestRequirements'
            } -ParameterFilter {
                ($AsJob -eq $true) -and
                ($ComputerName -eq $testComputerName2) -and
                ($JobName -eq 'TestRequirements')
            }

            @(
                [PSCustomObject]@{Status = 'Enabled'; ComputerName = $testComputerName; Path = 'E:\Department'; Action = 'Check' }
                [PSCustomObject]@{Status = 'Enabled'; ComputerName = $testComputerName2; Path = 'E:\Reports'; Action = 'Check' }
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
            Mock Invoke-Command {
                & $TestInvokeCommand -Scriptblock { 1 } -ComputerName $testComputerName -AsJob -JobName 'SetPermissions_1'
            } -ParameterFilter {
                ($AsJob -eq $true) -and
                ($JobName -eq 'SetPermissions_1')
            }
            Mock Invoke-Command {
                & $TestInvokeCommand -Scriptblock { throw 'failure' } -ComputerName $testComputerName -AsJob -JobName 'SetPermissions_2'
            } -ParameterFilter {
                ($AsJob -eq $true) -and
                ($JobName -eq 'SetPermissions_2')
            }

            @(
                [PSCustomObject]@{Status = 'Enabled'; ComputerName = $testComputerName; Path = 'E:\Department'; Action = 'Check' }
                [PSCustomObject]@{Status = 'Enabled'; ComputerName = $testComputerName; Path = 'E:\Reports'; Action = 'Check' }
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
                    [PSCustomObject]@{Path = 'Path'; ACL = @{'Mike' = 'L' }; Parent = $true; Ignore = $false }
                )
            }
            Mock Get-DefaultAclHC {
                @{
                    'Bob' = 'R'
                }
            }
            @(
                [PSCustomObject]@{Status = 'Enabled'; ComputerName = 'test'; Path = 'E:\Department'; Action = 'Check' }
            ) | Export-Excel @SettingsParams
            @(
                [PSCustomObject]@{P1 = $null      ; P2 = 'Mike' }
                [PSCustomObject]@{P1 = 'SiteCode' ; P2 = '' }
                [PSCustomObject]@{P1 = 'GroupName'; P2 = '' }
                [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
                [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
            ) | Export-Excel @PermissionsParams

            .$testScript @testParams

            $Actual = ($ImportedMatrix.Settings.Matrix.Where( { $_.Path -eq 'Path' })).ACL

            $Expected = @{
                'Bob'  = 'R'
                'Mike' = 'L'
            }
            Assert-Equivalent -Actual $Actual -Expected $Expected
        }
        It 'do not add default permissions to the matrix ACL when the folder has no ACL' {
            Mock Test-ExpandedMatrixHC
            Mock ConvertTo-MatrixAclHC {
                @(
                    [PSCustomObject]@{
                        Path   = 'Path'; 
                        ACL    = @{}; 
                        Parent = $true; 
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
                [PSCustomObject]@{Status = 'Enabled'; ComputerName = 'test'; Path = 'E:\Department'; Action = 'Check' }
            ) | Export-Excel @SettingsParams
            @(
                [PSCustomObject]@{P1 = $null      ; P2 = 'Mike' }
                [PSCustomObject]@{P1 = 'SiteCode' ; P2 = '' }
                [PSCustomObject]@{P1 = 'GroupName'; P2 = '' }
                [PSCustomObject]@{P1 = 'Path'     ; P2 = '' }
                [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'L' }
            ) | Export-Excel @PermissionsParams

            .$testScript @testParams

            $Actual = ($ImportedMatrix.Settings.Matrix.Where( { 
                        $_.Path -eq 'Path' })).ACL

            Assert-Equivalent -Actual $Actual -Expected @{}
        } 
        It 'do not overwrite permissions to the matrix ACL when they are also in the default ACL' {
            Mock Test-ExpandedMatrixHC
            Mock ConvertTo-MatrixAclHC {
                @(
                    [PSCustomObject]@{Path = 'Path'; ACL = @{'Mike' = 'L'; 'Bob' = 'L' }; Parent = $true; Ignore = $false }
                )
            }
            Mock Get-DefaultAclHC {
                @{
                    'Bob' = 'R'
                }
            }
            @(
                [PSCustomObject]@{Status = 'Enabled'; ComputerName = 'test'; Path = 'E:\Department'; Action = 'Check' }
            ) | Export-Excel @SettingsParams
            @(
                [PSCustomObject]@{P1 = $null      ; P2 = 'Mike' }
                [PSCustomObject]@{P1 = 'SiteCode' ; P2 = '' }
                [PSCustomObject]@{P1 = 'GroupName'; P2 = '' }
                [PSCustomObject]@{P1 = 'Path'     ; P2 = 'L' }
                [PSCustomObject]@{P1 = 'Folder'   ; P2 = 'W' }
            ) | Export-Excel @PermissionsParams

            .$testScript @testParams

            $Actual = ($ImportedMatrix.Settings.Matrix.Where( { $_.Path -eq 'Path' })).ACL

            $Expected = @{
                'Bob'  = 'L'
                'Mike' = 'L'
            }
            Assert-Equivalent -Actual $Actual -Expected $Expected
        }
    }
}
Describe 'when a FatalError occurs while executing the matrix' {
    AfterEach {
        $Error.Clear()
        Remove-Item -Path "$($testParams.LogFolder)\*" -Recurse -Force -EA Ignore
        Remove-Item -Path "$($testParams.ImportDir)\*" -Exclude $TestDefaultsFileName -Recurse -Force -EA Ignore
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
            [PSCustomObject]@{Status = 'Enabled'; ComputerName = 'S1'; Path = 'E:\Department'; Action = 'Check' }
            [PSCustomObject]@{Status = 'Enabled'; ComputerName = 'S2'; Path = 'E:\Department'; Action = 'Check' }
        ) | Export-Excel @SettingsParams
        $testPermissions | Export-Excel @PermissionsParams

        .$testScript @testParams

        $testMatrixLogFolder = Get-ChildItem -Path "$($testParams.LogFolder)\Permission matrix\$($testParams.ScriptName)" -Directory
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
            [PSCustomObject]@{Status = 'Enabled'; ComputerName = 'S1'; Path = 'E:\Department'; Action = 'Check'; GroupName = 'Group'; SiteCode = 'Site' }
            [PSCustomObject]@{Status = 'Enabled'; ComputerName = 'S2'; Path = 'E:\Department'; Action = 'Check' }
        ) | Export-Excel @SettingsParams
        $testPermissions | Export-Excel @PermissionsParams

        .$testScript @testParams

        $testMatrixLogFolder = Get-ChildItem -Path "$($testParams.LogFolder)\Permission matrix\$($testParams.ScriptName)" -Directory
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
            [PSCustomObject]@{Status = 'Enabled'; ComputerName = 'S1'; Path = 'E:\Department'; Action = 'Check'; GroupName = 'Group'; SiteCode = 'Site' }
            [PSCustomObject]@{Status = 'Enabled'; ComputerName = 'S2'; Path = 'E:\Department'; Action = 'Check' }
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
                [PSCustomObject]@{Status = 'Enabled'; ComputerName = $testComputerName; Path = 'E:\Department'; Action = 'Check' }
            ) | Export-Excel @SettingsParams

            $testPermissions | Export-Excel @PermissionsParams

            .$testScript @testParams -CherwellFolder $testCherwellFolder.FullName
        }
        It 'a FatalError is registered for the file' {
            $actual = $ImportedMatrix.File.Check
            $actual.Type | Should -Contain "FatalError"
            $actual.Name | Should -Contain "Worksheet 'FormData' not found"
        }
        It 'the permissions script is not executed' {
            Should -Not -Invoke Invoke-Command
        }
        It 'an email is sent to the user with the error' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
                ($To -eq 'Bob@contoso.com') -and
                ($Save -like "$($testParams.LogFolder.FullName)* - Mail - 1 matrix file, 1 error.html") -and
                ($Subject -eq '1 matrix file, 1 error') -and
                ($Priority -eq 'High') -and
                ($Message -like "*Worksheet 'FormData' not found*") -and
                ($Message -notLike '*Check the*overview*for details*')
            }
        }
    }
    Context 'but the worksheet FormData contains incorrect data' {
        AfterAll {
            Mock Test-FormDataHC
            Mock Get-AdUserPrincipalNameHC
        }
        BeforeAll {
            Mock Test-FormDataHC { 
                @{
                    Type = 'FatalError'
                    Name = 'incorrect data'
                } 
            }

            @(
                [PSCustomObject]@{Status = 'Enabled'; ComputerName = $testComputerName; Path = 'E:\Department'; Action = 'Check' }
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
                ($Save -like "$($testParams.LogFolder.FullName)* - Mail - 1 matrix file, 1 error.html") -and
                ($Subject -eq '1 matrix file, 1 error') -and
                ($Priority -eq 'High') -and
                ($Message -like "*Errors*Warnings*FormData*") -and
                ($Message -like "*FormData*incorrect data*") -and
                ($Message -notLike '*Check the*overview*for details*')
            }
        }
    }
    Context 'but the worksheet FormData has a non existing MatrixResponsible' {
        BeforeAll {
            Mock Test-FormDataHC
            Mock Get-AdUserPrincipalNameHC { 
                @{
                    UserPrincipalName = 'mike@contoso.com'
                    notFound          = 'bob@contoso.com'
                }
            }

            @(
                [PSCustomObject]@{Status = 'Enabled'; ComputerName = $testComputerName; Path = 'E:\Department'; Action = 'Check' }
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
                ($Save -like "$($testParams.LogFolder.FullName)* - Mail - 1 matrix file, 1 warning.html") -and
                ($Subject -eq '1 matrix file, 1 warning') -and
                ($Priority -eq 'High') -and
                ($Message -like "*Errors*Warnings*FormData*") -and
                ($Message -like "*FormData*AD object not found*") -and
                ($Message -like '*Check the*overview*for details*')
            }
        }
    }
}
Describe 'when the argument CherwellFolder is used on a successful run' {
    BeforeAll {
        Mock Get-AdUserPrincipalNameHC { 
            @{
                UserPrincipalName = @('bob@contoso.com', 'mike@contoso.com')
                notFound          = $null
            }
        }
        Mock Test-FormDataHC

        @(
            [PSCustomObject]@{Status = 'Enabled'; ComputerName = 'SERVER1'; Path = 'E:\Department'; Action = 'Check' }
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

        Mock ConvertTo-MatrixADNamesHC {
            @{
                'P2' = @{
                    SamAccountName = 'BEGIN Manager'
                    Original       = @{
                        Begin  = 'A'
                        Middle = ''
                        End    = 'Manager'
                    }
                    Converted      = @{
                        Begin  = 'BEGIN'
                        Middle = ''
                        End    = 'Manager'
                    }
                } 
            }
        }

        $testPermissions | Export-Excel @PermissionsParams

        $testCherwellParams = @{
            CherwellFolder                = $testCherwellFolder.FullName
            CherwellAdObjectsFileName     = 'AD object names.csv'
            CherwellFormDataFileName      = 'Form data.csv'
            CherwellExcelOverviewFileName = 'Overview.xlsx'
        }
        .$testScript @testParams @testCherwellParams

        $testCherwellExport = Get-ChildItem $testCherwellFolder.FullName

        $testLogFolderExport = Get-ChildItem $testParams.LogFolder.FullName  -Recurse -File |
        Where-Object Extension -Match '.xlsx$|.csv$'

        $testLogFolder = @{
            ExcelFile        = $testLogFolderExport | 
            Where-Object Name -Like '*Overview.xlsx'
            FormDataCsvFile  = $testLogFolderExport | 
            Where-Object Name -Like '*Form data.csv'
            AdObjectsCsvFile = $testLogFolderExport | 
            Where-Object Name -Like '*AD object names.csv'
        }

        $testCherwellFolder = @{
            ExcelFile        = $testCherwellExport | 
            Where-Object Name -Like '*Overview.xlsx'
            FormDataCsvFile  = $testCherwellExport | 
            Where-Object Name -EQ 'Form data.csv'
            AdObjectsCsvFile = $testCherwellExport | 
            Where-Object Name -EQ 'AD object names.csv'
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
            It '<Name>' -TestCases @(
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
            It '<Name>' -TestCases @(
                @{ Name = 'MatrixFileName'; Value = 'Matrix' }
                @{ Name = 'SamAccountName'; Value = 'BEGIN Manager' }
                @{ Name = 'GroupName'; Value = 'BEGIN' }
                @{ Name = 'SiteCode'; Value = '' }
                @{ Name = 'Name'; Value = 'Manager' }
            ) {
                $actual.cherwellFolder.AdObjectNames.$Name | Should -Be $Value
                $actual.cherwellFolder.Excel.$Name | Should -Be $Value
                $actual.logFolder.AdObjectNames.$Name | Should -Be $Value
                $actual.logFolder.Excel.$Name | Should -Be $Value
            }
        }
    }
    It 'an email is sent to the user in the default settings file' {
        Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($To -eq 'Bob@contoso.com') -and
            ($Save -like "$($testParams.LogFolder.FullName)* - Mail - 1 matrix file.html") -and
            ($Subject -eq '1 matrix file') -and
            ($Priority -eq 'Normal') -and
            ($Message -like '*Export to*Cherwell*') -and
            ($Message -like '*Check the*overview*for details*') -and
            ($Message -like '*AD objects*2*') -and
            ($Message -like '*Form data*1*') -and
            ($Message -like '*Matrix results per file*') -and
            ($Message -like '*Matrix.xlsx*') -and
            ($Message -like '*Settings*') -and
            ($Message -like '*ID*ComputerName*Path*Action*Duration*') -and
            ($Message -like "*1*SERVER1*E:\Department*Check*") -and
            ($Message -like '*Error*Warning*Information*')
        }
    }
}