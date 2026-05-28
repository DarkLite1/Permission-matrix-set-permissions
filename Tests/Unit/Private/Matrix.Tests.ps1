#Requires -Version 7
#Requires -Modules Pester

Describe 'Matrix Logic Tests' {
    BeforeDiscovery {
        . "$PSScriptRoot/../../Helpers/Fixtures.Matrix.ps1"

        $script:MatrixSettingsFixtures = Get-MatrixSettingsFixtures
        $script:MatrixPermissionsFixtures = Get-MatrixPermissionsFixtures
        $script:DisabledMatrixFixtures = Get-DisabledMatrixFixtures
        $script:DuplicatePathFixtures = Get-DuplicateMatrixFixtures
        $script:DefaultMergeFixtures = Get-DefaultPermissionsMergeFixtures
        $script:ADBuildFixtures = Get-AdObjectBuildFixtures
        $script:MatrixBuildFixtures = Get-MatrixBuildFixtures

        $script:testScript = "$PSScriptRoot\..\..\..\Scripts\Entrypoints\PermissionMatrix.ps1"
    }

    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        . "$moduleRoot\Private\Utils.ps1"
        . "$moduleRoot\Private\Matrix.ps1"

        . "$root/Tests/Helpers/Helpers.HC.ps1"
        . "$root/Tests/Helpers/Fixtures.Excel.ps1"
        . "$root/Tests/Helpers/Fixtures.Matrix.ps1"
        . "$root/Tests/Helpers/Fixtures.Json.ps1"

        if (-not (Test-Path $testScript)) {
            throw "Script '$testScript' not found"
        }

        $jsonFile = New-Item 'TestDrive:\Input.json' -ItemType File

        $testInputTemplate = New-JsonFixture

        $testInputTemplate.Matrix.FolderPath =
        (New-Item 'TestDrive:\Matrix' -ItemType Directory).FullName

        $testInputTemplate.Matrix.DefaultsFile =
        (New-ValidDefaultsExcelFixture -Path 'TestDrive:\Defaults.xlsx')

        $testInputTemplate.Settings.SaveLogFiles.Where.Folder =
        (New-Item 'TestDrive:\MatrixLogs' -ItemType Directory).FullName
    
        $testParams = @{
            ConfigurationJsonFile = $jsonFile.FullName
        }

        $testInputTemplate | 
        ConvertTo-Json -Depth 20 | Set-Content $jsonFile.FullName

        $script:TestJsonFile = $jsonFile
        $script:TestInput = $testInputTemplate
        $script:TestScript = $testScript
        $script:TestParams = $testParams

        Import-Module "$moduleRoot\PermissionMatrix.psd1" -Force

        Mock Import-Module {
            Write-Verbose "Pester: Intercepted and skipped Import-Module for '$Name'"
        } -ParameterFilter { $Name -match 'PermissionMatrix' }

        Mock Send-MailKitMessageHC {
            Write-Verbose "Pester: Intercepted email to $($To -join ', ')"
        } -ModuleName 'PermissionMatrix'
        
        Mock Write-EventLog
    }

    BeforeEach {
        Clear-TestLogFoldersHC `
            -ConfiguredLogFolder $TestInput.Settings.SaveLogFiles.Where.Folder
    }

    Describe 'Matrix: Settings sheet validation' {
        It '<Issue> should be detected' -TestCases $MatrixSettingsFixtures {
            param($Issue, $SheetMutation, $ExpectedMessage)

            Invoke-Expression $SheetMutation | Out-Null

            $updated = Copy-ObjectHC $TestInput
            Save-TestJson $updated $TestJsonFile

            & $TestScript @TestParams

            Assert-HtmlLogContainsPatternHC `
                -LogFolderPath $TestInput.Settings.SaveLogFiles.Where.Folder `
                -Pattern "*$ExpectedMessage*"
        }
    }

    Describe 'Matrix: Permissions sheet validation' {
        It '<Issue> should be detected' -TestCases $MatrixPermissionsFixtures {
            param($Issue, $Mutation, $Expected)

            Invoke-Expression $Mutation | Out-Null

            $updated = Copy-ObjectHC $TestInput
            Save-TestJson $updated $TestJsonFile

            & $TestScript @TestParams

            Assert-HtmlLogContainsPatternHC `
                -LogFolderPath $TestInput.Settings.SaveLogFiles.Where.Folder `
                -Pattern "*$Expected*"
        }
    }

    Describe 'Matrix: Disabled matrices' {

        It '<Description>' -TestCases $DisabledMatrixFixtures {
            param($Description, $FixtureBuilder, $Expected, $NotExpected)

            & $FixtureBuilder  # creates .xlsx files in TestDrive:\Matrix

            $updated = Copy-ObjectHC $TestInput
            Save-TestJson $updated $TestJsonFile

            & $TestScript @TestParams

            foreach ($assert in $Assertions) {
                $assertParams = @{
                    LogFolderPath = $TestInput.Settings.SaveLogFiles.Where.Folder
                    Pattern       = $assert.Pattern
                }
                
                if ($assert.FileMatch) { 
                    $assertParams.FileMatch = $assert.FileMatch
                }
                if ($assert.Not) { 
                    $assertParams.Not = $true 
                }

                Assert-HtmlLogContainsPatternHC @assertParams
            }
        }
    }

    Describe 'Matrix: Duplicate combinations' {

        It 'detects duplicates correctly' -TestCases $DuplicatePathFixtures {
            param($FixtureBuilder, $ExpectedError)

            & $FixtureBuilder

            $updated = Copy-ObjectHC $TestInput
            Save-TestJson $updated $TestJsonFile

            & $TestScript @TestParams

            Assert-HtmlLogContainsPatternHC `
                -LogFolderPath $TestInput.Settings.SaveLogFiles.Where.Folder `
                -Pattern "*$ExpectedError*"
        }
    }

    Describe 'Default permissions merging' {

        It '<Description>' -TestCases $DefaultMergeFixtures {
            param($Description, $ApplyDefaultPermissions, $DefaultsRows, $MatrixRows, $ExpectedMerged, $ExpectedError)

            if ($ExpectedError) {
                { 
                    Merge-DefaultPermissionsHC `
                        -Defaults $DefaultsRows `
                        -Matrix $MatrixRows `
                        -ApplyDefaultPermissions $ApplyDefaultPermissions 
                } | Should -Throw -ExpectedMessage $ExpectedError
            }
            else {
                $output = Merge-DefaultPermissionsHC `
                    -Defaults $DefaultsRows `
                    -Matrix $MatrixRows `
                    -ApplyDefaultPermissions $ApplyDefaultPermissions
                
                $output.Count | Should -Be $ExpectedMerged.Count
                
                foreach ($expected in $ExpectedMerged) {
                    $match = $output | Where-Object { 
                        $_.ADObject -eq $expected.ADObject 
                    }
                    $match | Should -Not -BeNullOrEmpty
                    $match.Permission | Should -Be $expected.Permission
                }
            }
        }
    }

    Describe 'AD Object build logic' {

        It '<Description>' -TestCases $ADBuildFixtures {
            param($Description, $PermissionsRows, $SettingRow, $ExpectedMap)

            # Call the actual mapping function
            $actualMap = Get-MatrixADObjectsMapHC `
                -PermissionsSheet $PermissionsRows `
                -SettingRow $SettingRow

            $actualMap.Count | Should -Be $ExpectedMap.Count

            foreach ($key in $ExpectedMap.Keys) {
                $actualMap.Keys | Should -Contain $key -Because "Column '$key' should be mapped"
                $actualMap[$key] | Should -Be $ExpectedMap[$key] -Because 'The AD Object name should be correctly constructed'
            }
        }
    }

    Describe 'Matrix building logic' {
        BeforeEach {
            if (Test-Path 'TestDrive:\Matrix') {
                Remove-Item 'TestDrive:\Matrix\*' -Recurse -Force -EA Ignore
            }
        }

        It '<Description>' -TestCases $MatrixBuildFixtures {
            param($Description, $FixtureBuilder, $ExpectedFiles)

            & $FixtureBuilder

            $updated = Copy-ObjectHC $TestInput
            Save-TestJson $updated $TestJsonFile

            & $TestScript @TestParams

            $logFolder = Get-LatestLogFolderHC -Root $TestInput.Settings.SaveLogFiles.Where.Folder
            $htmlFiles = Get-ChildItem -Path $logFolder -Recurse -Filter '*.html'

            $htmlFiles.Count | Should -Be $ExpectedFiles
        }
    }

    Describe 'Get-DefaultAclHC validation' {
        BeforeEach {
            $script:errors = [System.Collections.Generic.List[object]]::new()
        }

        It 'accepts a complete row' {
            $sheet = @([PSCustomObject]@{ ADObjectName = 'IT Demand Management'; Permission = 'L'; MailTo = 'bob@contoso.com' })
            $result = Get-DefaultAclHC -Sheet $sheet -SystemErrors ([ref]$errors)

            $errors.Count | Should -Be 0
            $result['IT Demand Management'] | Should -Be 'L'
        }

        It 'silently skips MailTo-only rows' {
            $sheet = @([PSCustomObject]@{ ADObjectName = $null; Permission = $null; MailTo = 'mike@contoso.com' })
            $result = Get-DefaultAclHC -Sheet $sheet -SystemErrors ([ref]$errors)

            $errors.Count | Should -Be 0
            $result.Count | Should -Be 0
        }

        It 'flags ADObjectName without Permission' {
            $sheet = @([PSCustomObject]@{ ADObjectName = 'Orphaned'; Permission = $null })
            $null = Get-DefaultAclHC -Sheet $sheet -SystemErrors ([ref]$errors)

            $errors.Count | Should -Be 1
            $errors[0].Type | Should -Be 'FatalError'
            $errors[0].Message | Should -Match 'Orphaned'
        }

        It 'flags Permission without ADObjectName' {
            $sheet = @([PSCustomObject]@{ ADObjectName = $null; Permission = 'R' })
            $null = Get-DefaultAclHC -Sheet $sheet -SystemErrors ([ref]$errors)

            $errors.Count | Should -Be 1
            $errors[0].Type | Should -Be 'FatalError'
        }

        It 'flags invalid permission characters' {
            $sheet = @([PSCustomObject]@{ ADObjectName = 'IT'; Permission = 'X' })
            $null = Get-DefaultAclHC -Sheet $sheet -SystemErrors ([ref]$errors)

            $errors[0].Message | Should -Match "invalid permission 'X'"
        }

        It "rejects 'I' (ignore) in defaults" {
            $sheet = @([PSCustomObject]@{ ADObjectName = 'IT'; Permission = 'I' })
            $null = Get-DefaultAclHC -Sheet $sheet -SystemErrors ([ref]$errors)

            $errors.Count | Should -Be 1
            $errors[0].Message | Should -Match "invalid permission 'I'"
        }

        It 'flags duplicate ADObjectName entries' {
            $sheet = @(
                [PSCustomObject]@{ ADObjectName = 'IT'; Permission = 'L' }
                [PSCustomObject]@{ ADObjectName = 'IT'; Permission = 'R' }
            )
            $result = Get-DefaultAclHC -Sheet $sheet -SystemErrors ([ref]$errors)

            $errors.Count | Should -Be 1
            $errors[0].Name | Should -Be 'Duplicate default ACL entry'
            $result['IT'] | Should -Be 'L'
        }

        It 'normalizes case and trims whitespace' {
            $sheet = @([PSCustomObject]@{ ADObjectName = '  IT  '; Permission = ' l ' })
            $result = Get-DefaultAclHC -Sheet $sheet -SystemErrors ([ref]$errors)

            $errors.Count | Should -Be 0
            $result['IT'] | Should -Be 'L'
        }

        It 'continues processing after one bad row' {
            $sheet = @(
                [PSCustomObject]@{ ADObjectName = 'BadOne'; Permission = 'X' }
                [PSCustomObject]@{ ADObjectName = 'GoodOne'; Permission = 'L' }
            )
            $result = Get-DefaultAclHC -Sheet $sheet -SystemErrors ([ref]$errors)

            $errors.Count | Should -Be 1
            $result.Count | Should -Be 1
            $result['GoodOne'] | Should -Be 'L'
        }
    }
}