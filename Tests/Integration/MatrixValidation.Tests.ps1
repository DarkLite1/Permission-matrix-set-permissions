#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

# -----------------------------------------------------------------------------
# Integration tests for matrix validation and building. These drive the real
# entrypoint (Scripts\Entrypoints\PermissionMatrix.ps1) end to end and assert
# on the HTML execution report, so they live in Tests\Integration\ rather than
# in Tests\Unit\Private\. The per-function unit tests for Matrix.ps1 live in
# Tests\Unit\Private\Matrix.Tests.ps1.
#
# Moved here from the old Matrix.Tests.ps1 unchanged except for path depth
# ($PSScriptRoot is now two levels under the repo root, not three).
# -----------------------------------------------------------------------------

Describe 'Matrix validation (integration)' {
    BeforeDiscovery {
        . "$PSScriptRoot/../Helpers/Fixtures.Matrix.ps1"

        $script:MatrixSettingsFixtures = Get-MatrixSettingsFixtures
        $script:MatrixPermissionsFixtures = Get-MatrixPermissionsFixtures
        $script:DisabledMatrixFixtures = Get-DisabledMatrixFixtures
        $script:DuplicatePathFixtures = Get-DuplicateMatrixFixtures
        $script:MatrixBuildFixtures = Get-MatrixBuildFixtures

        $script:testScript = "$PSScriptRoot\..\..\Scripts\Entrypoints\PermissionMatrix.ps1"
    }

    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        # Copy-ObjectHC and friends are needed at test scope; the module under
        # test is exercised through the entrypoint, so Matrix.ps1 is not
        # dot-sourced here.
        . "$moduleRoot\Private\Utils.ps1"

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
            $htmlFiles = Get-ChildItem -Path $logFolder -Recurse -Filter '00 - Execution Report.html'

            $htmlFiles.Count | Should -Be $ExpectedFiles
        }
    }
}
