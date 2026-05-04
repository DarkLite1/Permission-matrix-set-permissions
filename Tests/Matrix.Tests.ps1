#Requires -Version 7
#Requires -Modules Pester

Describe 'Matrix Logic Tests' {
    BeforeDiscovery {
        . "$PSScriptRoot\Helpers\Fixtures.Matrix.ps1"

        $script:MatrixSettingsFixtures = Get-MatrixSettingsFixtures
        $script:MatrixPermissionsFixtures = Get-MatrixPermissionsFixtures
        $script:DisabledMatrixFixtures = Get-DisabledMatrixFixtures
        $script:DuplicatePathFixtures = Get-DuplicateMatrixFixtures
        $script:DefaultMergeFixtures = Get-DefaultPermissionsMergeFixtures
        $script:ADBuildFixtures = Get-AdObjectBuildFixtures
        $script:MatrixBuildFixtures = Get-MatrixBuildFixtures

        $script:testScript = Join-Path `
            $PSScriptRoot `
            '..\Scripts\Entrypoints\PermissionMatrix.ps1'
    }

    BeforeAll {
        . "$PSScriptRoot\Helpers\Helpers.HC.ps1"
        . "$PSScriptRoot\Helpers\Fixtures.Json.ps1"
        . "$PSScriptRoot\Helpers\Fixtures.Matrix.ps1"
        . "$PSScriptRoot\Helpers\Fixtures.Excel.ps1"
        
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

        $scriptPath = @{
            TestRequirementsFile = (New-Item 'TestDrive:\TestReq.ps1' -ItemType File).FullName
            SetPermissionFile    = (New-Item 'TestDrive:\SetPerm.ps1' -ItemType File).FullName
            UpdateServiceNow     = (New-Item 'TestDrive:\SNOW.ps1' -ItemType File).FullName
        }

        $testParams = @{
            ConfigurationJsonFile = $jsonFile.FullName
            ScriptPath            = $scriptPath
        }

        $testInputTemplate | 
        ConvertTo-Json -Depth 20 | Set-Content $jsonFile.FullName

        # Share objects for tests
        $script:TestJsonFile = $jsonFile
        $script:TestInput = $testInputTemplate
        $script:TestScript = $testScript
        $script:TestParams = $testParams

        $modulePath = Join-Path $PSScriptRoot '..\Modules\PermissionMatrix\PermissionMatrix.psd1'
        Import-Module $modulePath -Force

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
    } -Tag test #-Skip

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
    } -Skip

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
    } -Skip

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
    } -Skip

    Describe 'Default permissions merging' {

        It '<Description>' -TestCases $DefaultMergeFixtures {
            param($Description, $DefaultsRows, $MatrixRows, $ExpectedMerged)

            $output = Merge-DefaultPermissionsHC -Defaults $DefaultsRows -Matrix $MatrixRows
            $output | Should -Be $ExpectedMerged
        }
    }

    # ------------------------------------------------------------------
    # 7. AD Object Build Logic
    # ------------------------------------------------------------------
    Describe 'AD Object build logic' {

        It '<Description>' -TestCases $ADBuildFixtures {
            param($Description, $FixtureBuilder, $Expected)

            $hash = & $FixtureBuilder
            $hash.Keys.Count | Should -Be $Expected
        }
    }

    # ------------------------------------------------------------------
    # 8. Matrix building logic
    # ------------------------------------------------------------------
    Describe 'Matrix building logic' {

        It '<Description>' -TestCases $MatrixBuildFixtures {
            param($Description, $FixtureBuilder, $ExpectedFiles)

            & $FixtureBuilder

            $updated = Copy-ObjectHC $TestInput
            Save-TestJson $updated $TestJsonFile

            & $TestScript @TestParams
            $LASTEXITCODE | Should -Be 0

            $logFolder = Get-LatestLogFolderHC -Root $TestInput.Settings.SaveLogFiles.Where.Folder
            $htmlFiles = Get-ChildItem -Path $logFolder -Recurse -Filter '*.html'

            $htmlFiles.Count | Should -Be $ExpectedFiles
        }
    }
}