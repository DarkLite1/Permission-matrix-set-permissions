#Requires -Version 7
#Requires -Modules Pester

Describe 'Matrix Logic Tests' {
    BeforeDiscovery {
        . "$PSScriptRoot\Helpers\Fixtures.Matrix.ps1"

        $script:MatrixSettingsFixtures = Get-MatrixSettingsFixtures
        $script:MatrixPermissionsFixtures = Get-MatrixPermissionsFixtures
        $script:DisabledMatrixFixtures = Get-DisabledMatrixFixtures
        $script:DuplicatePathFixtures = Get-DuplicateMatrixFixtures
        $script:AclFixtures = Get-AclConversionFixtures
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
    }

    BeforeEach {
        Mock Write-EventLog
        Mock Send-MailKitMessageHC
        # DO NOT MOCK Invoke-Command

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

            $LASTEXITCODE | Should -Be 1

            Assert-LogContainsSystemErrorHC `
                -LogFolderPath $TestInput.Settings.SaveLogFiles.Where.Folder `
                -Pattern "*$ExpectedMessage*"
        } -Tag test
    }

    Describe 'Matrix: Permissions sheet validation' {
        It '<Issue> should be detected' -TestCases $MatrixPermissionsFixtures {
            param($Issue, $Mutation, $Expected)

            Invoke-Expression $Mutation | Out-Null

            $updated = Copy-ObjectHC $TestInput
            Save-TestJson $updated $TestJsonFile

            & $TestScript @TestParams
            $LASTEXITCODE | Should -Be 1

            Assert-LogContainsSystemErrorHC `
                -LogFolderPath $TestInput.Settings.SaveLogFiles.Where.Folder `
                -Pattern "*$Expected*"
        }
    }

    # ------------------------------------------------------------------
    # 3. Disabled Matrices
    # ------------------------------------------------------------------
    Describe 'Matrix: Disabled matrices' {

        It '<Description>' -TestCases $DisabledMatrixFixtures {
            param($Description, $FixtureBuilder, $ExpectedCount)

            & $FixtureBuilder  # creates .xlsx files in TestDrive:\Matrix

            $updated = Copy-ObjectHC $TestInput
            Save-TestJson $updated $TestJsonFile

            & $TestScript @TestParams
            $LASTEXITCODE | Should -Be 0

            $logFolder = Get-LatestLogFolderHC -Root $TestInput.Settings.SaveLogFiles.Where.Folder
            $htmlFile = Get-ChildItem -Path $logFolder -Recurse -Filter '*.html' | Select-Object -First 1

            $htmlFile | Should -Not -BeNullOrEmpty

            (Get-Content $htmlFile.FullName -Raw) |
            Should -Match "Processed $ExpectedCount enabled matrix"
        }
    }

    # ------------------------------------------------------------------
    # 4. Duplicate combinations
    # ------------------------------------------------------------------
    Describe 'Matrix: Duplicate combinations' {

        It 'detects duplicates correctly' -TestCases $DuplicatePathFixtures {
            param($FixtureBuilder, $ExpectedError)

            & $FixtureBuilder

            $updated = Copy-ObjectHC $TestInput
            Save-TestJson $updated $TestJsonFile

            & $TestScript @TestParams
            $LASTEXITCODE | Should -Be 1

            Assert-LogContainsSystemErrorHC `
                -LogFolderPath $TestInput.Settings.SaveLogFiles.Where.Folder `
                -Pattern "*$ExpectedError*"
        }
    }

    # ------------------------------------------------------------------
    # 5. ACL Conversion (now requires SettingsRows matching real structure)
    # ------------------------------------------------------------------
    Describe 'ACL conversion' {

        It '<Description>' -TestCases $AclFixtures {
            param($Description, $SettingsRows, $ExpectedAclCount)

            # Create full matrix file
            $path = 'TestDrive:\Matrix\File1.xlsx'

            New-MatrixExcelFixture `
                -Path $path `
                -SettingsRows $SettingsRows

            # Script doesn't need execution for this
            $result = ConvertTo-MatrixAclHC -Sheet $SettingsRows
            $result.Count | Should -Be $ExpectedAclCount
        }
    }

    # ------------------------------------------------------------------
    # 6. Default permissions merging
    # ------------------------------------------------------------------
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