#Requires -Version 7
#Requires -Modules Pester

Describe 'Input Validation Tests' {

    BeforeDiscovery {
        <#
        DISCOVERY CONTEXT
        - static only
        - no filesystem
        - no module loading
        #>

        . "$PSScriptRoot\Helpers\Fixtures.TestCases.ps1"

        $script:MissingTopLevelProps = Get-MissingTopLevelProperties
        $script:MissingMaxConcurrentProps = Get-MissingMaxConcurrentProperties
        $script:MissingMatrixProps = Get-MissingMatrixProperties
        $script:InvalidPathTests = Get-InvalidMatrixPaths

        $script:TestScript = Join-Path `
            $PSScriptRoot `
            '..\Scripts\Entrypoints\PermissionMatrix.ps1'
    }

    BeforeAll {
        <#
        EXECUTION CONTEXT
        - module loading happens via entrypoint script
        - filesystem + TestDrive allowed
        #>

        . "$PSScriptRoot\Helpers\Helpers.HC.ps1"
        . "$PSScriptRoot\Helpers\Fixtures.Json.ps1"

        if (-not (Test-Path $TestScript)) {
            throw "Script '$TestScript' not found"
        }

        # ------------------------------------------------------------------
        # Test input preparation
        # ------------------------------------------------------------------
        $jsonFile = New-Item 'TestDrive:\Input.json' -ItemType File

        $testInputTemplate = New-JsonFixture

        $testInputTemplate.Matrix.FolderPath =
        (New-Item 'TestDrive:\Matrix' -ItemType Directory).FullName

        $testInputTemplate.Matrix.DefaultsFile =
        (New-ValidDefaultsExcelFixture -Path 'TestDrive:\Defaults.xlsx')

        $testInputTemplate.Settings.SaveLogFiles.Where.Folder =
        (New-Item 'TestDrive:\Logs' -ItemType Directory).FullName

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
        ConvertTo-Json -Depth 20 |
        Set-Content $jsonFile.FullName

        # ------------------------------------------------------------------
        # Shared state
        # ------------------------------------------------------------------
        $script:TestJsonFile = $jsonFile
        $script:TestInput = $testInputTemplate
        $script:TestScript = $TestScript
        $script:TestParams = $testParams
    }

    BeforeEach {
        Mock Write-EventLog
        Mock Send-MailKitMessageHC
        Mock Invoke-Command

        Clear-TestLogFoldersHC `
            -ConfiguredLogFolder $TestInput.Settings.SaveLogFiles.Where.Folder
    }

    Describe 'missing top-level JSON properties' {
        It '<Property> should produce an error' -TestCases $MissingTopLevelProps {
            param($Property)

            $updated = Copy-ObjectHC $TestInput
            $updated.$Property = $null

            Save-TestJson -InputObject $updated -JsonFile $TestJsonFile
            & $TestScript @TestParams

            $LASTEXITCODE | Should -Be 1

            if ($Property -eq 'Settings') {
                $fallback = Get-FallbackLogFolderHC

                Assert-LogContainsSystemErrorHC `
                    -LogFolderPath $fallback `
                    -Pattern "*Property '$Property' not found*"
            }
            else {
                Assert-LogContainsSystemErrorHC `
                    -LogFolderPath $TestInput.Settings.SaveLogFiles.Where.Folder `
                    -Pattern "*Property '$Property' not found*"
            }
        }
    }

    Describe 'missing MaxConcurrent sub-properties' {
        It 'MaxConcurrent.<Property> not found' -TestCases $MissingMaxConcurrentProps {
            param($Property)

            $updated = Copy-ObjectHC $TestInput
            $updated.MaxConcurrent.$Property = $null

            Save-TestJson -InputObject $updated -JsonFile $TestJsonFile
            & $TestScript @TestParams

            $LASTEXITCODE | Should -Be 1

            Assert-LogContainsSystemErrorHC `
                -LogFolderPath $TestInput.Settings.SaveLogFiles.Where.Folder `
                -Pattern "*Property 'MaxConcurrent.$Property' must be numeric*"
        }
    }

    Describe 'missing Matrix sub-properties' {
        It 'Matrix.<Property> not found' -TestCases $MissingMatrixProps {
            param($Property)

            $updated = Copy-ObjectHC $TestInput
            $updated.Matrix.$Property = $null

            Save-TestJson -InputObject $updated -JsonFile $TestJsonFile
            & $TestScript @TestParams

            $LASTEXITCODE | Should -Be 1

            Assert-LogContainsSystemErrorHC `
                -LogFolderPath $TestInput.Settings.SaveLogFiles.Where.Folder `
                -Pattern "*Property 'Matrix.$Property' not found*"
        }
    }

    Describe 'invalid filesystem paths' {
        It 'fails when <Property> path is invalid' -TestCases $InvalidPathTests {
            param($Property, $Value)

            $updated = Copy-ObjectHC $TestInput
            Set-NestedPropertyHC -Object $updated -Path $Property -Value $Value

            Save-TestJson -InputObject $updated -JsonFile $TestJsonFile
            & $TestScript @TestParams

            $LASTEXITCODE | Should -Be 1

            Assert-LogContainsSystemErrorHC `
                -LogFolderPath $TestInput.Settings.SaveLogFiles.Where.Folder `
                -Pattern "*Property '$Property' path '$Value' not found*"
        }
    }

    Describe 'ScriptPath validation' {
        It 'ScriptPath.<Property> not found' -ForEach @(
            'TestRequirementsFile', 'SetPermissionFile', 'UpdateServiceNow'
        ) {
            $ScriptKey = $_

            $badParams = Copy-ObjectHC $TestParams
            $badParams.ScriptPath[$ScriptKey] = 'x:\doesnotexist.ps1'

            & $TestScript @badParams

            $LASTEXITCODE | Should -Be 1

            Assert-LogContainsSystemErrorHC `
                -LogFolderPath $TestInput.Settings.SaveLogFiles.Where.Folder `
                -Pattern "*The required script '$ScriptKey' was not found at 'x:\doesnotexist.ps1'*"
        }
    }

    Describe 'log folder creation failure' {
        It 'fails if SaveLogFiles.Where.Folder cannot be created' {

            $updated = Copy-ObjectHC $TestInput
            $updated.Settings.SaveLogFiles.Where.Folder = 'x:\nope'

            Save-TestJson -InputObject $updated -JsonFile $TestJsonFile

            & $TestScript @TestParams

            $LASTEXITCODE | Should -Be 1

            $fallback = Join-Path $env:TEMP 'PermissionMatrixLogs'

            Assert-LogContainsSystemErrorHC `
                -LogFolderPath $fallback `
                -Pattern "*Failed to create configured log folder 'x:\nope'*"
        }
    } -Tag test
}