#Requires -Version 7
#Requires -Modules Pester

Describe 'Input Validation Tests' {
    BeforeDiscovery {
        $root = Resolve-Path "$PSScriptRoot\..\..\..\.."

        . "$root\Tests\Helpers\Fixtures.TestCases.ps1"

        $script:MissingTopLevelProps = Get-MissingTopLevelProperties
        $script:MissingMaxConcurrentProps = Get-MissingMaxConcurrentProperties
        $script:MissingMatrixProps = Get-MissingMatrixProperties
        $script:InvalidPathTests = Get-InvalidMatrixPaths

        $script:TestScript = "$root\Scripts\Entrypoints\PermissionMatrix.ps1"
    }

    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\..\.."

        . "$root\Tests\Helpers\Helpers.HC.ps1"
        . "$root\Tests\Helpers\Fixtures.Json.ps1"

        if (-not (Test-Path $TestScript)) {
            throw "Script '$TestScript' not found"
        }

        $jsonFile = New-Item 'TestDrive:\Input.json' -ItemType File

        $testInputTemplate = New-JsonFixture

        $testInputTemplate.Matrix.FolderPath =
        (New-Item 'TestDrive:\Matrix' -ItemType Directory).FullName

        $testInputTemplate.Matrix.DefaultsFile =
        (New-ValidDefaultsExcelFixture -Path 'TestDrive:\Defaults.xlsx')

        $testInputTemplate.Settings.SaveLogFiles.Where.Folder =
        (New-Item 'TestDrive:\Logs' -ItemType Directory).FullName

        $testParams = @{
            ConfigurationJsonFile = $jsonFile.FullName
        }

        $testInputTemplate |
        ConvertTo-Json -Depth 20 | Set-Content $jsonFile.FullName
        
        # Share objects for tests
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

            if ($Property -eq 'Settings') {
                'pause'
            }

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
        } -Tag test
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

    Describe 'log folder creation failure' {
        It 'fallback to temp folder when log folder creation fails' {

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
    }
}