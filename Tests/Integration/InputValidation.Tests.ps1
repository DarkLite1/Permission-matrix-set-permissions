#Requires -Version 7
#Requires -Modules Pester

Describe 'Input Validation Tests' {
    BeforeDiscovery {
        $root = Resolve-Path "$PSScriptRoot\..\.."

        . "$root\Tests\Helpers\Fixtures.TestCases.ps1"

        $script:MissingTopLevelProps = Get-MissingTopLevelProperties
        $script:MissingMaxConcurrentProps = Get-MissingMaxConcurrentProperties
        $script:MissingMatrixProps = Get-MissingMatrixProperties
        $script:InvalidPathTests = Get-InvalidMatrixPaths
    }

    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        Import-Module "$moduleRoot\PermissionMatrix.psm1" -Force

        . "$root\Tests\Helpers\Helpers.HC.ps1"        
        . "$root\Tests\Helpers\Fixtures.Json.ps1"

        $script:TestScript = "$root\Scripts\Entrypoints\PermissionMatrix.ps1"

        if (-not (Test-Path $TestScript)) {
            throw "Script '$TestScript' not found"
        }

        # Template config — each test clones this, modifies one property, saves, invokes.
        $script:TestJsonFile = New-Item 'TestDrive:\Input.json' -ItemType File

        $script:TestInput = New-JsonFixture
        $TestInput.Matrix.FolderPath = (New-Item 'TestDrive:\Matrix' -ItemType Directory).FullName
        $TestInput.Matrix.DefaultsFile = (New-ValidDefaultsExcelFixture -Path 'TestDrive:\Defaults.xlsx')
        $TestInput.Settings.SaveLogFiles.Where.Folder = (New-Item 'TestDrive:\Logs' -ItemType Directory).FullName

        $TestInput | ConvertTo-Json -Depth 20 | 
        Set-Content $TestJsonFile.FullName

        $script:TestParams = @{
            ConfigurationJsonFile = $TestJsonFile.FullName
        }

        # Run the entrypoint with a modified config and return its exit code.
        function Invoke-WithConfig {
            param([pscustomobject]$Config)

            Save-TestJson -InputObject $Config -JsonFile $TestJsonFile
            & $TestScript @TestParams
            return $LASTEXITCODE
        }
    }

    BeforeEach {
        Mock Write-EventLog -ModuleName PermissionMatrix
        Mock Send-MailKitMessageHC -ModuleName PermissionMatrix
        Mock Invoke-Command -ModuleName PermissionMatrix

        Clear-TestLogFoldersHC `
            -ConfiguredLogFolder $TestInput.Settings.SaveLogFiles.Where.Folder
    }

    Context 'missing top-level JSON properties' {
        It '<Property> should produce an error' -TestCases $MissingTopLevelProps {
            param($Property)

            $updated = Copy-ObjectHC $TestInput
            $updated.$Property = $null

            (Invoke-WithConfig $updated) | Should -Be 1

            # When Settings is null, the configured log folder is unreachable —
            # the script falls back to a temp folder for logging.
            $logFolder = if ($Property -eq 'Settings') {
                Get-FallbackLogFolderHC
            }
            else {
                $TestInput.Settings.SaveLogFiles.Where.Folder
            }

            Assert-LogContainsSystemErrorHC `
                -LogFolderPath $logFolder `
                -Pattern "*Property '$Property' not found*"
        }
    }

    Context 'missing MaxConcurrent sub-properties' {
        It 'MaxConcurrent.<Property> not found' -TestCases $MissingMaxConcurrentProps {
            param($Property)

            $updated = Copy-ObjectHC $TestInput
            $updated.MaxConcurrent.$Property = $null

            (Invoke-WithConfig $updated) | Should -Be 1

            Assert-LogContainsSystemErrorHC `
                -LogFolderPath $TestInput.Settings.SaveLogFiles.Where.Folder `
                -Pattern "*Property 'MaxConcurrent.$Property' must be numeric*"
        }
    }

    Context 'missing Matrix sub-properties' {
        It 'Matrix.<Property> not found' -TestCases $MissingMatrixProps {
            param($Property)

            $updated = Copy-ObjectHC $TestInput
            $updated.Matrix.$Property = $null

            (Invoke-WithConfig $updated) | Should -Be 1

            Assert-LogContainsSystemErrorHC `
                -LogFolderPath $TestInput.Settings.SaveLogFiles.Where.Folder `
                -Pattern "*Property 'Matrix.$Property' not found*"
        }
    }

    Context 'invalid filesystem paths' {
        It 'fails when <Property> path is invalid' -TestCases $InvalidPathTests {
            param($Property, $Value)

            $updated = Copy-ObjectHC $TestInput
            Set-NestedPropertyHC -Object $updated -Path $Property -Value $Value

            (Invoke-WithConfig $updated) | Should -Be 1

            Assert-LogContainsSystemErrorHC `
                -LogFolderPath $TestInput.Settings.SaveLogFiles.Where.Folder `
                -Pattern "*Property '$Property' path '$Value' not found*"
        }
    }

    Context 'log folder creation failure' {
        It 'falls back to temp folder when configured log folder cannot be created' {
            $updated = Copy-ObjectHC $TestInput
            $updated.Settings.SaveLogFiles.Where.Folder = 'x:\nope'

            (Invoke-WithConfig $updated) | Should -Be 1

            $fallback = Join-Path $env:TEMP 'PermissionMatrixLogs'

            Assert-LogContainsSystemErrorHC `
                -LogFolderPath $fallback `
                -Pattern "*Failed to create configured log folder 'x:\nope'*"
        }
    }
}