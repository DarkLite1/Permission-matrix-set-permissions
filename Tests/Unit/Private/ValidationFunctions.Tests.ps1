#Requires -Module @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
#Requires -Module ImportExcel

Describe 'Validation.ps1 - Updated Validation Functions' {
    BeforeDiscovery {
        . "$PSScriptRoot/../../Helpers/Fixtures.Matrix.ps1"

        $script:PermissionFixtures = Get-MatrixPermissionsFixtures
    }

    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        . "$moduleRoot\Private\Utils.ps1"
        . "$moduleRoot\Private\Validation.ps1"
        . "$root/Tests/Helpers/Fixtures.Excel.ps1"
        . "$root/Tests/Helpers/Fixtures.Matrix.ps1"

        # Build a Permissions sheet from a scenario, round-trip it through Excel,
        # and return the imported objects exactly as production sees them.
        function Get-RoundTripPermissions {
            param([Parameter(Mandatory)][string]$Scenario)

            $spec = New-MatrixPermissionsFixtureRows -Scenario $Scenario

            $dir = Join-Path 'TestDrive:' 'Matrix'
            if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
            $path = Join-Path $dir "Permissions_$Scenario.xlsx"

            New-MatrixPermissionsExcelFixture -Path $path -Spec $spec | Out-Null

            return @(Import-Excel -Path $path -WorksheetName 'Permissions' -NoHeader -DataOnly -ErrorAction Stop)
        }
    }

    Context 'Test-MatrixFileHC' {
        It 'Warns for missing settings' {
            $M = @{ Settings = @(); Permissions = @('x') }
            $res = Test-MatrixFileHC -MatrixObject $M
            $res.Type | Should -Contain 'Warning'
        }

        It 'Errors for missing permissions' {
            $M = @{ Settings = @('x'); Permissions = @() }
            $res = Test-MatrixFileHC -MatrixObject $M
            $res.Type | Should -Contain 'FatalError'
        }
    }

    Context 'Test-MatrixPermissionsHC' {

        Context 'Happy path' {
            It 'returns nothing when the Valid fixture is supplied' {
                $perms = Get-RoundTripPermissions -Scenario 'Valid'

                $result = Test-MatrixPermissionsHC -Permissions $perms

                # Function only returns $checks when Count -gt 0, so success => $null.
                $result | Should -BeNullOrEmpty
            }
        }

        Context 'Data-driven checks from Get-MatrixPermissionsFixtures' {
            It 'flags <Issue> with check name <Expected>' -ForEach $PermissionFixtures {

                # The fixture 'Mutation' strings map 1:1 to a scenario name in
                # New-MatrixPermissionsFixtureRows; derive it from the Issue so we
                # can round-trip in-process rather than Invoke-Expression a string
                # that writes its own file.
                $scenario = switch ($Issue) {
                    'MissingADObjectName' { 'MissingADObjectName' }
                    'InvalidPermissionChar' { 'InvalidPermissionChar' }
                    'MissingRows' { 'MissingRows' }
                    'MissingColumns' { 'MissingColumns' }
                    'MissingFolderName' { 'MissingFolderName' }
                    'DuplicateFolderName' { 'DuplicateFolderName' }
                    'InaccessibleFolders' { 'InaccessibleFolders' }
                    default { throw "No scenario mapping for Issue '$Issue'" }
                }

                $perms = Get-RoundTripPermissions -Scenario $scenario

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $result | Should -Not -BeNullOrEmpty
                ($result.Name) | Should -Contain $Expected
            }
        }

        Context 'fatal errors exit immediately' {
            It 'returns ONLY "Missing rows" for the MissingRows fixture' {
                $perms = Get-RoundTripPermissions -Scenario 'MissingRows'

                $result = Test-MatrixPermissionsHC -Permissions $perms

                @($result).Count | Should -Be 1
                $result[0].Type | Should -Be 'FatalError'
                $result[0].Name | Should -Be 'Missing rows'
            }

            It 'returns ONLY "Missing columns" for the MissingColumns fixture' {
                $perms = Get-RoundTripPermissions -Scenario 'MissingColumns'

                $result = Test-MatrixPermissionsHC -Permissions $perms

                @($result).Count | Should -Be 1
                $result[0].Type | Should -Be 'FatalError'
                $result[0].Name | Should -Be 'Missing columns'
            }
        }

        Context 'Check types are correct' {
            It 'classifies InaccessibleFolders as a Warning, not a FatalError' {
                $perms = Get-RoundTripPermissions -Scenario 'InaccessibleFolders'
                $result = Test-MatrixPermissionsHC -Permissions $perms

                $warn = $result | Where-Object Name -EQ 'Inaccessible folders'
                $warn | Should -Not -BeNullOrEmpty
                $warn.Type | Should -Be 'Warning'
            }

            It 'classifies InvalidPermissionChar as a FatalError' {
                $perms = Get-RoundTripPermissions -Scenario 'InvalidPermissionChar'
                $result = Test-MatrixPermissionsHC -Permissions $perms

                $err = $result | Where-Object Name -EQ 'Invalid permission character'
                $err | Should -Not -BeNullOrEmpty
                $err.Type | Should -Be 'FatalError'
            }
        }

        Context 'Error handling' {
            It 'rejects an empty array at parameter binding' {
                { Test-MatrixPermissionsHC -Permissions @() } |
                Should -Throw -ExpectedMessage '*empty array*'
            }
        }
    }

    Context 'Test-MatrixFormDataHC' {
        It 'Warns if FormData missing' {
            (Test-MatrixFormDataHC -FormData $null).Type | Should -Be 'Warning'
        }
    }

    Context 'Test-MatrixSettingRowHC' {
        It 'Validates missing properties' {
            $S = @{ }
            $r = Test-MatrixSettingRowHC -SettingRow $S
            $r.Type | Should -Contain 'FatalError'
        }
    }

    Context 'Test-AdObjectsHC' {
        It 'Warns if AD object missing' {
            $res = Test-AdObjectsHC -ADObjects @('A', 'B') -AdInfo @('A')
            $res.Type | Should -Contain 'Warning'
        }
    }

    Context 'Validate-ConfigurationStructure' {

        It 'Calls Add-JsonSchemaErrorHC for missing required properties' {

            Mock Add-JsonSchemaErrorHC

            $json = @{
                Matrix                 = @{}
                Export                 = $null
                Settings               = @{}
                ServiceNow             = $null
                MaxConcurrent          = @{}
                PSSessionConfiguration = @{}
            }

            $sys = @()
            Validate-ConfigurationStructure -Json $json -SystemErrors ([ref]$sys)

            Should -Invoke Add-JsonSchemaErrorHC -Times 1
        }
    }

    Context 'Validate-RuntimeSettings' {

        It 'Warns when ScriptName missing and adds default' {

            Mock Add-RuntimeErrorHC

            $settings = @{
                ScriptName     = $null
                SaveLogFiles   = @{ Detailed = $true }
                SaveInEventLog = @{ Save = $true }
                SendMail       = @{
                    From = 'a'
                    To   = 'b'
                    Body = 'c'
                    Smtp = @{ Port = '25'; ConnectionType = 'None' }
                }
            }

            $matrix = @{
                DefaultsFile        = $PSCommandPath
                FolderPath          = 'C:\'
                AdGroupPlaceHolders = @()
            }

            $export = @{}
            $sn = @{}
            $maxcon = @{ Computers = '1'; FoldersPerMatrix = '1'; JobsPerRemoteComputer = '1' }

            $sys = @()
            Validate-RuntimeSettings `
                -Settings $settings `
                -Matrix $matrix `
                -Export $export `
                -ServiceNow $sn `
                -MaxConcurrent $maxcon `
                -SystemErrors ([ref]$sys)

            Should -Invoke Add-RuntimeErrorHC -Times 1
            $settings.ScriptName | Should -Be 'Default script name'
        }
    }
}