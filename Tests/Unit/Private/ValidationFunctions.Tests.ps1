#requires -Modules Pester

Describe 'Validation.ps1 - Updated Validation Functions' {

    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        . "$moduleRoot\Private\Utils.ps1"
        . "$moduleRoot\Private\Validation.ps1"
    }

    #
    # Matrix-level validation
    #

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
        It 'Errors when <4 rows' {
            (Test-MatrixPermissionsHC -Permissions @('a', 'b')).Type | Should -Be 'FatalError'
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


    #
    # AD validation
    #
    Context 'Test-AdObjectsHC' {
        It 'Warns if AD object missing' {
            $res = Test-AdObjectsHC -ADObjects @('A', 'B') -AdInfo @('A')
            $res.Type | Should -Contain 'Warning'
        }
    }


    #
    # Expanded matrix validation
    #
    Context 'Test-ExpandedMatrixHC' {
        It 'Warns for unknown ACL principals' {
            $mat = @(
                [pscustomobject]@{
                    ACL = @{ MissingUser = 'R' }
                }
            )
            $res = Test-ExpandedMatrixHC `
                -Matrix $mat `
                -ADObject @('GoodUser') `
                -DefaultAcl @{} `
                -AdGroupPlaceHolders @()

            $res.Type | Should -Contain 'Warning'
        }
    }


    #
    # JSON Schema validation
    #
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


    #
    # Runtime settings validation
    #
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
                DefaultsFile           = $PSCommandPath
                FolderPath             = 'C:\'
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