#requires -Modules Pester

Describe 'Utils.ps1 - Shared Utility Functions' {

    BeforeAll {
        $root = Split-Path -Parent $MyInvocation.MyCommand.Path
        $utils = Join-Path $root '../Modules/Toolbox.PermissionMatrixHC/Private/Utils.ps1'
        . $utils
    }


    Context 'Add-ErrorByCategoryHC' {
        It 'Adds properly formatted error objects' {
            $errors = @()
            Add-ErrorByCategoryHC -Type 'FatalError' -Name 'X' -Message 'Y' -Category 'Test' -SystemErrors ([ref]$errors)

            $errors.Count | Should -Be 1
            $errors[0].Type | Should -Be 'FatalError'
            $errors[0].Category | Should -Be 'Test'
        }
    }


    Context 'Add-MatrixErrorHC' {
        It 'Sets Category = Matrix' {
            $errors = @()
            Add-MatrixErrorHC -Type 'Warning' -Name 'W' -Message 'Msg' -SystemErrors ([ref]$errors)

            $errors[0].Category | Should -Be 'Matrix'
        }
    }


    Context 'Add-PermissionsErrorHC' {
        It 'Sets Category = Permissions' {
            $errors = @()
            Add-PermissionsErrorHC -Type 'FatalError' -Name 'N' -Message 'M' -SystemErrors ([ref]$errors)

            $errors[0].Category | Should -Be 'Permissions'
        }
    }


    Context 'Get-StringValueHC' {

        It 'Returns literal string for non-ENV values' {
            Get-StringValueHC -Name 'ABC' | Should -Be 'ABC'
        }

        It 'Resolves an environment variable' {
            $env:TEST_VALUE = 'Hello123'
            Get-StringValueHC -Name 'ENV:TEST_VALUE' | Should -Be 'Hello123'
        }

        It 'Throws when ENV variable missing' {
            { Get-StringValueHC -Name 'ENV:NOPEVAR' } | Should -Throw
        }
    }


    Context 'Plural' {
        It 'Returns plural form' {
            Plural -Count 5 -Word 'File' | Should -Be 'Files'
        }

        It 'Returns singular when Count=1' {
            Plural -Count 1 -Word 'File' | Should -Be 'File'
        }
    }


    Context 'Test-ItemHasFatalErrorHC' {
        It 'Detects FatalError in error list' {
            $errors = @(
                @{ Type = 'Warning' },
                @{ Type = 'FatalError' }
            )
            Test-ItemHasFatalErrorHC -CheckList $errors | Should -BeTrue
        }
    }


    Context 'Get-DatedLogFolderPathHC' {
        It 'Creates a folder and returns its path' {
            $folder = Join-Path $TestDrive 'logs'
            New-Item -ItemType Directory -Path $folder | Out-Null

            $result = Get-DatedLogFolderPathHC `
                -LogFolder $folder `
                -ScriptStartTime (Get-Date '2020-01-01T12:00:00') `
                -JsonFile @{ BaseName = 'MyConfig' }

            Test-Path $result | Should -BeTrue
        }
    }
}