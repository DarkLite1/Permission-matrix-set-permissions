#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
#Requires -Module ImportExcel

# -----------------------------------------------------------------------------
# Tests for Import-MatrixDefaultsFileHC.
#
# Strategy: the private collaborators (Add-ErrorHC, Get-DefaultAclHC,
# Test-ItemHasFatalErrorHC) run for real, so the SUT is driven through
# InModuleScope after Import-Module loads PermissionMatrix.psm1. Import-Excel
# also runs for real against a defaults workbook written to the real $TestDrive
# path (the function does its own Get-Item + Import-Excel, so the file must
# exist on disk with a 'Settings' worksheet).
#
# TWO ASSUMPTIONS worth checking against your code:
#   1. The default fixture rows use Permission = 'R' and a plain ADObjectName.
#      The success paths (happy path, "no MailTo") only pass if the REAL
#      Get-DefaultAclHC accepts that data without recording a FatalError. If
#      your ACL rules are stricter, adjust New-DefaultsExcelFixture's defaults.
#   2. Error assertions filter on .Name / .Type. If Add-ErrorHC stores the
#      -Name / -Type arguments under different property names, update the
#      Where-Object clauses accordingly.
# -----------------------------------------------------------------------------

Describe 'Import-MatrixDefaultsFileHC' {

    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"
        Import-Module "$moduleRoot\PermissionMatrix.psm1" -Force

        # Writes a defaults workbook to a real path and returns that path.
        function New-DefaultsExcelFixture {
            param(
                [Parameter(Mandatory)][string]$Path,
                [object[]]$Rows,
                [string]$WorksheetName = 'Settings'
            )

            if (-not $PSBoundParameters.ContainsKey('Rows')) {
                $Rows = @(
                    [pscustomobject]@{ MailTo = 'admin@contoso.com'; ADObjectName = 'Group Managers'; Permission = 'R' }
                    [pscustomobject]@{ MailTo = '  ops@contoso.com '; ADObjectName = 'Group Ops'; Permission = 'R' }
                    [pscustomobject]@{ MailTo = ''; ADObjectName = 'Group Readers'; Permission = 'R' }
                )
            }

            $Rows | Export-Excel -Path $Path -WorksheetName $WorksheetName -ClearSheet | Out-Null
            $Path
        }

        # Runs the SUT inside the module (real collaborators) and hands back both
        # the function output and the populated SystemErrors list.
        function Invoke-Sut {
            param([Parameter(Mandatory)][AllowNull()][string]$DefaultsFile)

            InModuleScope 'PermissionMatrix' -Parameters @{ DefaultsFile = $DefaultsFile } {
                param($DefaultsFile)

                $errors = [System.Collections.Generic.List[object]]::new()
                $matrix = [pscustomobject]@{ DefaultsFile = $DefaultsFile }

                $output = Import-MatrixDefaultsFileHC -Matrix $matrix -SystemErrors ([ref]$errors)

                [pscustomobject]@{ Output = $output; Errors = @($errors) }
            }
        }
    }

    Context 'when the defaults file cannot be read' {
        It 'records a FatalError and returns $null when the file does not exist' {
            $path = Join-Path $TestDrive 'does-not-exist.xlsx'

            $result = Invoke-Sut -DefaultsFile $path

            $result.Output | Should -BeNullOrEmpty
            $check = $result.Errors | Where-Object Name -EQ 'Defaults file not found'
            $check | Should -Not -BeNullOrEmpty
            $check.Type | Should -Be 'FatalError'
        }

        It "records a FatalError and returns `$null when the 'Settings' worksheet is missing" {
            # A workbook that exists but only has a differently-named sheet.
            $path = New-DefaultsExcelFixture `
                -Path (Join-Path $TestDrive 'no-settings.xlsx') `
                -WorksheetName 'SomethingElse'

            $result = Invoke-Sut -DefaultsFile $path

            $result.Output | Should -BeNullOrEmpty
            ($result.Errors | Where-Object Name -EQ 'Defaults worksheet missing') |
                Should -Not -BeNullOrEmpty
        }
    }

    Context 'mandatory columns' {
        It 'records an "Invalid defaults format" FatalError when the <_> column is missing' -ForEach @(
            'MailTo', 'ADObjectName', 'Permission'
        ) {
            # Start from a complete row, then drop the column under test.
            $row = [ordered]@{
                MailTo       = 'admin@contoso.com'
                ADObjectName = 'Group Managers'
                Permission   = 'R'
            }
            $row.Remove($_)

            $path = New-DefaultsExcelFixture `
                -Path (Join-Path $TestDrive "missing-$_.xlsx") `
                -Rows @([pscustomobject]$row)

            $result = Invoke-Sut -DefaultsFile $path

            $result.Output | Should -BeNullOrEmpty
            $check = $result.Errors | Where-Object Name -EQ 'Invalid defaults format'
            $check | Should -Not -BeNullOrEmpty
            $check.Type | Should -Be 'FatalError'
        }
    }

    Context 'when no usable MailTo address is present' {
        It 'records a "No MailTo addresses" FatalError when every MailTo is blank' {
            # Columns are valid (so Get-DefaultAclHC passes) but no MailTo value.
            $rows = @(
                [pscustomobject]@{ MailTo = ''; ADObjectName = 'Group Managers'; Permission = 'R' }
                [pscustomobject]@{ MailTo = '   '; ADObjectName = 'Group Ops'; Permission = 'R' }
            )
            $path = New-DefaultsExcelFixture `
                -Path (Join-Path $TestDrive 'no-mailto.xlsx') -Rows $rows

            $result = Invoke-Sut -DefaultsFile $path

            $result.Output | Should -BeNullOrEmpty
            ($result.Errors | Where-Object Name -EQ 'No MailTo addresses') |
                Should -Not -BeNullOrEmpty
        }
    }

    Context 'when Get-DefaultAclHC reports a fatal error' {
        It 'returns $null and reads no further once the ACL check is fatal' {
            $path = New-DefaultsExcelFixture -Path (Join-Path $TestDrive 'acl-fatal.xlsx')

            $result = InModuleScope 'PermissionMatrix' -Parameters @{ DefaultsFile = $path } {
                param($DefaultsFile)

                # Force the real ACL builder to record a fatal error. It still uses
                # the real Add-ErrorHC against the SUT's own SystemErrors ref.
                Mock Get-DefaultAclHC {
                    Add-ErrorHC `
                        -Type 'FatalError' `
                        -Name 'Default ACL invalid' `
                        -Message 'Injected ACL failure.' `
                        -Category 'Matrix' `
                        -SystemErrors $SystemErrors
                }

                $errors = [System.Collections.Generic.List[object]]::new()
                $matrix = [pscustomobject]@{ DefaultsFile = $DefaultsFile }
                $output = Import-MatrixDefaultsFileHC -Matrix $matrix -SystemErrors ([ref]$errors)

                [pscustomobject]@{ Output = $output; Errors = @($errors) }
            }

            $result.Output | Should -BeNullOrEmpty
            ($result.Errors | Where-Object Name -EQ 'Default ACL invalid') |
                Should -Not -BeNullOrEmpty
            # It bailed before the MailTo loop, so no MailTo error was added.
            ($result.Errors | Where-Object Name -EQ 'No MailTo addresses') |
                Should -BeNullOrEmpty
        }
    }

    Context 'the happy path' {
        It 'returns a defaults object carrying FilePath, DefaultAcl and MailTo' {
            $path = New-DefaultsExcelFixture -Path (Join-Path $TestDrive 'valid.xlsx')

            $result = Invoke-Sut -DefaultsFile $path

            $result.Output | Should -Not -BeNullOrEmpty
            $result.Output.FilePath | Should -Be (Get-Item -LiteralPath $path).FullName
            $result.Output.DefaultAcl | Should -Not -BeNullOrEmpty
            # No FatalError was recorded on the success path.
            ($result.Errors | Where-Object Type -EQ 'FatalError') | Should -BeNullOrEmpty
        }

        It 'collects only non-blank MailTo values and trims surrounding whitespace' {
            $path = New-DefaultsExcelFixture -Path (Join-Path $TestDrive 'mailto.xlsx')

            $result = Invoke-Sut -DefaultsFile $path

            # Default fixture: admin@ (kept), '  ops@ ' (kept, trimmed), '' (skipped).
            $result.Output.MailTo | Should -HaveCount 2
            $result.Output.MailTo | Should -Contain 'admin@contoso.com'
            $result.Output.MailTo | Should -Contain 'ops@contoso.com'
        }
    }

    Context 'unexpected failure' {
        It 'records a catch-all "Defaults import failed" FatalError when a collaborator throws' {
            $path = New-DefaultsExcelFixture -Path (Join-Path $TestDrive 'boom.xlsx')

            $result = InModuleScope 'PermissionMatrix' -Parameters @{ DefaultsFile = $path } {
                param($DefaultsFile)

                Mock Get-DefaultAclHC { throw 'unexpected kaboom' }

                $errors = [System.Collections.Generic.List[object]]::new()
                $matrix = [pscustomobject]@{ DefaultsFile = $DefaultsFile }
                $output = Import-MatrixDefaultsFileHC -Matrix $matrix -SystemErrors ([ref]$errors)

                [pscustomobject]@{ Output = $output; Errors = @($errors) }
            }

            $result.Output | Should -BeNullOrEmpty
            $check = $result.Errors | Where-Object Name -EQ 'Defaults import failed'
            $check | Should -Not -BeNullOrEmpty
            $check.Type | Should -Be 'FatalError'
        }
    }
}
