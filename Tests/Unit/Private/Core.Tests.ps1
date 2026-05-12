#requires -Modules Pester

Describe 'Core.ps1 - Parallel Check Engine' {

    BeforeAll {
        $root = Split-Path -Parent $MyInvocation.MyCommand.Path
        $coreFile = Join-Path $root '../Modules/Toolbox.PermissionMatrixHC/Private/Core.ps1'
        . $coreFile
    }

    BeforeEach {
        # Mock all validation & helper functions
        Mock Test-MatrixFileHC { return @('FileCheck') }
        Mock Test-MatrixPermissionsHC { return @('PermCheck') }
        Mock Test-MatrixFormDataHC { return @('FormCheck') }
        Mock Test-MatrixSettingHC { return @('SettingCheck') }
        Mock ConvertTo-MatrixADNamesHC { return @('AD1', 'AD2') }
        Mock Test-AdObjectsHC { return @('AdCheck') }
        Mock Test-ExpandedMatrixHC { return @('ExpCheck') }
        Mock Get-ADObjectDetailHC { return @('AD Info Object') }
    }

    Context 'ConvertTo-WorkItemsHC' {
        It 'Produces DTOs for each matrix' {
            $import = @(
                [pscustomobject]@{
                    File        = @{ Item = @{Name = 'A.xlsx' }; LogFolder = 'L1' }
                    Permissions = @{ Import = @('P') }
                    FormData    = @{ Import = @{Property = 1 } }
                    Settings    = @(
                        @{ Import = @{Cn = 1 }; Matrix = @() }
                    )
                }
            )

            $res = ConvertTo-WorkItemsHC -ImportedMatrix $import
            $res.Count | Should -Be 1
            $res[0].Settings.Count | Should -Be 1
            $res[0].Permissions[0] | Should -Be 'P'
        }
    }


    Context 'Invoke-MatrixPhase1ParallelHC' {

        It 'Calls all Phase 1 validators' {

            $work = @(
                @{
                    FileMeta    = @{}
                    Permissions = @('P')
                    FormData    = @{Value = 1 }
                    Settings    = @( @{Import = @{X = 1 }; Matrix = @() } )
                }
            )

            $res = Invoke-MatrixPhase1ParallelHC -WorkItems $work -Throttle 1

            # Ensure mocks invoked
            Should -Invoke Test-MatrixFileHC -Times 1
            Should -Invoke Test-MatrixPermissionsHC -Times 1
            Should -Invoke Test-MatrixFormDataHC -Times 1
            Should -Invoke Test-MatrixSettingHC -Times 1
            Should -Invoke ConvertTo-MatrixADNamesHC -Times 1

            $res[0].FileChecks | Should -Contain 'FileCheck'
            $res[0].PermissionChecks | Should -Contain 'PermCheck'
            $res[0].FormDataChecks | Should -Contain 'FormCheck'
            $res[0].Settings[0].PreAdChecks | Should -Contain 'SettingCheck'
        }
    }


    Context 'Invoke-MatrixPhase2ParallelHC' {

        It 'Performs AD + Expanded checks' {

            $phase1 = @(
                @{
                    Settings = @(
                        @{
                            Import        = @{}
                            Matrix        = @('Matrix')
                            PreAdChecks   = @()
                            AdIdentifiers = @('AD1')
                        }
                    )
                }
            )

            $res = Invoke-MatrixPhase2ParallelHC `
                -WorkItems $phase1 `
                -AdInfo 'AD Object' `
                -DefaultAcl @{} `
                -AdGroupPlaceHolders @() `
                -Throttle 1

            Should -Invoke Test-AdObjectsHC -Times 1
            Should -Invoke Test-ExpandedMatrixHC -Times 1

            $res[0].Settings[0].AdChecks | Should -Contain 'AdCheck'
            $res[0].Settings[0].ExpandedChecks | Should -Contain 'ExpCheck'
        }
    }


    Context 'Merge-CheckResultsHC' {
        It 'Merges Phase1 and Phase2 results correctly' {

            $import = @(
                [pscustomobject]@{
                    File        = @{Check = @() }
                    Permissions = @{Check = @() }
                    FormData    = @{Check = @() }
                    Settings    = @(
                        [pscustomobject]@{Check = @(); Import = @{} }
                    )
                }
            )

            $p1 = @(
                @{
                    FileChecks       = 'File'
                    PermissionChecks = 'Perm'
                    FormDataChecks   = 'FD'
                    Settings         = @(
                        @{ PreAdChecks = 'Pre' }
                    )
                }
            )

            $p2 = @(
                @{
                    Settings = @(
                        @{ AdChecks = 'AC'; ExpandedChecks = 'EC' }
                    )
                }
            )

            $final = Merge-CheckResultsHC -ImportedMatrix $import -Phase1 $p1 -Phase2 $p2

            $final[0].File.Check | Should -Contain 'File'
            $final[0].Permissions.Check | Should -Contain 'Perm'
            $final[0].FormData.Check | Should -Contain 'FD'
            $final[0].Settings[0].Check | Should -Contain 'Pre'
            $final[0].Settings[0].Check | Should -Contain 'AC'
            $final[0].Settings[0].Check | Should -Contain 'EC'
        }
    }


    Context 'Invoke-MatrixChecksHC - Full Flow' {
        It 'Runs entire check pipeline' {

            # Simple fake matrix input
            $import = @(
                [pscustomobject]@{
                    File        = @{Item = @{Name = 'A.xlsx' }; LogFolder = 'X'; Check = @() }
                    Permissions = @{Import = @('P'); Check = @() }
                    FormData    = @{Import = @{X = 1 }; Check = @() }
                    Settings    = @(
                        [pscustomobject]@{Import = @{GroupName = 'G'; SiteCode = 'S' }; Matrix = @(); Check = @() }
                    )
                }
            )

            $result = Invoke-MatrixChecksHC `
                -ImportedMatrix $import `
                -DefaultAcl @{} `
                -AdGroupPlaceHolders @() `
                -Throttle 1

            # Ensure pipeline executed
            Should -Invoke Test-MatrixFileHC -Times 1
            Should -Invoke Test-MatrixPermissionsHC -Times 1
            Should -Invoke Test-MatrixFormDataHC -Times 1
            Should -Invoke Test-MatrixSettingHC -Times 1
            Should -Invoke Test-AdObjectsHC -Times 1
            Should -Invoke Test-ExpandedMatrixHC -Times 1
        }
    }
}
