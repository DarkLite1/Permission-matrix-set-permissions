#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

Describe 'Invoke-PermissionMatrixProcessHC' {
    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        . "$root\Tests\Helpers\Helpers.HC.ps1"
        . "$moduleRoot\Private\Utils.ps1"
        . "$moduleRoot\Private\Invoke-PermissionMatrixProcessHC.ps1"
        . "$moduleRoot\Private\Invoke-WithOptionalParallelismHC.ps1"
        . "$moduleRoot\Private\Validation.ps1"

        function New-TestContext {
            param(
                [array]$Matrices = @(),
                [array]$FileResults = $null,
                [hashtable]$ScriptPath = @{
                    TestRequirements = 'TestDrive:\TestReq.ps1'
                    SetPermissions   = 'TestDrive:\SetPerm.ps1'
                },
                [hashtable]$MaxConcurrent = @{
                    Computers        = 10
                    FoldersPerMatrix = 3
                },
                [string]$PSSessionConfiguration = 'PowerShell.7',
                [bool]$Detailed = $false
            )

            # Default FileResults = one file containing all matrices, no file-level errors
            if ($null -eq $FileResults) {
                $FileResults = @(
                    [PSCustomObject]@{
                        Check    = [System.Collections.Generic.List[pscustomobject]]::new()
                        Matrices = $Matrices
                    }
                )
            }

            [PSCustomObject]@{
                AllMatrices = $Matrices
                FileResults = $FileResults
                ScriptPath  = $ScriptPath
                Config      = [PSCustomObject]@{
                    MaxConcurrent = $MaxConcurrent
                    Settings      = [PSCustomObject]@{
                        PSSessionConfiguration = $PSSessionConfiguration
                        SaveLogFiles           = [PSCustomObject]@{ Detailed = $Detailed }
                    }
                }
            }
        }

        function New-TestMatrix {
            param(
                [string]$ComputerName = 'SERVER1',
                [string]$Path = 'C:\Data',
                [string]$Action = 'Fix',
                [string]$ID = ([guid]::NewGuid().ToString()),
                [pscustomobject[]]$Check = @(),
                [array]$Matrix = @()
            )

            [PSCustomObject]@{
                ID      = $ID
                Setting = [PSCustomObject]@{
                    Formatted = [PSCustomObject]@{
                        ComputerName = $ComputerName
                        Path         = $Path
                        Action       = $Action
                    }
                }
                Check   = [System.Collections.Generic.List[pscustomobject]](
                    [System.Collections.Generic.List[pscustomobject]]::new()
                )
                Matrix  = $Matrix
                JobTime = @{}
            }
        }

        function New-FatalCheck {
            param([string]$Name = 'TestFatal', [string]$Description = 'Test')
            [PSCustomObject]@{
                Type        = 'FatalError'
                Name        = $Name
                Description = $Description
                Value       = $null
            }
        }
    }

    BeforeEach {
        $script:systemErrors = [System.Collections.Generic.List[pscustomobject]]::new()

        Mock Invoke-WithOptionalParallelismHC {
            param($InputObject, $ScriptBlock, $ThrottleLimit, $ArgumentList = @())
            $results = foreach ($item in $InputObject) {
                & $ScriptBlock $item @ArgumentList
            }
            @($results)
        }

        Mock Invoke-Command { return $null }

        Mock ConvertTo-StructuredObjectHC {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory, ValueFromPipeline = $true)]
                $InputObject
            )
            process {
                foreach ($obj in $InputObject) {
                    if ($null -ne $obj) { $obj }
                }
            }
        }
    }

    Context 'Guard conditions' {
        It 'returns immediately when AllMatrices is empty' {
            $ctx = New-TestContext -Matrices @()

            $result = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            $result | Should -Be $ctx
            Should -Invoke Invoke-Command -Times 0
            Should -Invoke Invoke-WithOptionalParallelismHC -Times 0
        }

        It 'returns immediately when AllMatrices is $null' {
            $ctx = New-TestContext
            $ctx.AllMatrices = $null

            $result = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            $result | Should -Be $ctx
            Should -Invoke Invoke-Command -Times 0
        }

        It 'skips matrices belonging to a file with file-level FatalError' {
            $m = New-TestMatrix
            $fileResults = @(
                [PSCustomObject]@{
                    Check    = [System.Collections.Generic.List[pscustomobject]]@((New-FatalCheck))
                    Matrices = @($m)
                }
            )
            $ctx = New-TestContext -Matrices @($m) -FileResults $fileResults

            $null = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            # No remote work attempted because the only executable group is empty
            Should -Invoke Invoke-Command -Times 0
        }

        It 'skips matrices that already have FatalError in their own Check list' {
            $broken = New-TestMatrix
            $broken.Check.Add((New-FatalCheck))

            $ctx = New-TestContext -Matrices @($broken)

            $null = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Invoke-Command -Times 0
        }
    }

    Context 'Test Requirements phase' {
        It 'groups matrices by ComputerName and calls Invoke-Command once per computer' {
            $m1 = New-TestMatrix -ComputerName 'SERVER1' -Path 'C:\A'
            $m2 = New-TestMatrix -ComputerName 'SERVER1' -Path 'C:\B'
            $m3 = New-TestMatrix -ComputerName 'SERVER2' -Path 'C:\C'

            $ctx = New-TestContext -Matrices @($m1, $m2, $m3)

            $null = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            # One Invoke-Command per computer for requirements (2 total here)
            Should -Invoke Invoke-Command -Times 2 -ParameterFilter {
                $FilePath -eq 'TestDrive:\TestReq.ps1'
            }
        }

        It 'aggregates all paths for matrices on the same computer' {
            $m1 = New-TestMatrix -ComputerName 'SERVER1' -Path 'C:\A'
            $m2 = New-TestMatrix -ComputerName 'SERVER1' -Path 'C:\B'
            $ctx = New-TestContext -Matrices @($m1, $m2)

            $null = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Invoke-Command -Times 1 -ParameterFilter {
                $FilePath -eq 'TestDrive:\TestReq.ps1' -and
                $ArgumentList[0] -contains 'C:\A' -and
                $ArgumentList[0] -contains 'C:\B'
            }
        }

        It 'uses the configured PSSessionConfiguration' {
            $m = New-TestMatrix
            $ctx = New-TestContext -Matrices @($m) -PSSessionConfiguration 'CustomConfig'

            $null = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Invoke-Command -ParameterFilter {
                $ConfigurationName -eq 'CustomConfig'
            }
        }

        It 'defaults PSSessionConfiguration to PowerShell.7 when not set' {
            $m = New-TestMatrix
            $ctx = New-TestContext -Matrices @($m) -PSSessionConfiguration $null

            $null = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Invoke-Command -ParameterFilter {
                $ConfigurationName -eq 'PowerShell.7'
            }
        }

        It 'appends requirement errors to all matrices on the failing computer' {
            $m1 = New-TestMatrix -ComputerName 'SERVER1' -Path 'C:\A'
            $m2 = New-TestMatrix -ComputerName 'SERVER1' -Path 'C:\B'
            $ctx = New-TestContext -Matrices @($m1, $m2)

            # First Invoke-Command (requirements) throws; subsequent ones don't matter
            Mock Invoke-Command { throw 'unreachable host' } -ParameterFilter {
                $FilePath -eq 'TestDrive:\TestReq.ps1'
            }

            $null = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            $m1.Check.Count | Should -BeGreaterThan 0
            $m2.Check.Count | Should -BeGreaterThan 0
            $m1.Check[0].Type | Should -Be 'FatalError'
            $m1.Check[0].Name | Should -Be 'Computer requirements'
        }

        It 'excludes a matrix from Set Permissions phase if requirements added FatalError' {
            $m = New-TestMatrix
            $ctx = New-TestContext -Matrices @($m)

            Mock Invoke-Command { throw 'fail' } -ParameterFilter {
                $FilePath -eq 'TestDrive:\TestReq.ps1'
            }

            $null = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Invoke-Command -Times 0 -ParameterFilter {
                $FilePath -eq 'TestDrive:\SetPerm.ps1'
            }
        }
    }

    Context 'Set Permissions phase' {
        It 'calls SetPermissions for matrices that passed requirements with all expected arguments' {
            $matrixContent = @(
                [PSCustomObject]@{ Path = 'C:\Data\Sub1'; ACL = @{ 'user1' = 'R' } }
                [PSCustomObject]@{ Path = 'C:\Data\Sub2'; ACL = @{ 'user2' = 'M' } }
            )
            $m = New-TestMatrix `
                -ComputerName 'SERVER1' `
                -Path 'C:\Data' `
                -Action 'Fix' `
                -Matrix $matrixContent

            $ctx = New-TestContext `
                -Matrices @($m) `
                -MaxConcurrent @{ Computers = 10; FoldersPerMatrix = 5 } `
                -Detailed $true

            $null = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Invoke-Command -Times 1 -ParameterFilter {
                $FilePath -eq 'TestDrive:\SetPerm.ps1' -and
                $ComputerName -eq 'SERVER1' -and
                # Positional argument order matches Set_permissions.ps1's param block:
                # 0=Path, 1=Action, 2=Matrix, 3=JobThrottleLimit, 4=DetailedLog
                $ArgumentList[0] -eq 'C:\Data' -and
                $ArgumentList[1] -eq 'Fix' -and
                $ArgumentList[2].Count -eq 2 -and
                $ArgumentList[2][0].Path -eq 'C:\Data\Sub1' -and
                $ArgumentList[2][1].Path -eq 'C:\Data\Sub2' -and
                $ArgumentList[3] -eq 5 -and
                $ArgumentList[4] -eq $true
            }
        }

        It 'serializes the matrix as JSON for transport across the runspace boundary' {
            $matrixContent = @(
                [PSCustomObject]@{ Path = 'C:\Data'; ACL = @{ 'user1' = 'R' } }
            )
            $m = New-TestMatrix -Matrix $matrixContent
            $ctx = New-TestContext -Matrices @($m)

            $null = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            # The ArgumentList passed to Invoke-Command should include the
            # deserialized matrix, not raw JSON — deserialization happens
            # inside the scriptblock before Invoke-Command runs.
            Should -Invoke Invoke-Command -ParameterFilter {
                $FilePath -eq 'TestDrive:\SetPerm.ps1' -and
                $ArgumentList[2] -ne $null
            }
        }

        It 'passes MaxConcurrent.FoldersPerMatrix to SetPermissions' {
            $m = New-TestMatrix
            $ctx = New-TestContext -Matrices @($m) -MaxConcurrent @{
                Computers        = 5
                FoldersPerMatrix = 7
            }

            $null = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Invoke-Command -ParameterFilter {
                $FilePath -eq 'TestDrive:\SetPerm.ps1' -and
                $ArgumentList[3] -eq 7
            }
        }

        It 'passes DetailedLog flag to SetPermissions' {
            $m = New-TestMatrix
            $ctx = New-TestContext -Matrices @($m) -Detailed $true

            $null = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Invoke-Command -ParameterFilter {
                $FilePath -eq 'TestDrive:\SetPerm.ps1' -and
                $ArgumentList[4] -eq $true
            }
        }

        It 'records JobTime.Start, End, and Duration on the matching matrix' {
            $m = New-TestMatrix
            $ctx = New-TestContext -Matrices @($m)

            $null = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            $m.JobTime.Start | Should -Not -BeNullOrEmpty
            $m.JobTime.End | Should -Not -BeNullOrEmpty
            $m.JobTime.Duration | Should -BeOfType [TimeSpan]
        }

        It 'appends SetPermissions errors to the originating matrix only' {
            $m1 = New-TestMatrix -ComputerName 'SERVER1' -Path 'C:\A'
            $m2 = New-TestMatrix -ComputerName 'SERVER2' -Path 'C:\B'
            $ctx = New-TestContext -Matrices @($m1, $m2)

            # Only the SERVER1 set-permissions call fails
            Mock Invoke-Command {
                throw 'set permissions failed'
            } -ParameterFilter {
                $FilePath -eq 'TestDrive:\SetPerm.ps1' -and $ComputerName -eq 'SERVER1'
            }

            $null = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            $m1.Check.Where({ $_.Type -eq 'FatalError' -and $_.Name -eq 'Set permissions' }).Count |
            Should -Be 1
            $m2.Check.Where({ $_.Type -eq 'FatalError' -and $_.Name -eq 'Set permissions' }).Count |
            Should -Be 0
        } -Tag test

        It 'returns context untouched if all matrices failed requirements' {
            $m = New-TestMatrix
            $ctx = New-TestContext -Matrices @($m)

            Mock Invoke-Command { throw 'unreachable' } -ParameterFilter {
                $FilePath -eq 'TestDrive:\TestReq.ps1'
            }

            $result = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            $result | Should -Be $ctx
            Should -Invoke Invoke-Command -Times 0 -ParameterFilter {
                $FilePath -eq 'TestDrive:\SetPerm.ps1'
            }
        }
    }

    Context 'Outer try/catch' {
        It 'adds a FatalError to SystemErrors when an unhandled exception occurs' {
            $m = New-TestMatrix
            $ctx = New-TestContext -Matrices @($m)

            # Force a throw from the helper to simulate an unhandled internal failure
            Mock Invoke-WithOptionalParallelismHC { throw 'catastrophic' }

            $result = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            $result | Should -Be $ctx
            $systemErrors.Count | Should -BeGreaterThan 0
            $systemErrors[0].Type | Should -Be 'FatalError'
            $systemErrors[0].Name | Should -Be 'PROCESS stage failure'
        }
    }
}