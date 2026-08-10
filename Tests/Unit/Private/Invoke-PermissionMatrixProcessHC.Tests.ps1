#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

Describe 'Invoke-PermissionMatrixProcessHC' {
    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        . "$root\Tests\Helpers\Helpers.HC.ps1"
        
        Get-ChildItem "$moduleRoot\Private" -Filter '*.ps1' -File |
        ForEach-Object { . $_.FullName }

        function New-TestContext {
            param(
                [array]$Matrices = @(),
                [array]$FileResults = $null,
                [hashtable]$ScriptPath = @{
                    TestRequirements = 'TestDrive:\TestReq.ps1'
                    SetPermissions   = 'TestDrive:\SetPerm.ps1'
                },
                [hashtable]$MaxConcurrent = @{
                    JobsTotal        = 10
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
                    MaxConcurrent          = $MaxConcurrent
                    PSSessionConfiguration = $PSSessionConfiguration
                    Settings               = [PSCustomObject]@{
                        SaveLogFiles = [PSCustomObject]@{ Detailed = $Detailed }
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

        # SetPermissions phase now owns an explicit PSSession per job so it can
        # forcibly close (terminate) an orphaned remote command before retrying.
        # Return a real PSSession-typed mock object so the strongly-typed
        # -Session parameter on Invoke-Command / Remove-PSSession binds.
        Mock New-PSSession {
            New-MockObject -Type 'System.Management.Automation.Runspaces.PSSession'
        }
        Mock Remove-PSSession { }

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

            $result | Should-Be $ctx
            Should-Invoke Invoke-Command -Times 0
            Should-Invoke Invoke-WithOptionalParallelismHC -Times 0
        }

        It 'returns immediately when AllMatrices is $null' {
            $ctx = New-TestContext
            $ctx.AllMatrices = $null

            $result = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            $result | Should-Be $ctx
            Should-Invoke Invoke-Command -Times 0
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
            Should-Invoke Invoke-Command -Times 0
        }

        It 'skips matrices that already have FatalError in their own Check list' {
            $broken = New-TestMatrix
            $broken.Check.Add((New-FatalCheck))

            $ctx = New-TestContext -Matrices @($broken)

            $null = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            Should-Invoke Invoke-Command -Times 0
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
            Should-Invoke Invoke-Command -Times 2 -ParameterFilter {
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

            Should-Invoke Invoke-Command -Times 1 -ParameterFilter {
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

            Should-Invoke Invoke-Command -ParameterFilter {
                $ConfigurationName -eq 'CustomConfig'
            }
        }

        It 'defaults PSSessionConfiguration to PowerShell.7 when not set' {
            $m = New-TestMatrix
            $ctx = New-TestContext -Matrices @($m) -PSSessionConfiguration $null

            $null = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            Should-Invoke Invoke-Command -ParameterFilter {
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

            $m1.Check.Count | Should-BeGreaterThan 0
            $m2.Check.Count | Should-BeGreaterThan 0
            $m1.Check[0].Type | Should-Be 'FatalError'
            $m1.Check[0].Name | Should-Be 'Computer requirements'
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

            Should-Invoke Invoke-Command -Times 0 -ParameterFilter {
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
                -MaxConcurrent @{ JobsTotal = 10; FoldersPerMatrix = 5 } `
                -Detailed $true

            $null = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            Should-Invoke Invoke-Command -Times 1 -ParameterFilter {
                $FilePath -eq 'TestDrive:\SetPerm.ps1' -and
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

            # ComputerName is now supplied to the session, not to Invoke-Command
            Should-Invoke New-PSSession -Times 1 -ParameterFilter {
                $ComputerName -eq 'SERVER1'
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
            Should-Invoke Invoke-Command -ParameterFilter {
                $FilePath -eq 'TestDrive:\SetPerm.ps1' -and
                $ArgumentList[2] -ne $null
            }
        }

        It 'passes MaxConcurrent.FoldersPerMatrix to SetPermissions' {
            $m = New-TestMatrix
            $ctx = New-TestContext -Matrices @($m) -MaxConcurrent @{
                JobsTotal        = 5
                FoldersPerMatrix = 7
            }

            $null = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            Should-Invoke Invoke-Command -ParameterFilter {
                $FilePath -eq 'TestDrive:\SetPerm.ps1' -and
                $ArgumentList[3] -eq 7
            }
        }

        It 'raises a FoldersPerMatrix of zero to one before sending it on' {
            <# Defence in depth: validation rejects zero, but the value lands on
            ForEach-Object -ThrottleLimit inside SetPermissions.ps1, which
            rejects it. JobsTotal and JobsPerComputer were already guarded with
            [math]::Max(1, ...); this one was passed through raw. #>
            $m = New-TestMatrix
            $ctx = New-TestContext -Matrices @($m) -MaxConcurrent @{
                JobsTotal        = 5
                FoldersPerMatrix = 0
            }

            $null = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            Should-Invoke Invoke-Command -ParameterFilter {
                $FilePath -eq 'TestDrive:\SetPerm.ps1' -and
                $ArgumentList[3] -eq 1
            }
        }

        It 'passes DetailedLog flag to SetPermissions' {
            $m = New-TestMatrix
            $ctx = New-TestContext -Matrices @($m) -Detailed $true

            $null = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            Should-Invoke Invoke-Command -ParameterFilter {
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

            $m.JobTime.Start | Should-BeTruthy
            $m.JobTime.End | Should-BeTruthy
            $m.JobTime.Duration | Should-HaveType ([TimeSpan])
        }

        It 'appends SetPermissions errors to the originating matrix only' {
            $m1 = New-TestMatrix -ComputerName 'SERVER1' -Path 'C:\A'
            $m2 = New-TestMatrix -ComputerName 'SERVER2' -Path 'C:\B'
            $ctx = New-TestContext -Matrices @($m1, $m2)

            # Only the SERVER1 set-permissions call fails. The session is opaque
            # here, so distinguish computers by the job Path in ArgumentList[0]
            # (SERVER1 -> C:\A, SERVER2 -> C:\B).
            Mock Invoke-Command {
                throw 'set permissions failed'
            } -ParameterFilter {
                $FilePath -eq 'TestDrive:\SetPerm.ps1' -and $ArgumentList[0] -eq 'C:\A'
            }

            $null = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            $m1.Check.Where({ $_.Type -eq 'FatalError' -and $_.Name -eq 'Set permissions' }).Count | Should-Be 1
            $m2.Check.Where({ $_.Type -eq 'FatalError' -and $_.Name -eq 'Set permissions' }).Count | Should-Be 0
        } -Tag test

        It 'retries a transient WinRM I/O abort and succeeds without recording an error' {
            $m = New-TestMatrix -ComputerName 'SERVER1' -Path 'C:\A'
            $ctx = New-TestContext -Matrices @($m)

            Mock Start-Sleep {}

            $script:setPermAttempt = 0
            Mock Invoke-Command {
                $script:setPermAttempt++
                if ($script:setPermAttempt -lt 3) {
                    throw 'Processing data from remote server SERVER1 failed with the following error message: The I/O operation has been aborted because of either a thread exit or an application request.'
                }
                return $null
            } -ParameterFilter { $FilePath -eq 'TestDrive:\SetPerm.ps1' }

            $null = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            # Two transient aborts + one success = 3 attempts, 2 back-off waits
            Should-Invoke Invoke-Command -Times 3 -ParameterFilter {
                $FilePath -eq 'TestDrive:\SetPerm.ps1'
            }
            Should-Invoke Start-Sleep -Times 2
            # Every attempt opens then closes its own session, so an orphaned
            # remote command is terminated before the next attempt reconnects.
            Should-Invoke New-PSSession -Times 3
            Should-Invoke Remove-PSSession -Times 3
            $m.Check.Where({ $_.Type -eq 'FatalError' -and $_.Name -eq 'Set permissions' }).Count | Should-Be 0
        }

        It 'gives up after the maximum attempts on a persistent I/O abort and records one error' {
            $m = New-TestMatrix -ComputerName 'SERVER1' -Path 'C:\A'
            $ctx = New-TestContext -Matrices @($m)

            Mock Start-Sleep {}

            Mock Invoke-Command {
                throw 'The I/O operation has been aborted because of either a thread exit or an application request.'
            } -ParameterFilter { $FilePath -eq 'TestDrive:\SetPerm.ps1' }

            $null = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            # 3 attempts total (initial + 2 retries), 2 back-off waits
            Should-Invoke Invoke-Command -Times 3 -ParameterFilter {
                $FilePath -eq 'TestDrive:\SetPerm.ps1'
            }
            Should-Invoke Start-Sleep -Times 2
            # Session closed before each retry to terminate any orphaned run
            Should-Invoke Remove-PSSession -Times 3
            $m.Check.Where({ $_.Type -eq 'FatalError' -and $_.Name -eq 'Set permissions' }).Count | Should-Be 1
        }

        It 'does NOT retry a genuine (non-transient) SetPermissions failure' {
            $m = New-TestMatrix -ComputerName 'SERVER1' -Path 'C:\A'
            $ctx = New-TestContext -Matrices @($m)

            Mock Start-Sleep {}

            Mock Invoke-Command {
                throw 'Access to the path is denied.'
            } -ParameterFilter { $FilePath -eq 'TestDrive:\SetPerm.ps1' }

            $null = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            # Business error: single attempt, no retries, no waits
            Should-Invoke Invoke-Command -Times 1 -ParameterFilter {
                $FilePath -eq 'TestDrive:\SetPerm.ps1'
            }
            Should-Invoke Start-Sleep -Times 0
            # Single session, still closed exactly once
            Should-Invoke New-PSSession -Times 1
            Should-Invoke Remove-PSSession -Times 1
            $m.Check.Where({ $_.Type -eq 'FatalError' -and $_.Name -eq 'Set permissions' }).Count | Should-Be 1
        }

        It 'returns context untouched if all matrices failed requirements' {
            $m = New-TestMatrix
            $ctx = New-TestContext -Matrices @($m)

            Mock Invoke-Command { throw 'unreachable' } -ParameterFilter {
                $FilePath -eq 'TestDrive:\TestReq.ps1'
            }

            $result = Invoke-PermissionMatrixProcessHC `
                -Context $ctx `
                -SystemErrors ([ref]$systemErrors)

            $result | Should-Be $ctx
            Should-Invoke Invoke-Command -Times 0 -ParameterFilter {
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

            $result | Should-Be $ctx
            $systemErrors.Count | Should-BeGreaterThan 0
            $systemErrors[0].Type | Should-Be 'FatalError'
            $systemErrors[0].Name | Should-Be 'PROCESS stage failure'
        }
    }
}