#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

Describe 'Invoke-PermissionMatrix' {
    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        . "$root\Tests\Helpers\Helpers.HC.ps1"

        Get-ChildItem "$moduleRoot\Private" -Filter '*.ps1' -File |
        ForEach-Object { . $_.FullName }

        Get-ChildItem "$moduleRoot\Public" -Filter '*.ps1' -File -ErrorAction SilentlyContinue |
        ForEach-Object { . $_.FullName }

        # Defensive: make sure the function under test is loaded regardless of
        # whether it lives in Public, Private, or a sibling folder.
        if (-not (Get-Command Invoke-PermissionMatrix -ErrorAction SilentlyContinue)) {
            Get-ChildItem $moduleRoot -Recurse -Filter 'Invoke-PermissionMatrix.ps1' -File |
            Select-Object -First 1 |
            ForEach-Object { . $_.FullName }
        }

        #region Factory helpers
        function New-FatalCheck {
            param([string]$Name = 'TestFatal', [string]$Description = 'Test')
            [PSCustomObject]@{
                Type        = 'FatalError'
                Name        = $Name
                Description = $Description
                Value       = $null
            }
        }

        function New-Matrix {
            param([pscustomobject[]]$Check = @())
            [PSCustomObject]@{
                Check = [System.Collections.Generic.List[pscustomobject]]($Check)
            }
        }

        function New-Context {
            param(
                [bool]$FoundMatrices = $true,
                [array]$AllMatrices = @(),
                [string]$Marker
            )
            $c = [PSCustomObject]@{
                FoundMatrices = $FoundMatrices
                AllMatrices   = $AllMatrices
            }
            if ($Marker) {
                $c | Add-Member -NotePropertyName Marker -NotePropertyValue $Marker
            }
            $c
        }
        #endregion
    }

    BeforeEach {
        $script:systemErrors = [System.Collections.Generic.List[pscustomobject]]::new()

        $script:configFile = 'TestDrive:\config.json'
        $script:scriptPath = @{
            TestRequirementsFile   = 'TestDrive:\TestReq.ps1'
            SetPermissionFile      = 'TestDrive:\SetPerm.ps1'
            UpdateServiceNow       = 'TestDrive:\UpdateSnow.ps1'
            PermissionMatrixModule = 'TestDrive:\mod.psm1'
        }

        # Default: a healthy run with matrices found and no fatal errors.
        $script:beginContext = New-Context -FoundMatrices $true -AllMatrices @( New-Matrix )

        Mock Invoke-PermissionMatrixBeginHC { $script:beginContext }

        # Faithful stand-in: a check list is "fatal" when any item is a
        # FatalError. This is exercised twice by the orchestrator - once on the
        # system errors and once per matrix - so the mock must reflect its input.
        Mock Test-ItemHasFatalErrorHC {
            [bool](@($CheckList | Where-Object { $_.Type -eq 'FatalError' }).Count)
        }

        # By default the Process stage passes the context straight through.
        Mock Invoke-PermissionMatrixProcessHC { $Context }
        Mock Invoke-PermissionMatrixEndHC { }
        Mock Add-ErrorHC { }
        Mock Write-Error { }

        # The event-log fallback cmdlets do not exist on every edition/host
        # (e.g. they are absent in PowerShell 7). Mock them only when present so
        # the test neither fails to set up the mock nor writes to a real log.
        if (Get-Command New-EventLog -ErrorAction SilentlyContinue) { Mock New-EventLog { } }
        if (Get-Command Write-EventLog -ErrorAction SilentlyContinue) { Mock Write-EventLog { } }
    }

    Context 'Stage routing' {
        It 'runs Begin, Process, then End when matrices are found and there is no fatal error' {
            Invoke-PermissionMatrix `
                -ConfigurationJsonFile $configFile `
                -ScriptPath $scriptPath `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Invoke-PermissionMatrixBeginHC -Times 1 -Exactly
            Should -Invoke Invoke-PermissionMatrixProcessHC -Times 1 -Exactly
            Should -Invoke Invoke-PermissionMatrixEndHC -Times 1 -Exactly
            Should -Invoke Add-ErrorHC -Times 0 -Exactly
        }

        It 'passes the Process stage output on to the End stage' {
            # Process returns a *new* context; End must receive that one, not the
            # original from Begin.
            Mock Invoke-PermissionMatrixProcessHC {
                New-Context -FoundMatrices $true -Marker 'processed'
            }

            Invoke-PermissionMatrix `
                -ConfigurationJsonFile $configFile `
                -ScriptPath $scriptPath `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Invoke-PermissionMatrixEndHC -Times 1 -Exactly -ParameterFilter {
                $Context.Marker -eq 'processed'
            }
        }

        It 'skips the Process stage when no matrices were found but still runs End' {
            $script:beginContext = New-Context -FoundMatrices $false -AllMatrices @( New-Matrix )

            Invoke-PermissionMatrix `
                -ConfigurationJsonFile $configFile `
                -ScriptPath $scriptPath `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Invoke-PermissionMatrixProcessHC -Times 0 -Exactly
            Should -Invoke Invoke-PermissionMatrixEndHC -Times 1 -Exactly
        }
    }

    Context 'System-level fatal error' {
        It 'skips the Process stage and flags un-flagged matrices as "Run aborted"' {
            $m = New-Matrix
            $script:beginContext = New-Context -FoundMatrices $true -AllMatrices @($m)
            $systemErrors.Add((New-FatalCheck -Name 'System boom'))

            Invoke-PermissionMatrix `
                -ConfigurationJsonFile $configFile `
                -ScriptPath $scriptPath `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Invoke-PermissionMatrixProcessHC -Times 0 -Exactly

            $aborted = $m.Check.Where({ $_.Name -eq 'Run aborted' })
            $aborted.Count | Should -Be 1
            $aborted[0].Type | Should -Be 'FatalError'
        }

        It 'does not add a second fatal error to a matrix that already has one' {
            $m = New-Matrix -Check @( (New-FatalCheck -Name 'Pre-existing') )
            $script:beginContext = New-Context -FoundMatrices $true -AllMatrices @($m)
            $systemErrors.Add((New-FatalCheck -Name 'System boom'))

            Invoke-PermissionMatrix `
                -ConfigurationJsonFile $configFile `
                -ScriptPath $scriptPath `
                -SystemErrors ([ref]$systemErrors)

            $m.Check.Where({ $_.Name -eq 'Run aborted' }).Count | Should -Be 0
            $m.Check.Count | Should -Be 1
        }

        It 'still runs the End stage after a system-level fatal error' {
            $script:beginContext = New-Context -FoundMatrices $true -AllMatrices @( New-Matrix )
            $systemErrors.Add((New-FatalCheck -Name 'System boom'))

            Invoke-PermissionMatrix `
                -ConfigurationJsonFile $configFile `
                -ScriptPath $scriptPath `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Invoke-PermissionMatrixEndHC -Times 1 -Exactly
        }
    }

    Context 'Error handling' {
        It 'records an orchestrator failure when the Begin stage throws' {
            Mock Invoke-PermissionMatrixBeginHC { throw 'init failed' }

            Invoke-PermissionMatrix `
                -ConfigurationJsonFile $configFile `
                -ScriptPath $scriptPath `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Add-ErrorHC -Times 1 -Exactly -ParameterFilter {
                $Name -eq 'Unhandled orchestrator failure'
            }
            # No context was created, so the End stage must not run.
            Should -Invoke Invoke-PermissionMatrixEndHC -Times 0 -Exactly
        }

        It 'still runs the End stage when the Process stage throws (best-effort reporting)' {
            Mock Invoke-PermissionMatrixProcessHC { throw 'process boom' }

            Invoke-PermissionMatrix `
                -ConfigurationJsonFile $configFile `
                -ScriptPath $scriptPath `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Add-ErrorHC -Times 1 -Exactly
            # The context from Begin survives, so End still runs.
            Should -Invoke Invoke-PermissionMatrixEndHC -Times 1 -Exactly
        }

        It 'does not run the End stage when no context was created' {
            $script:beginContext = $null

            Invoke-PermissionMatrix `
                -ConfigurationJsonFile $configFile `
                -ScriptPath $scriptPath `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Invoke-PermissionMatrixEndHC -Times 0 -Exactly
        }
    }

    Context 'Event-log fallback' {
        It 'writes each system error to the error stream when no context exists' {
            $script:beginContext = $null
            $systemErrors.Add((New-FatalCheck -Name 'Config missing'))

            Invoke-PermissionMatrix `
                -ConfigurationJsonFile $configFile `
                -ScriptPath $scriptPath `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Write-Error -Times 1 -Exactly
            Should -Invoke Invoke-PermissionMatrixEndHC -Times 0 -Exactly
        }

        It 'reports every system error, not just the first' {
            $script:beginContext = $null
            $systemErrors.Add((New-FatalCheck -Name 'First'))
            $systemErrors.Add((New-FatalCheck -Name 'Second'))

            Invoke-PermissionMatrix `
                -ConfigurationJsonFile $configFile `
                -ScriptPath $scriptPath `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Write-Error -Times 2 -Exactly
        }
    }
}