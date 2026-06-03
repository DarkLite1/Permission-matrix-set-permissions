#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

# =============================================================================
# End-to-end tests for the Permission Matrix pipeline - NON-HAPPY paths.
#
# Companion to EndToEndHappyPath.Tests.ps1. Where that file proves the pipeline
# applies permissions correctly, this file proves it FAILS SAFELY:
#
#   - a fatal input error skips the PROCESS stage entirely, so no permission is
#     ever applied;
#   - a non-catastrophic failure still runs the END stage, so the notification
#     mail is sent (carrying the error report);
#   - a catastrophic failure (no execution context could be built) sends no
#     mail and falls back to the Event Log;
#   - an unexpected exception inside a stage is caught, recorded as a 'Runtime'
#     fatal, and does NOT crash the orchestrator;
#   - a failure resolving AD objects is downgraded to a warning, not a fatal;
#   - a failing SMTP send is caught and recorded, not propagated out of END.
#
# These assert PIPELINE OUTCOMES, not validation messages - the message-level
# checks already live in InputValidation.Tests.ps1 and
# MatrixValidation.Tests.ps1, which drive the entrypoint and match on log/HTML
# patterns. This file invokes Invoke-PermissionMatrix directly (like the
# happy-path test) and inspects SystemErrors + mock invocation counts.
#
# NO PREREQUISITES. Unlike the happy-path test, nothing here applies a real
# ACL, so there is no Administrator / PSRemoting / 'PowerShell.7' requirement.
# Every boundary that would touch the real world (AD, SMTP, remoting, Event
# Log) is mocked, and the permission-applying stage is either skipped by design
# or mocked. The whole file runs on any machine, unskipped.
# =============================================================================

Describe 'Permission Matrix - End to End (non-happy paths)' {

    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        # Order matches EndToEndHappyPath.Tests.ps1: Fixtures.Excel re-defines
        # New-ValidDefaultsExcelFixture, so it must be dot-sourced last.
        . "$root\Tests\Helpers\Fixtures.Json.ps1"
        . "$root\Tests\Helpers\Fixtures.Excel.ps1"

        Import-Module "$moduleRoot\PermissionMatrix.psm1" -Force

        # ScriptPath hashtable shape mirrors EndToEndHappyPath.Tests.ps1 - the
        # keys the entry point actually reads.
        $script:ScriptPath = @{
            PermissionMatrixModule = "$moduleRoot\PermissionMatrix.psm1"
            SetPermissions         = "$root\Scripts\Operations\SetPermissions.ps1"
            TestRequirements       = "$root\Scripts\Operations\TestRequirements.ps1"
            UpdateServiceNow       = "$root\Scripts\Operations\UpdateServiceNow.ps1"
        }

        # Serialise a config hashtable (from the JSON fixtures) to a JSON file,
        # the same way the happy-path test does.
        function Save-Config {
            param([hashtable]$Config, [string]$Path)
            $Config |
            ConvertTo-Json -Depth 20 |
            Out-File -LiteralPath $Path -Encoding utf8 -Force
        }
    }

    BeforeEach {
        # Hermetic safety net. None of the tests below should reach remoting or
        # the real Event Log, but mocking these guarantees a stray call can't
        # touch the host or fail on a non-admin machine. Write-Error is mocked
        # so the orchestrator's fallback logging can't turn into a terminating
        # error under ErrorActionPreference='Stop' and defeat -Not -Throw.
        Mock Invoke-Command -ModuleName PermissionMatrix { }
        Mock Write-EventLog -ModuleName PermissionMatrix { }
        Mock New-EventLog -ModuleName PermissionMatrix { }
        Mock Write-Error -ModuleName PermissionMatrix { }
    }

    It 'sends no mail and records errors when the config file is missing (catastrophic)' {
        # A missing config file means no execution context can be built, so the
        # orchestrator's finally block takes the Event-Log fallback path rather
        # than the END (mail) path.
        $missingConfig = Join-Path $TestDrive 'does-not-exist.json'

        Mock Send-MailKitMessageHC -ModuleName PermissionMatrix { }

        $systemErrors = [System.Collections.Generic.List[object]]::new()

        { Invoke-PermissionMatrix `
                -ConfigurationJsonFile $missingConfig `
                -ScriptPath $ScriptPath `
                -SystemErrors ([ref]$systemErrors) } |
        Should -Not -Throw -Because 'the orchestrator records the failure instead of propagating it'

        $systemErrors.Count |
        Should -BeGreaterThan 0 -Because 'a missing config file must be recorded as an error'

        Should -Invoke Send-MailKitMessageHC -ModuleName PermissionMatrix -Times 0 -Exactly -Because (
            'with no execution context the END stage never runs, so no mail is sent'
        )
    }

    It 'skips all permission work but still sends mail when input validation fails' {
        # A parseable config with an invalid Matrix.FolderPath: the context IS
        # built (so END runs and mail is sent), a fatal is recorded, and the
        # PROCESS stage is gated off - proving bad input never applies an ACL.
        $logsDir = (New-Item 'TestDrive:\Logs-validation' -ItemType Directory -Force).FullName

        $config = New-JsonFixtureWithModifiedValue -Path 'Matrix.FolderPath' -Value 'X:\does-not-exist'
        $config.Matrix.DefaultsFile = New-ValidDefaultsExcelFixture -Path (Join-Path $TestDrive 'Defaults-validation.xlsx')
        $config.Settings.SaveLogFiles.Where.Folder = $logsDir
        $config.MaxConcurrent.FoldersPerMatrix = 1

        $configPath = Join-Path $TestDrive 'Input-validation.json'
        Save-Config -Config $config -Path $configPath

        Mock Get-ADObjectDetailHC -ModuleName PermissionMatrix { @() }
        Mock Send-MailKitMessageHC -ModuleName PermissionMatrix { }
        # Mocked purely so we can prove it is NEVER invoked on a fatal input.
        Mock Invoke-PermissionMatrixProcessHC -ModuleName PermissionMatrix { }

        $systemErrors = [System.Collections.Generic.List[object]]::new()

        Invoke-PermissionMatrix `
            -ConfigurationJsonFile $configPath `
            -ScriptPath $ScriptPath `
            -SystemErrors ([ref]$systemErrors)

        ($systemErrors.Where({ $_.Type -eq 'FatalError' })).Count |
        Should -BeGreaterThan 0 -Because 'an invalid Matrix.FolderPath is a fatal input error'

        Should -Invoke Invoke-PermissionMatrixProcessHC -ModuleName PermissionMatrix -Times 0 -Exactly -Because (
            'a fatal input error must skip the PROCESS stage so no permission is applied'
        )

        Should -Invoke Send-MailKitMessageHC -ModuleName PermissionMatrix -Times 1 -Exactly -Because (
            'the END stage still runs on a non-catastrophic failure, so the error report is mailed'
        )
    }

    It 'records a Runtime fatal and still sends mail when a stage throws unexpectedly' {
        # Valid config + valid matrix so the BEGIN stage succeeds and a context
        # with matrices exists. PROCESS is then forced to throw; the orchestrator
        # must catch it, record a 'Runtime' fatal, run END (mail) from its
        # finally block, and not crash.
        $matrixDir = (New-Item 'TestDrive:\Matrix-rt' -ItemType Directory -Force).FullName
        $logsDir = (New-Item 'TestDrive:\Logs-rt' -ItemType Directory -Force).FullName

        $defaultsPath = Join-Path $matrixDir 'Defaults.xlsx'
        New-ValidDefaultsExcelFixture -Path $defaultsPath | Out-Null

        $matrixPath = Join-Path $matrixDir 'TeamA.xlsx'
        New-MatrixExcelFixture `
            -Path $matrixPath `
            -SettingsRows @(
            [pscustomobject]@{
                Status                  = 'Enabled'
                SiteName                = 'E2E'
                SiteCode                = 'E2E'
                ComputerName            = 'TESTHOST'      # never contacted: PROCESS is mocked
                Path                    = 'C:\DoesNotMatter'
                GroupName               = 'E2E-Test-Group'
                Action                  = 'Fix'
                ApplyDefaultPermissions = $false
            }
        )

        $config = New-JsonFixture
        $config.Matrix.FolderPath = $matrixDir
        $config.Matrix.DefaultsFile = $defaultsPath
        $config.Settings.SaveLogFiles.Where.Folder = $logsDir
        $config.MaxConcurrent.FoldersPerMatrix = 1

        $configPath = Join-Path $matrixDir 'Input.json'
        Save-Config -Config $config -Path $configPath

        # AD resolves cleanly (Bob/Mike are the 'Valid' Permissions header
        # names) so BEGIN produces a valid context with matrices. SIDs are
        # placeholders - nothing is applied because PROCESS is mocked.
        Mock Get-ADObjectDetailHC -ModuleName PermissionMatrix {
            @(
                @{ SamAccountName = 'Bob'; adObject = @{ ObjectSid = 'S-1-5-21-0-0-0-1001' } }
                @{ SamAccountName = 'Mike'; adObject = @{ ObjectSid = 'S-1-5-21-0-0-0-1002' } }
            )
        }
        Mock Send-MailKitMessageHC -ModuleName PermissionMatrix { }
        # Force an unexpected failure inside the PROCESS stage.
        Mock Invoke-PermissionMatrixProcessHC -ModuleName PermissionMatrix { throw 'boom' }

        $systemErrors = [System.Collections.Generic.List[object]]::new()

        { Invoke-PermissionMatrix `
                -ConfigurationJsonFile $configPath `
                -ScriptPath $ScriptPath `
                -SystemErrors ([ref]$systemErrors) } |
        Should -Not -Throw -Because 'the orchestrator catches stage exceptions instead of propagating them'

        ($systemErrors.Where({ $_.Type -eq 'FatalError' -and $_.Category -eq 'Runtime' })).Count |
        Should -BeGreaterThan 0 -Because 'an unhandled stage exception is recorded as a Runtime fatal'

        Should -Invoke Send-MailKitMessageHC -ModuleName PermissionMatrix -Times 1 -Exactly -Because (
            'the END stage runs in the finally block even after a stage throws'
        )
    }

    It 'downgrades an AD lookup failure to a warning and still completes' {
        # BeginHC wraps the bulk AD query, so an unreachable/erroring directory
        # becomes a 'Warning' (Name 'AD Bulk Lookup Failure') rather than a
        # fatal: the context is still returned and the run reports as normal.
        # (An AD name that merely resolves to nothing is handled separately, as
        # a per-matrix Check via Test-AdObjectInMatrixHC, not in SystemErrors -
        # that path belongs to MatrixValidation, not here.)
        $matrixDir = (New-Item 'TestDrive:\Matrix-ad' -ItemType Directory -Force).FullName
        $logsDir = (New-Item 'TestDrive:\Logs-ad' -ItemType Directory -Force).FullName

        $defaultsPath = Join-Path $matrixDir 'Defaults.xlsx'
        New-ValidDefaultsExcelFixture -Path $defaultsPath | Out-Null

        $matrixPath = Join-Path $matrixDir 'TeamA.xlsx'
        New-MatrixExcelFixture `
            -Path $matrixPath `
            -SettingsRows @(
            [pscustomobject]@{
                Status                  = 'Enabled'
                SiteName                = 'E2E'
                SiteCode                = 'E2E'
                ComputerName            = 'TESTHOST'
                Path                    = 'C:\DoesNotMatter'
                GroupName               = 'E2E-Test-Group'
                Action                  = 'Fix'
                ApplyDefaultPermissions = $false
            }
        )

        $config = New-JsonFixture
        $config.Matrix.FolderPath = $matrixDir
        $config.Matrix.DefaultsFile = $defaultsPath
        $config.Settings.SaveLogFiles.Where.Folder = $logsDir
        $config.MaxConcurrent.FoldersPerMatrix = 1

        $configPath = Join-Path $matrixDir 'Input.json'
        Save-Config -Config $config -Path $configPath

        # The directory query itself fails.
        Mock Get-ADObjectDetailHC -ModuleName PermissionMatrix { throw 'AD unreachable' }
        Mock Send-MailKitMessageHC -ModuleName PermissionMatrix { }
        # Neutralise PROCESS (this test is about BEGIN-stage AD handling, not
        # permission application). Return the context so the orchestrator's
        # `$context = Invoke-PermissionMatrixProcessHC ...` keeps it non-null and
        # the END stage still runs.
        Mock Invoke-PermissionMatrixProcessHC -ModuleName PermissionMatrix { $Context }

        $systemErrors = [System.Collections.Generic.List[object]]::new()

        { Invoke-PermissionMatrix `
                -ConfigurationJsonFile $configPath `
                -ScriptPath $ScriptPath `
                -SystemErrors ([ref]$systemErrors) } |
        Should -Not -Throw -Because 'an AD lookup failure is caught inside the BEGIN stage'

        ($systemErrors.Where({ $_.Name -eq 'AD Bulk Lookup Failure' })).Count |
        Should -BeGreaterThan 0 -Because 'a failed bulk AD lookup is recorded'

        ($systemErrors.Where({ $_.Name -eq 'AD Bulk Lookup Failure' -and $_.Type -eq 'FatalError' })).Count |
        Should -Be 0 -Because 'the AD lookup failure must be a warning, not a fatal that aborts the run'

        Should -Invoke Send-MailKitMessageHC -ModuleName PermissionMatrix -Times 1 -Exactly -Because (
            'the run continues to the END stage and mails the report after an AD warning'
        )
    }

    It 'survives a failing SMTP send by recording it as a warning' {
        # Everything succeeds except the SMTP send. EndHC wraps the send in its
        # own try/catch, so a throwing Send-MailKitMessageHC becomes a 'Warning'
        # (Name 'Email Failed') rather than escaping the finally block and
        # crashing the orchestrator.
        $matrixDir = (New-Item 'TestDrive:\Matrix-smtp' -ItemType Directory -Force).FullName
        $logsDir = (New-Item 'TestDrive:\Logs-smtp' -ItemType Directory -Force).FullName

        $defaultsPath = Join-Path $matrixDir 'Defaults.xlsx'
        New-ValidDefaultsExcelFixture -Path $defaultsPath | Out-Null

        $matrixPath = Join-Path $matrixDir 'TeamA.xlsx'
        New-MatrixExcelFixture `
            -Path $matrixPath `
            -SettingsRows @(
            [pscustomobject]@{
                Status                  = 'Enabled'
                SiteName                = 'E2E'
                SiteCode                = 'E2E'
                ComputerName            = 'TESTHOST'
                Path                    = 'C:\DoesNotMatter'
                GroupName               = 'E2E-Test-Group'
                Action                  = 'Fix'
                ApplyDefaultPermissions = $false
            }
        )

        $config = New-JsonFixture
        $config.Matrix.FolderPath = $matrixDir
        $config.Matrix.DefaultsFile = $defaultsPath
        $config.Settings.SaveLogFiles.Where.Folder = $logsDir
        $config.MaxConcurrent.FoldersPerMatrix = 1

        $configPath = Join-Path $matrixDir 'Input.json'
        Save-Config -Config $config -Path $configPath

        Mock Get-ADObjectDetailHC -ModuleName PermissionMatrix {
            @(
                @{ SamAccountName = 'Bob'; adObject = @{ ObjectSid = 'S-1-5-21-0-0-0-1001' } }
                @{ SamAccountName = 'Mike'; adObject = @{ ObjectSid = 'S-1-5-21-0-0-0-1002' } }
            )
        }
        # PROCESS neutralised (return the context so END still runs); the SMTP
        # send is the only thing that fails.
        Mock Invoke-PermissionMatrixProcessHC -ModuleName PermissionMatrix { $Context }
        Mock Send-MailKitMessageHC -ModuleName PermissionMatrix { throw 'smtp down' }

        $systemErrors = [System.Collections.Generic.List[object]]::new()

        { Invoke-PermissionMatrix `
                -ConfigurationJsonFile $configPath `
                -ScriptPath $ScriptPath `
                -SystemErrors ([ref]$systemErrors) } |
        Should -Not -Throw -Because 'EndHC swallows a send failure instead of letting it escape the finally'

        Should -Invoke Send-MailKitMessageHC -ModuleName PermissionMatrix -Times 1 -Exactly -Because (
            'the send is attempted exactly once before it fails'
        )

        ($systemErrors.Where({ $_.Name -eq 'Email Failed' })).Count |
        Should -BeGreaterThan 0 -Because 'a failed send is recorded as a warning so the operator can see mail did not go out'
    }
}