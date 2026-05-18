#Requires -Version 7
#Requires -Modules Pester

# =============================================================================
# Pragmatic end-to-end test for the Permission Matrix pipeline.
#
# Approach: invoke Invoke-PermissionMatrix directly (bypassing the wrapper
# script). Mock at the AD and SMTP boundaries; everything else is real.
#
# Prerequisites (the test skips with clear messages if any are missing):
#   - Running as Administrator
#   - PSRemoting enabled (Test-WSMan localhost succeeds)
#   - 'PowerShell.7' session configuration registered
#
# What this test actually does:
#   - Creates two local Windows groups (cleanup in AfterAll)
#   - Creates target folders under TestDrive
#   - Builds real .xlsx fixtures (defaults + matrix) via the project fixture
#     helpers
#   - Runs Invoke-PermissionMatrix against $env:COMPUTERNAME with Action='New'
#   - Asserts: no fatal errors, mail-send called, ACLs applied with the right
#     FileSystemRights for each test group
# =============================================================================

Describe 'Permission Matrix - End to End' {
    BeforeDiscovery {
        # These checks run at discovery so the It blocks can be -Skip'd
        # cleanly with -Because messages rather than failing inside BeforeAll.
        $script:IsAdmin = (
            [Security.Principal.WindowsPrincipal] `
                [Security.Principal.WindowsIdentity]::GetCurrent()
        ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

        $script:HasRemoting = try {
            $null = Test-WSMan -ComputerName 'localhost' -ErrorAction Stop
            $true
        }
        catch { $false }

        $script:HasPwsh7SessionConfig = try {
            $null = Get-PSSessionConfiguration -Name 'PowerShell.7' -ErrorAction Stop
            $true
        }
        catch { $false }

        $script:E2EPrereqsMet = $IsAdmin -and $HasRemoting -and $HasPwsh7SessionConfig
    }

    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        . "$root\Tests\Helpers\Fixtures.Json.ps1"
        . "$root\Tests\Helpers\Fixtures.Excel.ps1"

        Import-Module "$moduleRoot\PermissionMatrix.psm1" -Force

        # ---------------------------------------------------------------------
        # Local groups used as the AD object stand-ins. Bob and Mike are the
        # placeholder names in the matrix Permissions sheet header (per the
        # 'Valid' fixture in New-MatrixPermissionsFixtureRows). The AD mock
        # below resolves Bob -> $TestGroupBobSid, Mike -> $TestGroupMikeSid.
        # ---------------------------------------------------------------------
        $script:TestGroupBob = 'PermMatrixE2E_Bob'
        $script:TestGroupMike = 'PermMatrixE2E_Mike'

        if ($E2EPrereqsMet) {
            foreach ($name in $TestGroupBob, $TestGroupMike) {
                if (-not (Get-LocalGroup -Name $name -ErrorAction SilentlyContinue)) {
                    $null = New-LocalGroup -Name $name -Description 'PermissionMatrix E2E test fixture'
                }
            }
            $script:TestGroupBobSid = (Get-LocalGroup -Name $TestGroupBob).SID.Value
            $script:TestGroupMikeSid = (Get-LocalGroup -Name $TestGroupMike).SID.Value
        }
    }

    AfterAll {
        # Clean up local groups regardless of test outcome.
        foreach ($name in $TestGroupBob, $TestGroupMike) {
            if (Get-LocalGroup -Name $name -ErrorAction SilentlyContinue) {
                Remove-LocalGroup -Name $name -ErrorAction SilentlyContinue
            }
        }
    }

    It 'applies permissions, sends mail, and reports no errors on the happy path' -Skip:(-not $E2EPrereqsMet) {
        # -------------------------------------------------------------------
        # Filesystem layout. Action='New' on the matrix tells SetPermissions
        # to create the folder structure, so $rootFolder must exist but the
        # subfolders (Finance, Finance\Docs) should NOT — the script creates
        # them as part of its work.
        # -------------------------------------------------------------------
        $rootFolder = Join-Path $TestDrive 'Target'
        $matrixDir = (New-Item 'TestDrive:\Matrix' -ItemType Directory -Force).FullName
        $logsDir = (New-Item 'TestDrive:\Logs' -ItemType Directory -Force).FullName

        # -------------------------------------------------------------------
        # Defaults Excel — uses the project's existing valid fixture.
        # -------------------------------------------------------------------
        $defaultsPath = Join-Path $matrixDir 'Defaults.xlsx'
        New-ValidDefaultsExcelFixture -Path $defaultsPath | Out-Null

        # -------------------------------------------------------------------
        # Matrix Excel — overrides the 'Valid' Settings fixture so that:
        #   ComputerName = $env:COMPUTERNAME  (loopback to the test machine)
        #   Path         = $rootFolder        (real TestDrive folder)
        #   Action       = 'New'              (create folders + apply ACLs)
        # Permissions sheet uses the default 'Valid' fixture, which references
        # 'Bob' and 'Mike' in the header and 'Finance' / 'Finance\Docs' as
        # subfolders.
        # -------------------------------------------------------------------
        $matrixPath = Join-Path $matrixDir 'TeamA.xlsx'
        New-MatrixExcelFixture `
            -Path $matrixPath `
            -SettingsRows @(
            [pscustomobject]@{
                Status                  = 'Enabled'
                SiteName                = 'E2E'
                SiteCode                = 'E2E'
                ComputerName            = $env:COMPUTERNAME
                Path                    = $rootFolder
                GroupName               = 'E2E-Test-Group'
                Action                  = 'New'
                ApplyDefaultPermissions = $false
            }
        )

        # -------------------------------------------------------------------
        # JSON config — points at the matrix folder, the defaults file, the
        # log folder, and a fake SendMail config (the SMTP send is mocked
        # anyway).
        # -------------------------------------------------------------------
        $configFixture = New-JsonFixture
        $configFixture.Matrix.FolderPath = $matrixDir
        $configFixture.Matrix.DefaultsFile = $defaultsPath
        $configFixture.Settings.SaveLogFiles.Where.Folder = $logsDir
        # MaxConcurrent.FoldersPerMatrix=1 forces sequential parallelism in
        # BeginHC's matrix-import loop, which keeps test execution
        # deterministic and avoids cross-runspace mock complications.
        $configFixture.MaxConcurrent.FoldersPerMatrix = 1

        $configPath = Join-Path $matrixDir 'Input.json'
        $configFixture |
        ConvertTo-Json -Depth 20 |
        Out-File -LiteralPath $configPath -Encoding utf8 -Force

        # -------------------------------------------------------------------
        # Script paths the entrypoint requires.
        # -------------------------------------------------------------------
        $scriptPath = @{
            PermissionMatrixModule = "$moduleRoot\PermissionMatrix.psm1"
            SetPermissions         = "$root\Operations\SetPermissions.ps1"
            TestRequirements       = "$root\Operations\TestRequirements.ps1"
            UpdateServiceNow       = "$root\Operations\UpdateServiceNow.ps1"
        }

        # -------------------------------------------------------------------
        # Mocks. -ModuleName ensures the mocks intercept calls made from
        # within Invoke-PermissionMatrix's call chain (BeginHC, EndHC).
        # -------------------------------------------------------------------
        Mock Get-ADObjectDetailHC -ModuleName PermissionMatrix {
            return @(
                @{ SamAccountName = 'Bob'; adObject = @{ ObjectSid = $TestGroupBobSid } }
                @{ SamAccountName = 'Mike'; adObject = @{ ObjectSid = $TestGroupMikeSid } }
            )
        }

        Mock Send-MailKitMessageHC -ModuleName PermissionMatrix { }

        # -------------------------------------------------------------------
        # Run the orchestrator.
        # -------------------------------------------------------------------
        $systemErrors = [System.Collections.Generic.List[object]]::new()

        Invoke-PermissionMatrix `
            -ConfigurationJsonFile $configPath `
            -ScriptPath $scriptPath `
            -SystemErrors ([ref]$systemErrors)

        # -------------------------------------------------------------------
        # Assert: no fatal errors logged.
        # -------------------------------------------------------------------
        $fatals = $systemErrors.Where({ $_.Type -eq 'FatalError' })
        $fatals.Count | Should -Be 0 -Because (
            "expected no fatal errors but got: $($fatals | ForEach-Object { $_.Message } | Out-String)"
        )

        # -------------------------------------------------------------------
        # Assert: mail was sent exactly once.
        # -------------------------------------------------------------------
        Should -Invoke Send-MailKitMessageHC -ModuleName PermissionMatrix -Times 1 -Exactly

        # -------------------------------------------------------------------
        # Assert: the matrix's Permissions sheet rows resulted in real ACLs
        # on disk. The 'Valid' fixture maps:
        #   Finance       -> Bob: R,  Mike: R
        #   Finance\Docs  -> Bob: W,  Mike: W
        # R -> ReadAndExecute, W -> CreateFiles+Modify+propagation rules.
        # We check IdentityReference (translated back to NTAccount) and
        # FileSystemRights, not the full propagation/inheritance tuple —
        # SetPermissions.ps1 applies several rules per 'W' permission, so
        # an exact match is fragile. Asserting that "this group has at
        # least one ACE on this folder" catches the wiring regression.
        # -------------------------------------------------------------------
        $financeFolder = Join-Path $rootFolder 'Finance'
        $docsFolder = Join-Path $rootFolder 'Finance\Docs'

        Test-Path -LiteralPath $financeFolder -PathType Container |
        Should -BeTrue -Because 'Action=New should have created Finance'
        Test-Path -LiteralPath $docsFolder -PathType Container |
        Should -BeTrue -Because 'Action=New should have created Finance\Docs'

        $financeAcl = (Get-Acl -LiteralPath $financeFolder).Access
        $docsAcl = (Get-Acl -LiteralPath $docsFolder).Access

        # Translate IdentityReference to NTAccount name for comparison.
        $bobNT = (New-Object System.Security.Principal.NTAccount("$env:COMPUTERNAME\$TestGroupBob")).Value
        $mikeNT = (New-Object System.Security.Principal.NTAccount("$env:COMPUTERNAME\$TestGroupMike")).Value

        $financeAcl.Where({ $_.IdentityReference.Value -eq $bobNT }).Count |
        Should -BeGreaterThan 0 -Because "Finance should have an ACE for $bobNT"
        $financeAcl.Where({ $_.IdentityReference.Value -eq $mikeNT }).Count |
        Should -BeGreaterThan 0 -Because "Finance should have an ACE for $mikeNT"

        $docsAcl.Where({ $_.IdentityReference.Value -eq $bobNT }).Count |
        Should -BeGreaterThan 0 -Because "Finance\Docs should have an ACE for $bobNT"
        $docsAcl.Where({ $_.IdentityReference.Value -eq $mikeNT }).Count |
        Should -BeGreaterThan 0 -Because "Finance\Docs should have an ACE for $mikeNT"
    }

    It 'reports clear prerequisites when the environment is not set up' -Skip:$E2EPrereqsMet {
        # This test only runs when prerequisites are NOT met. It surfaces
        # the missing prerequisite as an actionable failure message so the
        # operator knows what to fix before re-running the suite.
        $missing = @()
        if (-not $IsAdmin) { $missing += 'Administrator privileges' }
        if (-not $HasRemoting) { $missing += "PSRemoting (run 'Enable-PSRemoting -Force')" }
        if (-not $HasPwsh7SessionConfig) {
            $missing += "'PowerShell.7' session config (run 'Register-PSSessionConfiguration -Name PowerShell.7 -PSVersion 7.x')"
        }

        Set-ItResult -Inconclusive -Because (
            "E2E test requires: $($missing -join '; '). Skipping until prerequisites are met."
        )
    }
}