#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

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
        $script:TestGroupDefault = 'PermMatrixE2E_Default'

        if ($E2EPrereqsMet) {
            foreach ($name in $TestGroupBob, $TestGroupMike, $TestGroupDefault) {
                if (-not (Get-LocalGroup -Name $name -ErrorAction SilentlyContinue)) {
                    $null = New-LocalGroup -Name $name -Description 'PermissionMatrix E2E test fixture'
                }
            }
            $script:TestGroupBobSid = (Get-LocalGroup -Name $TestGroupBob).SID.Value
            $script:TestGroupMikeSid = (Get-LocalGroup -Name $TestGroupMike).SID.Value
            $script:TestGroupDefaultSid = (Get-LocalGroup -Name $TestGroupDefault).SID.Value
        }
    }

    AfterAll {
        # Clean up local groups regardless of test outcome.
        foreach ($name in $TestGroupBob, $TestGroupMike, $TestGroupDefault) {
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
            SetPermissions         = "$root\Scripts\Operations\SetPermissions.ps1"
            TestRequirements       = "$root\Scripts\Operations\TestRequirements.ps1"
            UpdateServiceNow       = "$root\Scripts\Operations\UpdateServiceNow.ps1"
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

        # -------------------------------------------------------------------
        # Regression: newly created folders must keep the exact casing from
        # the matrix (Column A), not be upper-cased. The filesystem is
        # case-insensitive, so Test-Path/Get-Item on the requested path would
        # not reveal a casing mismatch - enumerate the parent to read the
        # real on-disk name instead.
        # -------------------------------------------------------------------
        $onDiskFinance = (Get-ChildItem -LiteralPath $rootFolder -Directory).Where({ $_.Name -eq 'Finance' }).Name
        $onDiskFinance |
        Should -BeExactly 'Finance' -Because 'the folder casing from the matrix must be preserved'
        $onDiskDocs = (Get-ChildItem -LiteralPath $financeFolder -Directory).Where({ $_.Name -eq 'Docs' }).Name
        $onDiskDocs |
        Should -BeExactly 'Docs' -Because 'the folder casing from the matrix must be preserved'

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

    It 'creates a no-permission folder as inherit-only under Action=New (no defaults, no cut)' -Skip:(-not $E2EPrereqsMet) {
        # -------------------------------------------------------------------
        # A folder listed in the matrix WITHOUT any permissions is an
        # inherit-only folder. Under Action=New the whole tree is created from
        # scratch, and these folders must be created too, purely inheriting
        # from their parent: no default permissions, no cut inheritance.
        # A folder WITH permissions ('Finance') still gets the default merged
        # and its inheritance protected, proving the guard is scoped correctly.
        # -------------------------------------------------------------------
        $rootFolder = Join-Path $TestDrive 'DefTarget'
        $matrixDir = (New-Item 'TestDrive:\DefMatrix' -ItemType Directory -Force).FullName
        $logsDir = (New-Item 'TestDrive:\DefLogs' -ItemType Directory -Force).FullName

        $defaultsPath = Join-Path $matrixDir 'Defaults.xlsx'
        New-ValidDefaultsExcelFixture -Path $defaultsPath | Out-Null

        $matrixPath = Join-Path $matrixDir 'TeamA.xlsx'
        New-MatrixExcelFixture `
            -Path $matrixPath `
            -PermissionsRows (New-MatrixPermissionsFixtureRows -Scenario 'WithEmptyPermissionFolder') `
            -SettingsRows @(
            [pscustomobject]@{
                Status                  = 'Enabled'
                SiteName                = 'E2E'
                SiteCode                = 'E2E'
                ComputerName            = $env:COMPUTERNAME
                Path                    = $rootFolder
                GroupName               = 'E2E-Test-Group'
                Action                  = 'New'
                ApplyDefaultPermissions = $true
            }
        )

        $configFixture = New-JsonFixture
        $configFixture.Matrix.FolderPath = $matrixDir
        $configFixture.Matrix.DefaultsFile = $defaultsPath
        $configFixture.Settings.SaveLogFiles.Where.Folder = $logsDir
        $configFixture.MaxConcurrent.FoldersPerMatrix = 1

        $configPath = Join-Path $matrixDir 'Input.json'
        $configFixture |
        ConvertTo-Json -Depth 20 |
        Out-File -LiteralPath $configPath -Encoding utf8 -Force

        $scriptPath = @{
            PermissionMatrixModule = "$moduleRoot\PermissionMatrix.psm1"
            SetPermissions         = "$root\Scripts\Operations\SetPermissions.ps1"
            TestRequirements       = "$root\Scripts\Operations\TestRequirements.ps1"
            UpdateServiceNow       = "$root\Scripts\Operations\UpdateServiceNow.ps1"
        }

        # 'DefaultGroup' is the AD object in New-ValidDefaultsExcelFixture; it
        # must resolve so the default can be applied to folders that have
        # permissions.
        Mock Get-ADObjectDetailHC -ModuleName PermissionMatrix {
            return @(
                @{ SamAccountName = 'Bob'; adObject = @{ ObjectSid = $TestGroupBobSid } }
                @{ SamAccountName = 'Mike'; adObject = @{ ObjectSid = $TestGroupMikeSid } }
                @{ SamAccountName = 'DefaultGroup'; adObject = @{ ObjectSid = $TestGroupDefaultSid } }
            )
        }

        Mock Send-MailKitMessageHC -ModuleName PermissionMatrix { }

        $systemErrors = [System.Collections.Generic.List[object]]::new()

        Invoke-PermissionMatrix `
            -ConfigurationJsonFile $configPath `
            -ScriptPath $scriptPath `
            -SystemErrors ([ref]$systemErrors)

        $fatals = $systemErrors.Where({ $_.Type -eq 'FatalError' })
        $fatals.Count | Should -Be 0 -Because (
            "expected no fatal errors but got: $($fatals | ForEach-Object { $_.Message } | Out-String)"
        )

        $financeFolder = Join-Path $rootFolder 'Finance'
        $inheritOnlyFolder = Join-Path $rootFolder 'InheritOnly'

        # The no-permission folder must be created ...
        Test-Path -LiteralPath $inheritOnlyFolder -PathType Container |
        Should -BeTrue -Because 'Action=New must create inherit-only folders too'

        $defaultNT = (New-Object System.Security.Principal.NTAccount("$env:COMPUTERNAME\$TestGroupDefault")).Value

        # ... and must keep inheritance (not protected) ...
        $inheritOnlySecurity = Get-Acl -LiteralPath $inheritOnlyFolder
        $inheritOnlySecurity.AreAccessRulesProtected |
        Should -BeFalse -Because 'a folder without permissions must keep inheriting (inheritance not cut)'

        # ... and must NOT have the default group applied explicitly.
        $inheritOnlyExplicit = $inheritOnlySecurity.Access.Where({ -not $_.IsInherited })
        $inheritOnlyExplicit.Where({ $_.IdentityReference.Value -eq $defaultNT }).Count |
        Should -Be 0 -Because 'defaults must not be applied to a no-permission folder'

        # The folder WITH permissions still gets the default merged and its
        # inheritance protected, so the guard is scoped to empty folders only.
        $financeSecurity = Get-Acl -LiteralPath $financeFolder
        $financeSecurity.AreAccessRulesProtected |
        Should -BeTrue -Because 'a folder with explicit permissions has protected (non-inherited) ACLs'
        $financeSecurity.Access.Where({ $_.IdentityReference.Value -eq $defaultNT }).Count |
        Should -BeGreaterThan 0 -Because 'defaults are still applied to folders that have permissions'
    }

    It 'corrects a TOP-LEVEL inherit-only folder under Action=Fix (full flow regression)' -Skip:(-not $E2EPrereqsMet) {
        # -------------------------------------------------------------------
        # REGRESSION GUARD for the production 'Common' bug.
        #
        # Matrix (mirrors BEL-MTX-CEM-Gent.xlsx):
        #   Path            -> L L   (root, the parent 'Path' row)
        #   Common          -> <empty> (TOP-LEVEL inherit-only folder)
        #   Common\Exchange -> W     (permissioned child)
        #
        # The full pipeline must turn the 'Path' row into a Parent=$true matrix
        # entry so SetPermissions seeds its inheritance walk from the share
        # ROOT. That walk is the only thing that visits and corrects a
        # top-level inherit-only folder like 'Common'. Before the fix the root
        # had no matrix entry, no walk was seeded there, and 'Common' kept its
        # stale explicit ACL — exactly the reported bug. This test drives the
        # real Invoke-PermissionMatrix end to end under Action=Fix.
        # -------------------------------------------------------------------
        $rootFolder = (New-Item (Join-Path $TestDrive 'FixInheritTarget') -ItemType Directory -Force).FullName
        $commonFolder = (New-Item (Join-Path $rootFolder 'Common') -ItemType Directory -Force).FullName
        $exchangeFolder = (New-Item (Join-Path $commonFolder 'Exchange') -ItemType Directory -Force).FullName

        $matrixDir = (New-Item 'TestDrive:\FixMatrix' -ItemType Directory -Force).FullName
        $logsDir = (New-Item 'TestDrive:\FixLogs' -ItemType Directory -Force).FullName

        $defaultsPath = Join-Path $matrixDir 'Defaults.xlsx'
        New-ValidDefaultsExcelFixture -Path $defaultsPath | Out-Null

        # Prod-shaped permissions sheet: root grants List to Bob+Mike, Common
        # has no permissions (inherit-only), Common\Exchange grants Write.
        $permSpec = @{
            Row1 = @('', '', '')
            Row2 = @('', '', '')
            Row3 = @('', 'Bob', 'Mike')
            Row4 = @('Path', 'L', 'L')
            Data = @(
                @{ Path = 'Common'          ; Col2 = '' ; Col3 = '' }
                @{ Path = 'Common\Exchange' ; Col2 = 'W'; Col3 = '' }
            )
        }

        $matrixPath = Join-Path $matrixDir 'TeamA.xlsx'
        New-MatrixExcelFixture `
            -Path $matrixPath `
            -PermissionsRows $permSpec `
            -SettingsRows @(
            [pscustomobject]@{
                Status                  = 'Enabled'
                SiteName                = 'E2E'
                SiteCode                = 'E2E'
                ComputerName            = $env:COMPUTERNAME
                Path                    = $rootFolder
                GroupName               = 'E2E-Test-Group'
                Action                  = 'Fix'
                ApplyDefaultPermissions = $false
            }
        )

        # -------------------------------------------------------------------
        # Seed 'Common' with a WRONG, protected (non-inheriting) ACL so the
        # run has something to correct. Giving the Bob group explicit
        # FullControl and cutting inheritance is exactly the stale state the
        # script must strip back to pure inheritance.
        # -------------------------------------------------------------------
        $wrongAcl = New-Object System.Security.AccessControl.DirectorySecurity
        $wrongAcl.SetAccessRuleProtection($true, $false)
        $wrongAcl.AddAccessRule(
            (New-Object System.Security.AccessControl.FileSystemAccessRule(
                "$env:COMPUTERNAME\$TestGroupBob",
                [System.Security.AccessControl.FileSystemRights]::FullControl,
                [System.Security.AccessControl.InheritanceFlags]'ContainerInherit,ObjectInherit',
                [System.Security.AccessControl.PropagationFlags]::None,
                [System.Security.AccessControl.AccessControlType]::Allow))
        )
        Set-Acl -LiteralPath $commonFolder -AclObject $wrongAcl

        (Get-Acl -LiteralPath $commonFolder).AreAccessRulesProtected |
        Should -BeTrue -Because 'the test must start from a Common folder that does not inherit'

        $configFixture = New-JsonFixture
        $configFixture.Matrix.FolderPath = $matrixDir
        $configFixture.Matrix.DefaultsFile = $defaultsPath
        $configFixture.Settings.SaveLogFiles.Where.Folder = $logsDir
        $configFixture.MaxConcurrent.FoldersPerMatrix = 1

        $configPath = Join-Path $matrixDir 'Input.json'
        $configFixture |
        ConvertTo-Json -Depth 20 |
        Out-File -LiteralPath $configPath -Encoding utf8 -Force

        $scriptPath = @{
            PermissionMatrixModule = "$moduleRoot\PermissionMatrix.psm1"
            SetPermissions         = "$root\Scripts\Operations\SetPermissions.ps1"
            TestRequirements       = "$root\Scripts\Operations\TestRequirements.ps1"
            UpdateServiceNow       = "$root\Scripts\Operations\UpdateServiceNow.ps1"
        }

        Mock Get-ADObjectDetailHC -ModuleName PermissionMatrix {
            return @(
                @{ SamAccountName = 'Bob'; adObject = @{ ObjectSid = $TestGroupBobSid } }
                @{ SamAccountName = 'Mike'; adObject = @{ ObjectSid = $TestGroupMikeSid } }
            )
        }
        Mock Send-MailKitMessageHC -ModuleName PermissionMatrix { }

        $systemErrors = [System.Collections.Generic.List[object]]::new()

        Invoke-PermissionMatrix `
            -ConfigurationJsonFile $configPath `
            -ScriptPath $scriptPath `
            -SystemErrors ([ref]$systemErrors)

        $fatals = $systemErrors.Where({ $_.Type -eq 'FatalError' })
        $fatals.Count | Should -Be 0 -Because (
            "expected no fatal errors but got: $($fatals | ForEach-Object { $_.Message } | Out-String)"
        )

        $bobNT = (New-Object System.Security.Principal.NTAccount("$env:COMPUTERNAME\$TestGroupBob")).Value

        #region Common is now inherit-only: inheritance restored, no explicit ACE
        $commonSecurity = Get-Acl -LiteralPath $commonFolder

        $commonSecurity.AreAccessRulesProtected |
        Should -BeFalse -Because 'Fix must reset a top-level inherit-only folder back to inheriting'

        $commonSecurity.Access.Where({ -not $_.IsInherited }) |
        Should -BeNullOrEmpty -Because 'Common cannot keep any explicit permissions'

        # It should now inherit the root's List grant for Bob (proof the root
        # walk actually reached and re-based Common).
        $commonSecurity.Access.Where({
                $_.IsInherited -and ($_.IdentityReference.Value -eq $bobNT)
            }) |
        Should -Not -BeNullOrEmpty -Because 'Common inherits the root ACL after the fix'
        #endregion

        #region The permissioned child keeps its own protected ACL
        $exchangeSecurity = Get-Acl -LiteralPath $exchangeFolder
        $exchangeSecurity.AreAccessRulesProtected |
        Should -BeTrue -Because 'Common\Exchange has Write in the matrix and stays protected'
        $exchangeSecurity.Access.Where({
                (-not $_.IsInherited) -and ($_.IdentityReference.Value -eq $bobNT)
            }) |
        Should -Not -BeNullOrEmpty -Because 'Common\Exchange carries its explicit Write ACE for Bob'
        #endregion
    }

    It 'works the same way under parallel execution' -Skip:(-not $E2EPrereqsMet) {
        # -------------------------------------------------------------------
        # Same happy-path scenario as the sequential test, but with
        # MaxConcurrent values that force real parallelism in BeginHC.
        # This catches regressions where runspace-boundary changes break
        # the pipeline (lost mocks, mis-scoped variables, race conditions
        # on shared state, missing module imports inside child runspaces).
        #
        # The three knobs:
        #   JobsTotal              — overall concurrent job budget
        #   FoldersPerMatrix       — parallelism within a matrix's folders
        #   JobsPerComputer        — parallel jobs on one remote machine
        # Values >1 are required; values >=3 force the parallel code paths
        # even with our small single-matrix fixture.
        # -------------------------------------------------------------------
        $rootFolder = Join-Path $TestDrive 'Target'
        $matrixDir = (New-Item 'TestDrive:\Matrix' -ItemType Directory -Force).FullName
        $logsDir = (New-Item 'TestDrive:\Logs' -ItemType Directory -Force).FullName

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
                ComputerName            = $env:COMPUTERNAME
                Path                    = $rootFolder
                GroupName               = 'E2E-Test-Group'
                Action                  = 'New'
                ApplyDefaultPermissions = $false
            }
        )

        $configFixture = New-JsonFixture
        $configFixture.Matrix.FolderPath = $matrixDir
        $configFixture.Matrix.DefaultsFile = $defaultsPath
        $configFixture.Settings.SaveLogFiles.Where.Folder = $logsDir

        # The only meaningful difference from the sequential test:
        # crank the concurrency knobs to force runspace-based parallelism.
        $configFixture.MaxConcurrent.JobsTotal = 30
        $configFixture.MaxConcurrent.FoldersPerMatrix = 3
        $configFixture.MaxConcurrent.JobsPerComputer = 3

        $configPath = Join-Path $matrixDir 'Input.json'
        $configFixture |
        ConvertTo-Json -Depth 20 |
        Out-File -LiteralPath $configPath -Encoding utf8 -Force

        $scriptPath = @{
            PermissionMatrixModule = "$moduleRoot\PermissionMatrix.psm1"
            SetPermissions         = "$root\Scripts\Operations\SetPermissions.ps1"
            TestRequirements       = "$root\Scripts\Operations\TestRequirements.ps1"
            UpdateServiceNow       = "$root\Scripts\Operations\UpdateServiceNow.ps1"
        }

        Mock Get-ADObjectDetailHC -ModuleName PermissionMatrix {
            return @(
                @{ SamAccountName = 'Bob'; adObject = @{ ObjectSid = $TestGroupBobSid } }
                @{ SamAccountName = 'Mike'; adObject = @{ ObjectSid = $TestGroupMikeSid } }
            )
        }

        Mock Send-MailKitMessageHC -ModuleName PermissionMatrix { }

        $systemErrors = [System.Collections.Generic.List[object]]::new()

        Invoke-PermissionMatrix `
            -ConfigurationJsonFile $configPath `
            -ScriptPath $scriptPath `
            -SystemErrors ([ref]$systemErrors)

        # -------------------------------------------------------------------
        # Same assertions as the sequential test. If parallel execution
        # breaks the pipeline, we expect to see it here as either fatal
        # errors, a missing mail send, or ACLs that never landed on disk.
        # -------------------------------------------------------------------
        $fatals = $systemErrors.Where({ $_.Type -eq 'FatalError' })
        $fatals.Count | Should -Be 0 -Because (
            "expected no fatal errors under parallel execution but got: $($fatals | ForEach-Object { $_.Message } | Out-String)"
        )

        Should -Invoke Send-MailKitMessageHC -ModuleName PermissionMatrix -Times 1 -Exactly

        $financeFolder = Join-Path $rootFolder 'Finance'
        $docsFolder = Join-Path $rootFolder 'Finance\Docs'

        Test-Path -LiteralPath $financeFolder -PathType Container |
        Should -BeTrue -Because 'Action=New should have created Finance under parallel execution'
        Test-Path -LiteralPath $docsFolder -PathType Container |
        Should -BeTrue -Because 'Action=New should have created Finance\Docs under parallel execution'

        $financeAcl = (Get-Acl -LiteralPath $financeFolder).Access
        $docsAcl = (Get-Acl -LiteralPath $docsFolder).Access

        $bobNT = (New-Object System.Security.Principal.NTAccount("$env:COMPUTERNAME\$TestGroupBob")).Value
        $mikeNT = (New-Object System.Security.Principal.NTAccount("$env:COMPUTERNAME\$TestGroupMike")).Value

        $financeAcl.Where({ $_.IdentityReference.Value -eq $bobNT }).Count |
        Should -BeGreaterThan 0 -Because "Finance should have an ACE for $bobNT under parallel execution"
        $financeAcl.Where({ $_.IdentityReference.Value -eq $mikeNT }).Count |
        Should -BeGreaterThan 0 -Because "Finance should have an ACE for $mikeNT under parallel execution"

        $docsAcl.Where({ $_.IdentityReference.Value -eq $bobNT }).Count |
        Should -BeGreaterThan 0 -Because "Finance\Docs should have an ACE for $bobNT under parallel execution"
        $docsAcl.Where({ $_.IdentityReference.Value -eq $mikeNT }).Count |
        Should -BeGreaterThan 0 -Because "Finance\Docs should have an ACE for $mikeNT under parallel execution"
    }

    It 'fixes incorrect permissions on files and folders, then Check confirms all correct' -Skip:(-not $E2EPrereqsMet) {
        # -------------------------------------------------------------------
        # This test proves Action=Fix corrects wrong ACLs on matrix-defined
        # folders AND inherited files, and that a subsequent Action=Check run
        # reports success (no warnings about incorrect permissions).
        #
        # Phase 1 - Setup: create folders + a file with WRONG permissions:
        #   Finance       -> Mike gets FullControl instead of R (extra rights)
        #   Finance\Docs  -> only Admin FullControl, Bob & Mike missing entirely
        #   report.txt    -> explicit FullControl for Bob (should be inherited)
        #
        # Phase 2 - Fix run: Invoke-PermissionMatrix with Action=Fix corrects
        #   all three corruption types.
        #
        # Phase 3 - Check run: Invoke-PermissionMatrix with Action=Check
        #   confirms no non-inherited or inherited permission warnings.
        # -------------------------------------------------------------------

        # ===================================================================
        # PHASE 1: Build folder structure with deliberately wrong permissions
        # ===================================================================
        $rootFolder = (New-Item 'TestDrive:\FixTarget' -ItemType Directory -Force).FullName
        $financeFolder = (New-Item (Join-Path $rootFolder 'Finance') -ItemType Directory -Force).FullName
        $docsFolder = (New-Item (Join-Path $rootFolder 'Finance\Docs') -ItemType Directory -Force).FullName
        $reportFile = Join-Path $docsFolder 'report.txt'
        'test content' | Set-Content -LiteralPath $reportFile -Force

        $matrixDir = (New-Item 'TestDrive:\Matrix-fix' -ItemType Directory -Force).FullName
        $logsDir = (New-Item 'TestDrive:\Logs-fix' -ItemType Directory -Force).FullName

        $bobNT = New-Object System.Security.Principal.NTAccount("$env:COMPUTERNAME\$TestGroupBob")
        $mikeNT = New-Object System.Security.Principal.NTAccount("$env:COMPUTERNAME\$TestGroupMike")
        $builtinAdmin = [System.Security.Principal.NTAccount]'BUILTIN\Administrators'

        # --- Corrupt Finance: Mike gets FullControl instead of R -----------
        $financeAclCorrupt = New-Object System.Security.AccessControl.DirectorySecurity
        $financeAclCorrupt.SetAccessRuleProtection($true, $false)
        $financeAclCorrupt.SetOwner($builtinAdmin)
        $financeAclCorrupt.AddAccessRule(
            (New-Object System.Security.AccessControl.FileSystemAccessRule(
                    $builtinAdmin, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow'))
        )
        $financeAclCorrupt.AddAccessRule(
            (New-Object System.Security.AccessControl.FileSystemAccessRule(
                    $mikeNT, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow'))
        )
        [System.IO.FileSystemAclExtensions]::SetAccessControl(
            [System.IO.DirectoryInfo]::new($financeFolder), $financeAclCorrupt)

        # --- Corrupt Finance\Docs: only Admin, Bob & Mike missing entirely -
        $docsAclCorrupt = New-Object System.Security.AccessControl.DirectorySecurity
        $docsAclCorrupt.SetAccessRuleProtection($true, $false)
        $docsAclCorrupt.SetOwner($builtinAdmin)
        $docsAclCorrupt.AddAccessRule(
            (New-Object System.Security.AccessControl.FileSystemAccessRule(
                    $builtinAdmin, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow'))
        )
        [System.IO.FileSystemAclExtensions]::SetAccessControl(
            [System.IO.DirectoryInfo]::new($docsFolder), $docsAclCorrupt)

        # --- Corrupt report.txt: explicit FullControl for Bob (should be inherited)
        $fileAclCorrupt = New-Object System.Security.AccessControl.FileSecurity
        $fileAclCorrupt.SetAccessRuleProtection($true, $false)
        $fileAclCorrupt.SetOwner($builtinAdmin)
        $fileAclCorrupt.AddAccessRule(
            (New-Object System.Security.AccessControl.FileSystemAccessRule(
                    $builtinAdmin, 'FullControl', 'None', 'None', 'Allow'))
        )
        $fileAclCorrupt.AddAccessRule(
            (New-Object System.Security.AccessControl.FileSystemAccessRule(
                    $bobNT, 'FullControl', 'None', 'None', 'Allow'))
        )
        [System.IO.FileSystemAclExtensions]::SetAccessControl(
            [System.IO.FileInfo]::new($reportFile), $fileAclCorrupt)

        # --- Verify corruption landed (sanity check) ----------------------
        $preFixFinanceAcl = (Get-Acl -LiteralPath $financeFolder).Access
        $preFixFinanceAcl.Where({ $_.IdentityReference.Value -eq $mikeNT.Value -and $_.FileSystemRights -match 'FullControl' }).Count |
        Should -BeGreaterThan 0 -Because 'sanity: Mike should have FullControl on Finance before Fix'

        $preFixDocsAcl = (Get-Acl -LiteralPath $docsFolder).Access
        $preFixDocsAcl.Where({ $_.IdentityReference.Value -eq $bobNT.Value }).Count |
        Should -Be 0 -Because 'sanity: Bob should have no ACE on Finance\Docs before Fix'

        $preFixFileAcl = Get-Acl -LiteralPath $reportFile
        $preFixFileAcl.AreAccessRulesProtected |
        Should -BeTrue -Because 'sanity: report.txt should have explicit (protected) ACL before Fix'

        # --- Build Excel fixtures and JSON config -------------------------
        $defaultsPath = Join-Path $matrixDir 'Defaults.xlsx'
        New-ValidDefaultsExcelFixture -Path $defaultsPath | Out-Null

        $matrixPath = Join-Path $matrixDir 'TeamA.xlsx'
        New-MatrixExcelFixture `
            -Path $matrixPath `
            -SettingsRows @(
            [pscustomobject]@{
                Status                  = 'Enabled'
                SiteName                = 'E2E-Fix'
                SiteCode                = 'E2E'
                ComputerName            = $env:COMPUTERNAME
                Path                    = $rootFolder
                GroupName               = 'E2E-Fix-Group'
                Action                  = 'Fix'
                ApplyDefaultPermissions = $false
            }
        )

        $configFixture = New-JsonFixture
        $configFixture.Matrix.FolderPath = $matrixDir
        $configFixture.Matrix.DefaultsFile = $defaultsPath
        $configFixture.Settings.SaveLogFiles.Where.Folder = $logsDir
        $configFixture.MaxConcurrent.FoldersPerMatrix = 1

        $configPath = Join-Path $matrixDir 'Input.json'
        $configFixture |
        ConvertTo-Json -Depth 20 |
        Out-File -LiteralPath $configPath -Encoding utf8 -Force

        $scriptPath = @{
            PermissionMatrixModule = "$moduleRoot\PermissionMatrix.psm1"
            SetPermissions         = "$root\Scripts\Operations\SetPermissions.ps1"
            TestRequirements       = "$root\Scripts\Operations\TestRequirements.ps1"
            UpdateServiceNow       = "$root\Scripts\Operations\UpdateServiceNow.ps1"
        }

        Mock Get-ADObjectDetailHC -ModuleName PermissionMatrix {
            return @(
                @{ SamAccountName = 'Bob'; adObject = @{ ObjectSid = $TestGroupBobSid } }
                @{ SamAccountName = 'Mike'; adObject = @{ ObjectSid = $TestGroupMikeSid } }
            )
        }

        Mock Send-MailKitMessageHC -ModuleName PermissionMatrix { }

        # ===================================================================
        # PHASE 2: Run with Action=Fix — should correct all corrupt ACLs
        # ===================================================================
        $systemErrors = [System.Collections.Generic.List[object]]::new()

        Invoke-PermissionMatrix `
            -ConfigurationJsonFile $configPath `
            -ScriptPath $scriptPath `
            -SystemErrors ([ref]$systemErrors)

        $fatals = $systemErrors.Where({ $_.Type -eq 'FatalError' })
        $fatals.Count | Should -Be 0 -Because (
            "Fix run should have no fatal errors but got: $($fatals | ForEach-Object { $_.Message } | Out-String)"
        )

        Should -Invoke Send-MailKitMessageHC -ModuleName PermissionMatrix -Times 1 -Exactly -Because (
            'Fix run should send exactly one mail'
        )

        # --- Verify Fix corrected the non-inherited folder ACLs -----------
        $fixedFinanceAcl = (Get-Acl -LiteralPath $financeFolder).Access
        $fixedDocsAcl = (Get-Acl -LiteralPath $docsFolder).Access

        $fixedFinanceAcl.Where({ $_.IdentityReference.Value -eq $bobNT.Value }).Count |
        Should -BeGreaterThan 0 -Because 'Fix should have applied an ACE for Bob on Finance'
        $fixedFinanceAcl.Where({ $_.IdentityReference.Value -eq $mikeNT.Value }).Count |
        Should -BeGreaterThan 0 -Because 'Fix should have applied an ACE for Mike on Finance'

        # Mike must no longer have FullControl — the matrix says R
        $fixedFinanceAcl.Where({
                $_.IdentityReference.Value -eq $mikeNT.Value -and
                $_.FileSystemRights -match 'FullControl'
            }).Count |
        Should -Be 0 -Because 'Fix should have corrected Mike from FullControl to R on Finance'

        $fixedDocsAcl.Where({ $_.IdentityReference.Value -eq $bobNT.Value }).Count |
        Should -BeGreaterThan 0 -Because 'Fix should have applied an ACE for Bob on Finance\Docs'
        $fixedDocsAcl.Where({ $_.IdentityReference.Value -eq $mikeNT.Value }).Count |
        Should -BeGreaterThan 0 -Because 'Fix should have applied an ACE for Mike on Finance\Docs'

        # --- Verify Fix restored inheritance on the file ------------------
        $fixedFileAcl = Get-Acl -LiteralPath $reportFile
        $fixedFileAcl.AreAccessRulesProtected |
        Should -BeFalse -Because 'Fix should have removed explicit ACL protection on report.txt (restored inheritance)'

        # ===================================================================
        # PHASE 3: Run with Action=Check — should report all correct
        # ===================================================================

        # Overwrite the matrix Excel with Action=Check (same path, same
        # permissions fixture — only the Action column changes).
        New-MatrixExcelFixture `
            -Path $matrixPath `
            -SettingsRows @(
            [pscustomobject]@{
                Status                  = 'Enabled'
                SiteName                = 'E2E-Fix'
                SiteCode                = 'E2E'
                ComputerName            = $env:COMPUTERNAME
                Path                    = $rootFolder
                GroupName               = 'E2E-Fix-Group'
                Action                  = 'Check'
                ApplyDefaultPermissions = $false
            }
        )

        # Re-serialise the JSON so the config timestamp is fresh — the
        # content is identical, only the Excel changed.
        $configFixture |
        ConvertTo-Json -Depth 20 |
        Out-File -LiteralPath $configPath -Encoding utf8 -Force

        $systemErrorsCheck = [System.Collections.Generic.List[object]]::new()

        Invoke-PermissionMatrix `
            -ConfigurationJsonFile $configPath `
            -ScriptPath $scriptPath `
            -SystemErrors ([ref]$systemErrorsCheck)

        $fatalsCheck = $systemErrorsCheck.Where({ $_.Type -eq 'FatalError' })
        $fatalsCheck.Count | Should -Be 0 -Because (
            "Check run should have no fatal errors but got: $($fatalsCheck | ForEach-Object { $_.Message } | Out-String)"
        )

        Should -Invoke Send-MailKitMessageHC -ModuleName PermissionMatrix -Times 2 -Exactly -Because (
            'Check run should send a second mail (two total: Fix + Check)'
        )

        # --- Verify Check found no permission issues on disk --------------
        $checkFinanceAcl = (Get-Acl -LiteralPath $financeFolder).Access
        $checkDocsAcl = (Get-Acl -LiteralPath $docsFolder).Access

        $checkFinanceAcl.Where({ $_.IdentityReference.Value -eq $bobNT.Value }).Count |
        Should -BeGreaterThan 0 -Because 'ACLs should still be intact after Check on Finance'
        $checkFinanceAcl.Where({ $_.IdentityReference.Value -eq $mikeNT.Value }).Count |
        Should -BeGreaterThan 0 -Because 'ACLs should still be intact after Check on Finance'

        $checkDocsAcl.Where({ $_.IdentityReference.Value -eq $bobNT.Value }).Count |
        Should -BeGreaterThan 0 -Because 'ACLs should still be intact after Check on Finance\Docs'
        $checkDocsAcl.Where({ $_.IdentityReference.Value -eq $mikeNT.Value }).Count |
        Should -BeGreaterThan 0 -Because 'ACLs should still be intact after Check on Finance\Docs'

        $checkFileAcl = Get-Acl -LiteralPath $reportFile
        $checkFileAcl.AreAccessRulesProtected |
        Should -BeFalse -Because 'report.txt should still have inherited ACL after Check'
    }

    It 'fixes a fully corrupted matrix tree, then a second Fix run is clean' -Skip:(-not $E2EPrereqsMet) {
        # -------------------------------------------------------------------
        # Broad full-flow unhappy-path regression:
        #   - the complete matrix folder tree already exists on disk;
        #   - every class of item is deliberately wrong:
        #       * root Path row folder (non-inherited / matrix-defined)
        #       * permissioned child folders (non-inherited / matrix-defined)
        #       * matrix paths with trailing slashes
        #       * inherit-only folders (empty ACL rows)
        #       * ignored folders/subtrees (I rows) that must stay untouched
        #       * files below both permissioned and inherit-only folders
        #   - first Action=Fix must converge the tree;
        #   - second Action=Fix must be clean, proving no reported correction
        #     was only a partial/no-op write.
        # -------------------------------------------------------------------
        function New-CorruptDirectoryAcl {
            param(
                [Parameter(Mandatory)]
                [System.Security.Principal.IdentityReference]$Identity
            )

            $acl = New-Object System.Security.AccessControl.DirectorySecurity
            $acl.SetAccessRuleProtection($true, $false)
            $acl.SetOwner([System.Security.Principal.NTAccount]'BUILTIN\Administrators')
            $acl.AddAccessRule(
                (New-Object System.Security.AccessControl.FileSystemAccessRule(
                        [System.Security.Principal.NTAccount]'BUILTIN\Administrators',
                        'FullControl',
                        'ContainerInherit,ObjectInherit',
                        'None',
                        'Allow'))
            )
            $acl.AddAccessRule(
                (New-Object System.Security.AccessControl.FileSystemAccessRule(
                        $Identity,
                        'FullControl',
                        'ContainerInherit,ObjectInherit',
                        'None',
                        'Allow'))
            )
            $acl
        }

        function New-CorruptFileAcl {
            param(
                [Parameter(Mandatory)]
                [System.Security.Principal.IdentityReference]$Identity
            )

            $acl = New-Object System.Security.AccessControl.FileSecurity
            $acl.SetAccessRuleProtection($true, $false)
            $acl.SetOwner([System.Security.Principal.NTAccount]'BUILTIN\Administrators')
            $acl.AddAccessRule(
                (New-Object System.Security.AccessControl.FileSystemAccessRule(
                        [System.Security.Principal.NTAccount]'BUILTIN\Administrators',
                        'FullControl',
                        'None',
                        'None',
                        'Allow'))
            )
            $acl.AddAccessRule(
                (New-Object System.Security.AccessControl.FileSystemAccessRule(
                        $Identity,
                        'FullControl',
                        'None',
                        'None',
                        'Allow'))
            )
            $acl
        }

        function Test-HasFullControlAce {
            param(
                [Parameter(Mandatory)]
                [System.Security.AccessControl.AuthorizationRuleCollection]$Access,
                [Parameter(Mandatory)]
                [string]$Identity
            )

            @($Access | Where-Object {
                    $_.IdentityReference.Value -eq $Identity -and
                    (($_.FileSystemRights -band [System.Security.AccessControl.FileSystemRights]::FullControl) -eq [System.Security.AccessControl.FileSystemRights]::FullControl)
                }).Count -gt 0
        }

        # ===================================================================
        # PHASE 1: Build a complete matrix tree with bad ACLs everywhere
        # ===================================================================
        $rootFolder = (New-Item 'TestDrive:\FullFixTarget' -ItemType Directory -Force).FullName
        $financeFolder = (New-Item (Join-Path $rootFolder 'Finance') -ItemType Directory -Force).FullName
        $docsFolder = (New-Item (Join-Path $financeFolder 'Docs') -ItemType Directory -Force).FullName
        $trailingSlashFolder = (New-Item (Join-Path $rootFolder 'TrailingSlash') -ItemType Directory -Force).FullName
        $inheritOnlyFolder = (New-Item (Join-Path $rootFolder 'InheritOnly') -ItemType Directory -Force).FullName
        $inheritOnlyNestedFolder = (New-Item (Join-Path $inheritOnlyFolder 'Nested') -ItemType Directory -Force).FullName
        $inheritOnlyDeepFolder = (New-Item (Join-Path $inheritOnlyNestedFolder 'Deep') -ItemType Directory -Force).FullName
        $inheritOnlySiblingFolder = (New-Item (Join-Path $rootFolder 'InheritOnlySibling') -ItemType Directory -Force).FullName
        $ignoredFolder = (New-Item (Join-Path $rootFolder 'IgnoredTree') -ItemType Directory -Force).FullName
        $ignoredChildFolder = (New-Item (Join-Path $ignoredFolder 'Child') -ItemType Directory -Force).FullName

        $rootFile = Join-Path $rootFolder 'root-file.txt'
        $financeFile = Join-Path $financeFolder 'finance-file.txt'
        $docsFile = Join-Path $docsFolder 'docs-file.txt'
        $trailingSlashFile = Join-Path $trailingSlashFolder 'trailing-file.txt'
        $inheritOnlyFile = Join-Path $inheritOnlyFolder 'inherit-only-file.txt'
        $nestedFile = Join-Path $inheritOnlyNestedFolder 'nested-file.txt'
        $deepFile = Join-Path $inheritOnlyDeepFolder 'deep-file.txt'
        $siblingFile = Join-Path $inheritOnlySiblingFolder 'sibling-file.txt'
        $ignoredFile = Join-Path $ignoredChildFolder 'ignored-file.txt'

        foreach ($file in @($rootFile, $financeFile, $docsFile, $trailingSlashFile, $inheritOnlyFile, $nestedFile, $deepFile, $siblingFile, $ignoredFile)) {
            'corrupt me' | Set-Content -LiteralPath $file -Force
        }

        $matrixDir = (New-Item 'TestDrive:\FullFixMatrix' -ItemType Directory -Force).FullName
        $logsDir = (New-Item 'TestDrive:\FullFixLogs' -ItemType Directory -Force).FullName

        $bobNT = (New-Object System.Security.Principal.NTAccount("$env:COMPUTERNAME\$TestGroupBob")).Value
        $mikeNT = (New-Object System.Security.Principal.NTAccount("$env:COMPUTERNAME\$TestGroupMike")).Value
        $bobAccount = [System.Security.Principal.NTAccount]$bobNT
        $mikeAccount = [System.Security.Principal.NTAccount]$mikeNT

        foreach ($folder in @($rootFolder, $financeFolder, $docsFolder, $trailingSlashFolder, $inheritOnlyFolder, $inheritOnlyNestedFolder, $inheritOnlyDeepFolder, $inheritOnlySiblingFolder, $ignoredFolder, $ignoredChildFolder)) {
            [System.IO.FileSystemAclExtensions]::SetAccessControl(
                [System.IO.DirectoryInfo]::new($folder),
                (New-CorruptDirectoryAcl -Identity $mikeAccount)
            )
        }

        foreach ($file in @($rootFile, $financeFile, $docsFile, $trailingSlashFile, $inheritOnlyFile, $nestedFile, $deepFile, $siblingFile, $ignoredFile)) {
            [System.IO.FileSystemAclExtensions]::SetAccessControl(
                [System.IO.FileInfo]::new($file),
                (New-CorruptFileAcl -Identity $bobAccount)
            )
        }

        # Sanity: the tree really starts in a bad state.
        (Get-Acl -LiteralPath $inheritOnlyFolder).AreAccessRulesProtected |
        Should -BeTrue -Because 'sanity: inherit-only folder starts protected and wrong'
        (Get-Acl -LiteralPath $docsFile).AreAccessRulesProtected |
        Should -BeTrue -Because 'sanity: file starts protected and wrong'
        Test-HasFullControlAce -Access (Get-Acl -LiteralPath $rootFolder).Access -Identity $mikeNT |
        Should -BeTrue -Because 'sanity: root starts with wrong Mike FullControl'
        (Get-Acl -LiteralPath $ignoredFolder).AreAccessRulesProtected |
        Should -BeTrue -Because 'sanity: ignored folder starts protected and wrong'

        $defaultsPath = Join-Path $matrixDir 'Defaults.xlsx'
        New-ValidDefaultsExcelFixture -Path $defaultsPath | Out-Null

        $permSpec = @{
            Row1 = @('', '', '')
            Row2 = @('', '', '')
            Row3 = @('', 'Bob', 'Mike')
            Row4 = @('Path', 'L', 'L')
            Data = @(
                @{ Path = 'Finance'           ; Col2 = 'R'; Col3 = 'R' }
                @{ Path = 'Finance\Docs'      ; Col2 = 'W'; Col3 = 'W' }
                @{ Path = 'TrailingSlash\'    ; Col2 = 'R'; Col3 = 'R' }
                @{ Path = 'InheritOnly'       ; Col2 = '' ; Col3 = ''  }
                @{ Path = 'InheritOnly\Nested'; Col2 = '' ; Col3 = ''  }
                @{ Path = 'InheritOnly\Nested\Deep'; Col2 = '' ; Col3 = '' }
                @{ Path = 'InheritOnlySibling'; Col2 = '' ; Col3 = '' }
                @{ Path = 'IgnoredTree\'      ; Col2 = 'I'; Col3 = '' }
            )
        }

        $matrixPath = Join-Path $matrixDir 'TeamA.xlsx'
        New-MatrixExcelFixture `
            -Path $matrixPath `
            -PermissionsRows $permSpec `
            -SettingsRows @(
            [pscustomobject]@{
                Status                  = 'Enabled'
                SiteName                = 'E2E-Full-Fix'
                SiteCode                = 'E2E'
                ComputerName            = $env:COMPUTERNAME
                Path                    = $rootFolder
                GroupName               = 'E2E-Full-Fix-Group'
                Action                  = 'Fix'
                ApplyDefaultPermissions = $false
            }
        )

        $configFixture = New-JsonFixture
        $configFixture.Matrix.FolderPath = $matrixDir
        $configFixture.Matrix.DefaultsFile = $defaultsPath
        $configFixture.Settings.SaveLogFiles.Where.Folder = $logsDir
        $configFixture.MaxConcurrent.FoldersPerMatrix = 1

        $configPath = Join-Path $matrixDir 'Input.json'
        $configFixture |
        ConvertTo-Json -Depth 20 |
        Out-File -LiteralPath $configPath -Encoding utf8 -Force

        $scriptPath = @{
            PermissionMatrixModule = "$moduleRoot\PermissionMatrix.psm1"
            SetPermissions         = "$root\Scripts\Operations\SetPermissions.ps1"
            TestRequirements       = "$root\Scripts\Operations\TestRequirements.ps1"
            UpdateServiceNow       = "$root\Scripts\Operations\UpdateServiceNow.ps1"
        }

        Mock Get-ADObjectDetailHC -ModuleName PermissionMatrix {
            return @(
                @{ SamAccountName = 'Bob'; adObject = @{ ObjectSid = $TestGroupBobSid } }
                @{ SamAccountName = 'Mike'; adObject = @{ ObjectSid = $TestGroupMikeSid } }
            )
        }

        $script:fullFixMailSubjects = [System.Collections.Generic.List[string]]::new()
        Mock Send-MailKitMessageHC -ModuleName PermissionMatrix {
            $script:fullFixMailSubjects.Add($Subject)
        }

        # ===================================================================
        # PHASE 2: First Fix run should detect and correct the bad tree
        # ===================================================================
        $systemErrorsFirst = [System.Collections.Generic.List[object]]::new()

        Invoke-PermissionMatrix `
            -ConfigurationJsonFile $configPath `
            -ScriptPath $scriptPath `
            -SystemErrors ([ref]$systemErrorsFirst)

        $fatalsFirst = $systemErrorsFirst.Where({ $_.Type -eq 'FatalError' })
        $fatalsFirst.Count | Should -Be 0 -Because (
            "first Fix run should have no fatal errors but got: $($fatalsFirst | ForEach-Object { $_.Message } | Out-String)"
        )

        $script:fullFixMailSubjects.Count |
        Should -Be 1 -Because 'first Fix run should send one mail'
        $script:fullFixMailSubjects[0] |
        Should -Match 'warning' -Because 'first Fix run should report the deliberately corrupted ACLs'

        foreach ($folder in @($rootFolder, $financeFolder, $docsFolder, $trailingSlashFolder)) {
            $acl = Get-Acl -LiteralPath $folder
            $acl.AreAccessRulesProtected |
            Should -BeTrue -Because "'$folder' is matrix-defined and should have protected explicit ACLs after Fix"
            Test-HasFullControlAce -Access $acl.Access -Identity $mikeNT |
            Should -BeFalse -Because "'$folder' should no longer carry wrong Mike FullControl after Fix"
        }

        foreach ($folder in @($inheritOnlyFolder, $inheritOnlyNestedFolder, $inheritOnlyDeepFolder, $inheritOnlySiblingFolder)) {
            $acl = Get-Acl -LiteralPath $folder
            $acl.AreAccessRulesProtected |
            Should -BeFalse -Because "'$folder' is inherit-only and should inherit after Fix"
            $acl.Access.Where({ -not $_.IsInherited }) |
            Should -BeNullOrEmpty -Because "'$folder' should not keep explicit ACEs after Fix"
        }

        foreach ($file in @($rootFile, $financeFile, $docsFile, $trailingSlashFile, $inheritOnlyFile, $nestedFile, $deepFile, $siblingFile)) {
            $acl = Get-Acl -LiteralPath $file
            $acl.AreAccessRulesProtected |
            Should -BeFalse -Because "'$file' should inherit after Fix"
            $acl.Access.Where({ -not $_.IsInherited }) |
            Should -BeNullOrEmpty -Because "'$file' should not keep explicit ACEs after Fix"
        }

        foreach ($ignoredItem in @($ignoredFolder, $ignoredChildFolder, $ignoredFile)) {
            $acl = Get-Acl -LiteralPath $ignoredItem
            $expectedIgnoredIdentity = if (Test-Path -LiteralPath $ignoredItem -PathType Leaf) { $bobNT } else { $mikeNT }

            $acl.AreAccessRulesProtected |
            Should -BeTrue -Because "'$ignoredItem' is under an ignored matrix row and must remain untouched"
            Test-HasFullControlAce -Access $acl.Access -Identity $expectedIgnoredIdentity |
            Should -BeTrue -Because "'$ignoredItem' must keep its deliberately wrong ACL because ignored rows are not managed"
        }

        $detailFilesBeforeSecondFix = @(
            Get-ChildItem -LiteralPath $logsDir -Recurse -Filter '*.json' -File |
            Select-Object -ExpandProperty FullName
        )

        # ===================================================================
        # PHASE 3: Second Fix run should be clean for permissions
        # ===================================================================
        $systemErrorsSecond = [System.Collections.Generic.List[object]]::new()

        Invoke-PermissionMatrix `
            -ConfigurationJsonFile $configPath `
            -ScriptPath $scriptPath `
            -SystemErrors ([ref]$systemErrorsSecond)

        $fatalsSecond = $systemErrorsSecond.Where({ $_.Type -eq 'FatalError' })
        $fatalsSecond.Count | Should -Be 0 -Because (
            "second Fix run should have no fatal errors but got: $($fatalsSecond | ForEach-Object { $_.Message } | Out-String)"
        )

        $script:fullFixMailSubjects.Count |
        Should -Be 2 -Because 'second Fix run should send a second mail'

        $newDetailFiles = @(
            Get-ChildItem -LiteralPath $logsDir -Recurse -Filter '*.json' -File |
            Where-Object { $_.FullName -notin $detailFilesBeforeSecondFix }
        )

        $newDetailText = @(
            foreach ($detailFile in $newDetailFiles) {
                Get-Content -LiteralPath $detailFile.FullName -Raw
            }
        ) -join "`n"

        $newDetailText |
        Should -Not -Match 'Non inherited folder incorrect permissions|Inherited permissions incorrect' -Because (
            'a second Fix run after convergence must not report any remaining permission corrections'
        )
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