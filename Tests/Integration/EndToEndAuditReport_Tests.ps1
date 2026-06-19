#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

# =============================================================================
# End-to-end test for the Permission Matrix AUDIT REPORT.
#
# Approach: invoke the public 'Invoke-PermissionMatrixAuditReport' directly
# (bypassing the entry-point script). Mock at the AD and SMTP boundaries;
# everything else is real (Excel import, AD resolution flow, log-file writing,
# mail assembly).
#
# Unlike the main pipeline e2e test, this one needs NO Administrator rights, NO
# PSRemoting and NO 'PowerShell.7' session config: the audit only runs the
# Begin stage and then reads/reports — it never applies ACLs or remotes.
#
# Dependency note: the audit reads the matrix 'FormData' worksheet, which the
# import stage only loads when the config contains an 'AuditReport' section.
# That gate lives in Import-MatrixFileHC ('-or $Context.Config.AuditReport').
# If FormData is not imported, matrices look like they have no responsible and
# no mail is sent — so this test also guards that wiring.
# =============================================================================

Describe 'Permission Matrix Audit Report - End to End' {
    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\.."
        $script:moduleRoot = "$root\Modules\PermissionMatrix"

        # Project fixture helpers (same ones the pipeline e2e test uses).
        . "$root\Tests\Helpers\Fixtures.Json.ps1"
        . "$root\Tests\Helpers\Fixtures.Excel.ps1"

        Import-Module "$moduleRoot\PermissionMatrix.psm1" -Force

        # The audit runs the Begin stage only, so the single 'PermissionMatrixModule'
        # key is all that is needed (Begin uses it to dot-source the private
        # functions inside its runspaces).
        $script:scriptPath = @{
            PermissionMatrixModule = "$moduleRoot\PermissionMatrix.psm1"
        }

        # ------------------------------------------------------------------
        # Helper: append a valid 'FormData' worksheet (the audit needs it).
        # Mandatory columns per Test-MatrixFormDataHC.
        # ------------------------------------------------------------------
        function Add-FormDataSheet {
            param(
                [string]$Path,
                [string]$Responsible,
                [string]$FolderPath
            )
            @(
                [pscustomobject]@{
                    MatrixFormStatus        = 'Enabled'
                    MatrixCategoryName      = 'Finance'
                    MatrixSubCategoryName   = 'Accounting'
                    MatrixResponsible       = $Responsible
                    MatrixFolderDisplayName = 'Finance share'
                    MatrixFolderPath        = $FolderPath
                }
            ) | Export-Excel -Path $Path -WorksheetName 'FormData'
        }

        # ------------------------------------------------------------------
        # Helper: build a SLIM audit config - only the fields the audit itself
        # uses. It deliberately OMITS the schema-only blocks the shared
        # validator expects (Export, ServiceNow, PSSessionConfiguration,
        # MaxConcurrent, Matrix.Archive, Settings.SaveLogFiles.Detailed). The
        # audit fills those with defaults in memory before Begin validates, so
        # this test exercises that skeleton-injection path end to end: if it
        # broke, Begin would raise "Missing/Incorrect" fatal errors and no owner
        # mail would be sent.
        # ------------------------------------------------------------------
        function New-AuditConfig {
            param(
                [string]$MatrixDir,
                [string]$DefaultsPath,
                [string]$LogsDir,
                [string[]]$ScriptAdmin
            )
            return [ordered]@{
                Matrix      = [ordered]@{
                    FolderPath          = $MatrixDir
                    DefaultsFile        = $DefaultsPath
                    AdGroupPlaceHolders = @()
                }
                AuditReport = [ordered]@{
                    RequestTicketURL = 'https://portal/req'
                    ScriptAdmin      = $ScriptAdmin
                }
                Settings    = [ordered]@{
                    ScriptName     = 'Permission matrix audit report (test)'
                    SendMail       = [ordered]@{
                        From            = 'no-reply@example.com'
                        FromDisplayName = 'Audit'
                        To              = @()
                        Bcc             = @()
                        Subject         = 'Audit {{MatrixFileName}}: {{UniqueUserCount}} users, {{UniqueGroupCount}} groups'
                        Body            = '<p>Please review {{MatrixFileName}} at {{RequestTicketURL}}</p>'
                        Smtp            = [ordered]@{
                            ServerName     = 'smtp.example.com'
                            Port           = 25
                            ConnectionType = 'None'
                            UserName       = ''
                            Password       = ''
                        }
                        AssemblyPath    = [ordered]@{
                            MailKit = 'C:\MailKit.dll'
                            MimeKit = 'C:\MimeKit.dll'
                        }
                    }
                    SaveLogFiles   = [ordered]@{
                        Where               = [ordered]@{ Folder = $LogsDir }
                        DeleteLogsAfterDays = 30
                    }
                    SaveInEventLog = [ordered]@{ Save = $false; LogName = 'Scripts' }
                }
            }
        }
    }

    It 'e-mails the matrix responsible with the log file attached and reports no errors' {
        $targetFolder = Join-Path $TestDrive 'Target'
        $matrixDir = (New-Item 'TestDrive:\Matrix' -ItemType Directory -Force).FullName
        $logsDir = (New-Item 'TestDrive:\Logs' -ItemType Directory -Force).FullName

        $defaultsPath = Join-Path $matrixDir 'Defaults.xlsx'
        New-ValidDefaultsExcelFixture -Path $defaultsPath | Out-Null

        # Matrix (Settings + Permissions via the 'Valid' fixture, which
        # references Bob and Mike). Action=Check: the audit ignores Action,
        # it never applies anything.
        $matrixPath = Join-Path $matrixDir 'TeamA.xlsx'
        New-MatrixExcelFixture -Path $matrixPath -SettingsRows @(
            [pscustomobject]@{
                Status                  = 'Enabled'
                SiteName                = 'E2E'
                SiteCode                = 'E2E'
                ComputerName            = $env:COMPUTERNAME
                Path                    = $targetFolder
                GroupName               = 'E2E-Test-Group'
                Action                  = 'Check'
                ApplyDefaultPermissions = $false
            }
        )
        Add-FormDataSheet -Path $matrixPath -Responsible 'owner@example.com' -FolderPath $targetFolder

        $configPath = Join-Path $matrixDir 'AuditInput.json'
        New-AuditConfig -MatrixDir $matrixDir -DefaultsPath $defaultsPath `
            -LogsDir $logsDir -ScriptAdmin @('audit-admin@example.com') |
        ConvertTo-Json -Depth 20 |
        Out-File -LiteralPath $configPath -Encoding utf8 -Force

        Mock Get-ADObjectDetailHC -ModuleName PermissionMatrix {
            return @(
                @{ SamAccountName = 'Bob'; adObject = @{ ObjectSid = 'S-1-5-32-544' } }
                @{ SamAccountName = 'Mike'; adObject = @{ ObjectSid = 'S-1-5-32-545' } }
            )
        }
        Mock Send-MailKitMessageHC -ModuleName PermissionMatrix { }

        $systemErrors = [System.Collections.Generic.List[object]]::new()

        Invoke-PermissionMatrixAuditReport `
            -ConfigurationJsonFile $configPath `
            -ScriptPath $scriptPath `
            -SystemErrors ([ref]$systemErrors)

        # No fatal errors during initialization.
        $fatals = $systemErrors.Where({ $_.Type -eq 'FatalError' })
        $fatals.Count | Should -Be 0 -Because (
            "expected no fatal errors but got: $($fatals | ForEach-Object { $_.Message } | Out-String)"
        )

        # Exactly one mail was sent.
        Should -Invoke Send-MailKitMessageHC -ModuleName PermissionMatrix -Times 1 -Exactly

        # It is the owner audit mail, not an admin skip / init-failure mail.
        Should -Invoke Send-MailKitMessageHC -ModuleName PermissionMatrix -ParameterFilter {
            ($Subject -notmatch 'skipped') -and ($Subject -notmatch 'initialization')
        } -Times 1 -Exactly -Because 'the valid matrix should mail its responsible, not the admin'

        # The responsible is the recipient (lenient: To may be a string or array).
        Should -Invoke Send-MailKitMessageHC -ModuleName PermissionMatrix -ParameterFilter {
            (@($To) -join ';') -match 'owner@example\.com'
        } -Because 'the responsible should be the recipient'

        # The per-matrix log file is attached and exists on disk.
        Should -Invoke Send-MailKitMessageHC -ModuleName PermissionMatrix -ParameterFilter {
            $att = @($Attachments) | Where-Object { $_ } | Select-Object -First 1
            $att -and (Test-Path -LiteralPath $att)
        } -Because 'the responsible should be mailed their log file as an attachment'

        # The admin is BCC'd.
        Should -Invoke Send-MailKitMessageHC -ModuleName PermissionMatrix -ParameterFilter {
            (@($Bcc) -join ';') -match 'audit-admin@example\.com'
        } -Because 'the admin should be BCC''d on the audit mail'
    }

    It 'skips matrices with fatal errors and reports them to the admin' {
        $targetFolder = Join-Path $TestDrive 'Target'
        $matrixDir = (New-Item 'TestDrive:\Matrix' -ItemType Directory -Force).FullName
        $logsDir = (New-Item 'TestDrive:\Logs' -ItemType Directory -Force).FullName

        $defaultsPath = Join-Path $matrixDir 'Defaults.xlsx'
        New-ValidDefaultsExcelFixture -Path $defaultsPath | Out-Null

        # Two matrices with the SAME ComputerName + Path -> Begin flags both
        # with a 'Duplicate ComputerName/Path' fatal error. Both are skipped
        # and reported to the admin; neither responsible is mailed.
        foreach ($name in 'TeamA', 'TeamB') {
            $p = Join-Path $matrixDir "$name.xlsx"
            New-MatrixExcelFixture -Path $p -SettingsRows @(
                [pscustomobject]@{
                    Status                  = 'Enabled'
                    SiteName                = 'E2E'
                    SiteCode                = 'E2E'
                    ComputerName            = $env:COMPUTERNAME
                    Path                    = $targetFolder
                    GroupName               = 'E2E-Test-Group'
                    Action                  = 'Check'
                    ApplyDefaultPermissions = $false
                }
            )
            Add-FormDataSheet -Path $p -Responsible "$name@example.com" -FolderPath $targetFolder
        }

        $configPath = Join-Path $matrixDir 'AuditInput.json'
        New-AuditConfig -MatrixDir $matrixDir -DefaultsPath $defaultsPath `
            -LogsDir $logsDir -ScriptAdmin @('audit-admin@example.com') |
        ConvertTo-Json -Depth 20 |
        Out-File -LiteralPath $configPath -Encoding utf8 -Force

        Mock Get-ADObjectDetailHC -ModuleName PermissionMatrix {
            return @(
                @{ SamAccountName = 'Bob'; adObject = @{ ObjectSid = 'S-1-5-32-544' } }
                @{ SamAccountName = 'Mike'; adObject = @{ ObjectSid = 'S-1-5-32-545' } }
            )
        }
        Mock Send-MailKitMessageHC -ModuleName PermissionMatrix { }

        $systemErrors = [System.Collections.Generic.List[object]]::new()

        Invoke-PermissionMatrixAuditReport `
            -ConfigurationJsonFile $configPath `
            -ScriptPath $scriptPath `
            -SystemErrors ([ref]$systemErrors)

        # Exactly one mail: the admin summary. No owner mails for skipped files.
        Should -Invoke Send-MailKitMessageHC -ModuleName PermissionMatrix -Times 1 -Exactly

        Should -Invoke Send-MailKitMessageHC -ModuleName PermissionMatrix -ParameterFilter {
            ($To -contains 'audit-admin@example.com') -and
            ($Subject -match 'skipped')
        } -Because 'the admin should receive the skipped-matrix summary'

        Should -Invoke Send-MailKitMessageHC -ModuleName PermissionMatrix -ParameterFilter {
            ($To -contains 'TeamA@example.com') -or ($To -contains 'TeamB@example.com')
        } -Times 0 -Because 'responsibles of skipped matrices must not be mailed'
    }
}