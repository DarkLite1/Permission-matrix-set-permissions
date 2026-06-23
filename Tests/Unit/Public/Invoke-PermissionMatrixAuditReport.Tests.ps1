#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

Describe 'Invoke-PermissionMatrixAuditReport' {
    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        . "$root\Tests\Helpers\Helpers.HC.ps1"

        Get-ChildItem "$moduleRoot\Private" -Filter '*.ps1' -File |
        ForEach-Object { . $_.FullName }

        Get-ChildItem "$moduleRoot\Public" -Filter '*.ps1' -File -ErrorAction SilentlyContinue |
        ForEach-Object { . $_.FullName }

        # Defensive: the function under test and the mail builder may live in
        # Public, Private, or a sibling folder depending on the layout. Make sure
        # both are loaded so the call under test resolves and the mock targets a
        # known command.
        foreach ($fn in 'Invoke-PermissionMatrixAuditReport', 'Build-AuditReportMailHC') {
            if (-not (Get-Command $fn -ErrorAction SilentlyContinue)) {
                Get-ChildItem $moduleRoot -Recurse -Filter "$fn.ps1" -File |
                Select-Object -First 1 |
                ForEach-Object { . $_.FullName }
            }
        }

        #region Factory helpers
        function New-AuditConfig {
            param(
                [string[]]$To = @(),
                [string[]]$Bcc = @('cfgbcc@example.com'),
                [string[]]$ScriptAdmin = @('admin@example.com'),
                [string]$From = 'no-reply@example.com',
                [string]$RequestTicketURL = 'https://tickets.example.com',
                [string]$ScriptName = 'Permission matrix audit report',
                [string]$LogFolder = 'TestDrive:\AuditLogs',
                [string[]]$PlaceHolders = @('cnorris')
            )

            [PSCustomObject]@{
                Matrix      = [PSCustomObject]@{ AdGroupPlaceHolders = $PlaceHolders }
                AuditReport = [PSCustomObject]@{
                    RequestTicketURL = $RequestTicketURL
                    ScriptAdmin      = $ScriptAdmin
                }
                Settings    = [PSCustomObject]@{
                    ScriptName   = $ScriptName
                    SendMail     = [PSCustomObject]@{
                        From            = $From
                        FromDisplayName = 'Audit'
                        To              = $To
                        Bcc             = $Bcc
                        Subject         = 'Access review - {{MatrixFileName}}'
                        Body            = '<p>{{MatrixResponsible}}</p>'
                        Smtp            = [PSCustomObject]@{
                            ServerName     = 'SMTP1'
                            Port           = 25
                            ConnectionType = 'None'
                            UserName       = ''
                            Password       = ''
                        }
                        AssemblyPath    = [PSCustomObject]@{
                            MailKit = 'C:\MailKit.dll'
                            MimeKit = 'C:\MimeKit.dll'
                        }
                    }
                    SaveLogFiles = [PSCustomObject]@{
                        Where = [PSCustomObject]@{ Folder = $LogFolder }
                    }
                }
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

        function New-AuditFileResult {
            param(
                [string]$Name = 'Matrix.xlsx',
                [string]$Responsible = 'jdoe',
                [switch]$NoFormData,
                [switch]$NoResponsible,
                [pscustomobject[]]$Check = @(),
                [pscustomobject[]]$MatrixChecks = @(),
                [string]$LogFolder
            )

            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Name)

            $formData =
            if ($NoFormData) { $null }
            elseif ($NoResponsible) {
                [PSCustomObject]@{
                    MatrixFileName    = $baseName
                    MatrixResponsible = $null
                }
            }
            else {
                [PSCustomObject]@{
                    MatrixFileName          = $baseName
                    MatrixResponsible       = $Responsible
                    MatrixFilePath          = 'C:\Matrix.xlsx'
                    MatrixCategoryName      = 'Cat'
                    MatrixSubCategoryName   = 'Sub'
                    MatrixFolderPath        = 'C:\Folder'
                    MatrixFolderDisplayName = 'Folder'
                }
            }

            $fr = [PSCustomObject]@{
                Item     = [PSCustomObject]@{
                    Name     = $Name
                    BaseName = $baseName
                    FullName = "C:\src\$Name"
                }
                Check    = [System.Collections.Generic.List[pscustomobject]]($Check)
                Matrices = @(
                    $MatrixChecks | ForEach-Object {
                        [PSCustomObject]@{ Check = @($_) }
                    }
                )
                Sheets   = [PSCustomObject]@{
                    FormData = [PSCustomObject]@{ Formatted = $formData }
                }
            }

            if ($LogFolder) {
                $fr | Add-Member -NotePropertyName LogFolder -NotePropertyValue $LogFolder
            }

            $fr
        }

        function New-AuditContext {
            param(
                $Config = (New-AuditConfig),
                [array]$FileResults = @(),
                [bool]$FoundMatrices = $true,
                $AdObjectDetails = @{}
            )

            [PSCustomObject]@{
                Config          = $Config
                FileResults     = $FileResults
                FoundMatrices   = $FoundMatrices
                AdObjectDetails = $AdObjectDetails
            }
        }
        #endregion
    }

    BeforeEach {
        $script:systemErrors = [System.Collections.Generic.List[pscustomobject]]::new()

        # No on-disk config: the function reads it only to fill schema defaults
        # and falls back gracefully when it cannot be parsed. The mocked Begin
        # stage supplies the real configuration through the returned context.
        $script:configFile = 'TestDrive:\audit.json'
        $script:scriptPath = @{ PermissionMatrixModule = 'TestDrive:\mod.psm1' }

        # Default: one healthy matrix, per-matrix routing (SendMail.To empty)
        $script:auditContext = New-AuditContext -FileResults @(
            New-AuditFileResult -Name 'Matrix.xlsx' -Responsible 'jdoe'
        )

        Mock Invoke-PermissionMatrixBeginHC { $script:auditContext }

        Mock Resolve-ResponsibleEmailHC {
            [PSCustomObject]@{
                Emails     = @('jdoe@example.com')
                Unresolved = @()
            }
        }

        Mock Build-MatrixLogSheetRowsHC {
            [PSCustomObject]@{
                AccessList    = @()
                GroupManagers = @()
                AdObjects     = @()
            }
        }

        Mock Copy-MatrixFileToLogFolderHC { 'TestDrive:\AuditLogs\Matrix\Matrix.xlsx' }

        # Echo back the recipient the orchestrator decided on (To), and a
        # sentinel Bcc that represents "whatever Build computed". This lets the
        # tests assert both the recipient decision (Build inputs) and the
        # orchestrator's post-build Bcc handling (Send inputs) independently.
        Mock Build-AuditReportMailHC {
            @{
                From            = 'no-reply@example.com'
                FromDisplayName = 'Audit'
                To              = @($RecipientEmail)
                Bcc             = @('built-bcc@example.com')
                Subject         = 'Access review - Matrix'
                Body            = '<p>body</p>'
                Attachments     = 'TestDrive:\AuditLogs\Matrix\Matrix.xlsx'
            }
        }

        Mock Send-MailKitMessageHC { }

        # Keep the in-function mail-parameter builder deterministic and free of
        # filesystem / environment concerns. These mocks have no param block:
        # Pester exposes each bound parameter as a variable in the mock body, so
        # we read the real parameter names ('Name', 'Hashtable') directly. A
        # param block would force Pester to re-bind by name and throw if our
        # names did not match the real command's.
        Mock Get-StringValueHC { $Name }
        Mock Remove-BlankValueHC {
            foreach ($k in @($Hashtable.Keys)) {
                $v = $Hashtable[$k]
                if ($null -eq $v -or ($v -is [string] -and $v -eq '')) {
                    $Hashtable.Remove($k)
                }
            }
            $Hashtable
        }

        # Avoid touching the real filesystem for the log folder / mail html.
        Mock New-Item { }
        Mock Set-Content { }
    }

    Context 'Guard conditions' {
        It 'sends nothing when the Begin stage returns no context' {
            $script:auditContext = $null

            Invoke-PermissionMatrixAuditReport `
                -ConfigurationJsonFile $configFile `
                -ScriptPath $scriptPath `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Send-MailKitMessageHC -Times 0 -Exactly
        }

        It 'sends nothing when no matrices were found' {
            $script:auditContext = New-AuditContext -FoundMatrices $false -FileResults @(
                New-AuditFileResult
            )

            Invoke-PermissionMatrixAuditReport `
                -ConfigurationJsonFile $configFile `
                -ScriptPath $scriptPath `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Send-MailKitMessageHC -Times 0 -Exactly
        }

        It 'sends nothing when there are no file results' {
            $script:auditContext = New-AuditContext -FileResults @()

            Invoke-PermissionMatrixAuditReport `
                -ConfigurationJsonFile $configFile `
                -ScriptPath $scriptPath `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Send-MailKitMessageHC -Times 0 -Exactly
        }

        It 'mails the admin when initialization produced a FatalError' {
            # Begin appends initialization errors to the [ref] collection; a
            # FatalError there must stop the run and notify the admin.
            $systemErrors.Add((New-FatalCheck -Name 'Init' -Description 'boom'))

            Invoke-PermissionMatrixAuditReport `
                -ConfigurationJsonFile $configFile `
                -ScriptPath $scriptPath `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Send-MailKitMessageHC -Times 1 -Exactly -ParameterFilter {
                $To -contains 'admin@example.com' -and
                $Subject -like '*initialization failed*'
            }
            # The per-matrix audit work must not run.
            Should -Invoke Resolve-ResponsibleEmailHC -Times 0 -Exactly
        }
    }

    Context 'Per-matrix routing (SendMail.To empty)' {
        It 'sends the audit mail to the resolved responsible address' {
            Invoke-PermissionMatrixAuditReport `
                -ConfigurationJsonFile $configFile `
                -ScriptPath $scriptPath `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Resolve-ResponsibleEmailHC -Times 1 -Exactly -ParameterFilter {
                $Responsible -eq 'jdoe'
            }
            Should -Invoke Build-AuditReportMailHC -Times 1 -Exactly -ParameterFilter {
                $RecipientEmail -contains 'jdoe@example.com'
            }
            Should -Invoke Send-MailKitMessageHC -Times 1 -Exactly -ParameterFilter {
                $To -contains 'jdoe@example.com' -and $Subject -like 'Access review*'
            }
        }

        It 'BCCs the script admin on the audit mail' {
            Invoke-PermissionMatrixAuditReport `
                -ConfigurationJsonFile $configFile `
                -ScriptPath $scriptPath `
                -SystemErrors ([ref]$systemErrors)

            # The admin is handed to the builder as the extra Bcc...
            Should -Invoke Build-AuditReportMailHC -Times 1 -Exactly -ParameterFilter {
                $Bcc -contains 'admin@example.com'
            }
            # ...and the message Bcc survives all the way to the send (not cleared).
            Should -Invoke Send-MailKitMessageHC -Times 1 -Exactly -ParameterFilter {
                $Bcc -contains 'built-bcc@example.com'
            }
        }

        It 'skips a matrix with a file-level FatalError and reports it to the admin' {
            $script:auditContext = New-AuditContext -FileResults @(
                New-AuditFileResult -Name 'Broken.xlsx' -Check @((New-FatalCheck -Name 'Bad data'))
            )

            Invoke-PermissionMatrixAuditReport `
                -ConfigurationJsonFile $configFile `
                -ScriptPath $scriptPath `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Resolve-ResponsibleEmailHC -Times 0 -Exactly
            Should -Invoke Send-MailKitMessageHC -Times 0 -Exactly -ParameterFilter {
                $Subject -like 'Access review*'
            }
            Should -Invoke Send-MailKitMessageHC -Times 1 -Exactly -ParameterFilter {
                $To -contains 'admin@example.com' -and $Subject -like '*skipped*'
            }
        }

        It 'skips a matrix that has no MatrixResponsible and sends no mail' {
            $script:auditContext = New-AuditContext -FileResults @(
                New-AuditFileResult -NoResponsible
            )

            Invoke-PermissionMatrixAuditReport `
                -ConfigurationJsonFile $configFile `
                -ScriptPath $scriptPath `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Resolve-ResponsibleEmailHC -Times 0 -Exactly
            Should -Invoke Send-MailKitMessageHC -Times 0 -Exactly
        }

        It 'sends no audit mail and reports the admin when the responsible resolves to no e-mail' {
            Mock Resolve-ResponsibleEmailHC {
                [PSCustomObject]@{ Emails = @(); Unresolved = @() }
            }

            Invoke-PermissionMatrixAuditReport `
                -ConfigurationJsonFile $configFile `
                -ScriptPath $scriptPath `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Send-MailKitMessageHC -Times 0 -Exactly -ParameterFilter {
                $Subject -like 'Access review*'
            }
            Should -Invoke Send-MailKitMessageHC -Times 1 -Exactly -ParameterFilter {
                $To -contains 'admin@example.com' -and $Subject -like '*without an e-mail address*'
            }
        }

        It 'still mails the resolved recipients but reports unresolved members to the admin' {
            Mock Resolve-ResponsibleEmailHC {
                [PSCustomObject]@{
                    Emails     = @('jdoe@example.com')
                    Unresolved = @('GROUP-NoMail')
                }
            }

            Invoke-PermissionMatrixAuditReport `
                -ConfigurationJsonFile $configFile `
                -ScriptPath $scriptPath `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Send-MailKitMessageHC -Times 1 -Exactly -ParameterFilter {
                $To -contains 'jdoe@example.com' -and $Subject -like 'Access review*'
            }
            Should -Invoke Send-MailKitMessageHC -Times 1 -Exactly -ParameterFilter {
                $To -contains 'admin@example.com' -and $Subject -like '*without an e-mail address*'
            }
        }
    }

    Context 'SendMail.To override' {
        It 'redirects every matrix to SendMail.To and ignores the responsible' {
            $script:auditContext = New-AuditContext `
                -Config (New-AuditConfig -To @('test@example.com')) `
                -FileResults @( New-AuditFileResult -Responsible 'jdoe' )

            Invoke-PermissionMatrixAuditReport `
                -ConfigurationJsonFile $configFile `
                -ScriptPath $scriptPath `
                -SystemErrors ([ref]$systemErrors)

            # The Excel-derived responsible is never resolved in override mode.
            Should -Invoke Resolve-ResponsibleEmailHC -Times 0 -Exactly
            Should -Invoke Build-AuditReportMailHC -Times 1 -Exactly -ParameterFilter {
                $RecipientEmail -contains 'test@example.com'
            }
            Should -Invoke Send-MailKitMessageHC -Times 1 -Exactly -ParameterFilter {
                $To -contains 'test@example.com' -and $To -notcontains 'jdoe@example.com'
            }
        }

        It 'adds no Bcc in override mode even when config Bcc and ScriptAdmin are set' {
            $script:auditContext = New-AuditContext `
                -Config (New-AuditConfig `
                    -To @('test@example.com') `
                    -Bcc @('cfgbcc@example.com') `
                    -ScriptAdmin @('admin@example.com')) `
                -FileResults @( New-AuditFileResult )

            Invoke-PermissionMatrixAuditReport `
                -ConfigurationJsonFile $configFile `
                -ScriptPath $scriptPath `
                -SystemErrors ([ref]$systemErrors)

            # No admin handed to the builder...
            Should -Invoke Build-AuditReportMailHC -Times 1 -Exactly -ParameterFilter {
                -not $Bcc
            }
            # ...and the message Bcc is cleared, so the send carries no Bcc at all.
            Should -Invoke Send-MailKitMessageHC -Times 1 -Exactly -ParameterFilter {
                -not $Bcc
            }
        }

        It 'still sends for a matrix whose responsible is empty (skip bypassed)' {
            $script:auditContext = New-AuditContext `
                -Config (New-AuditConfig -To @('test@example.com')) `
                -FileResults @( New-AuditFileResult -NoResponsible )

            Invoke-PermissionMatrixAuditReport `
                -ConfigurationJsonFile $configFile `
                -ScriptPath $scriptPath `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Resolve-ResponsibleEmailHC -Times 0 -Exactly
            Should -Invoke Send-MailKitMessageHC -Times 1 -Exactly -ParameterFilter {
                $To -contains 'test@example.com'
            }
        }

        It 'skips a matrix that has no FormData at all, even in override mode' {
            $script:auditContext = New-AuditContext `
                -Config (New-AuditConfig -To @('test@example.com')) `
                -FileResults @( New-AuditFileResult -NoFormData )

            Invoke-PermissionMatrixAuditReport `
                -ConfigurationJsonFile $configFile `
                -ScriptPath $scriptPath `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Send-MailKitMessageHC -Times 0 -Exactly
        }

        It 'routes all matrices to the override address when several are present' {
            $script:auditContext = New-AuditContext `
                -Config (New-AuditConfig -To @('test@example.com')) `
                -FileResults @(
                    New-AuditFileResult -Name 'A.xlsx' -Responsible 'alice'
                    New-AuditFileResult -Name 'B.xlsx' -Responsible 'bob'
                )

            Invoke-PermissionMatrixAuditReport `
                -ConfigurationJsonFile $configFile `
                -ScriptPath $scriptPath `
                -SystemErrors ([ref]$systemErrors)

            Should -Invoke Send-MailKitMessageHC -Times 2 -Exactly -ParameterFilter {
                $To -contains 'test@example.com'
            }
        }
    }
}