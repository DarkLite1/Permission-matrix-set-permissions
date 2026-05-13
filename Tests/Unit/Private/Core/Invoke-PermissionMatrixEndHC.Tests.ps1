#Requires -Version 7
#Requires -Modules Pester

Describe 'Invoke-PermissionMatrixEndHC' {
    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        . "$moduleRoot\Private\Utils.ps1"
        . "$moduleRoot\Private\Html.ps1"
        . "$moduleRoot\Private\Mail.ps1"
        . "$moduleRoot\Private\Export.ps1"
        . "$moduleRoot\Private\Logging\Write-EventLogSafe.ps1"
        . "$moduleRoot\Private\Logging\Cleanup-OldLogsHC.ps1"
        . "$moduleRoot\Private\Logging\Write-SystemErrorLogHC.ps1"
        . "$moduleRoot\Private\Core\Invoke-PermissionMatrixEndHC.ps1"

        function New-EndContext {
            param(
                [hashtable]$Counter = @{},
                [array]$FileResults = @(),
                [array]$AllMatrices = @(),
                [bool]$FoundMatrices = $true,
                [string]$LogFolder = 'TestDrive:\Logs',
                [bool]$Archive = $false,
                [string]$JsonFileName = 'TestInput',
                [hashtable]$ServiceNow = @{},
                [hashtable]$Defaults = @{ MailTo = @() },
                [hashtable]$SaveLogFiles = @{
                    Detailed            = $false
                    DeleteLogsAfterDays = 0
                    Where               = @{ Folder = 'TestDrive:\Logs' }
                },
                [hashtable]$SaveInEventLog = @{ Save = $false; LogName = 'Application' },
                [hashtable]$SendMail = @{
                    To           = @('test@example.com')
                    From         = 'noreply@example.com'
                    AssemblyPath = @{
                        MailKit = 'TestDrive:\fake-mailkit.dll'
                        MimeKit = 'TestDrive:\fake-mimekit.dll'
                    }
                    Smtp         = @{
                        ServerName = 'smtp.example.com'
                        Port       = 25
                    }
                }, 
                [hashtable]$Export = @{},
                [hashtable]$ScriptPath = @{ UpdateServiceNow = 'TestDrive:\Snow.ps1' },
                [string]$ScriptName = 'Permission Matrix'
            )

            $SaveLogFiles.Where.Folder = $LogFolder

            [PSCustomObject]@{
                Counter       = $Counter
                FileResults   = $FileResults
                AllMatrices   = $AllMatrices
                FoundMatrices = $FoundMatrices
                StartTime     = (Get-Date).AddMinutes(-5)
                JsonFileName  = $JsonFileName
                Defaults      = [PSCustomObject]$Defaults
                ScriptPath    = $ScriptPath
                ExportedFiles = $null
                Config        = [PSCustomObject]@{
                    Settings   = [PSCustomObject]@{
                        ScriptName     = $ScriptName
                        SendMail       = $SendMail
                        SaveLogFiles   = [PSCustomObject]$SaveLogFiles
                        SaveInEventLog = [PSCustomObject]$SaveInEventLog
                    }
                    Export     = [PSCustomObject]$Export
                    ServiceNow = [PSCustomObject]$ServiceNow
                }
            }
        }

        function New-EndMatrix {
            param([string]$Name = 'TestMatrix', [pscustomobject[]]$Check = @())
            [PSCustomObject]@{
                ID    = [guid]::NewGuid().ToString()
                Check = [System.Collections.Generic.List[pscustomobject]]@($Check)
                Item  = [PSCustomObject]@{ BaseName = $Name; Name = "$Name.xlsx" }
            }
        }

        function New-EndFileResult {
            param(
                [pscustomobject[]]$Check = @(),
                [pscustomobject[]]$Matrices = @(),
                [string]$Name = 'TestFile'
            )
            [PSCustomObject]@{
                Check          = [System.Collections.Generic.List[pscustomobject]]@($Check)
                Matrices       = $Matrices
                Item           = [PSCustomObject]@{ BaseName = $Name; Name = "$Name.xlsx" }
                LogFolder      = $null
                ReportFileName = "$Name.html"
                ReportFilePath = $null
            }
        }

        function New-FatalCheck {
            param([string]$Name = 'TestFatal', $Value = $null)
            [PSCustomObject]@{
                Type        = 'FatalError'
                Name        = $Name
                Description = 'Test fatal'
                Value       = $Value
            }
        }
    }

    BeforeEach {
        $script:systemErrors = [System.Collections.Generic.List[pscustomobject]]::new()
        Remove-Item 'TestDrive:\*' -Recurse -Force -ErrorAction Ignore

        # ALL helpers mocked. EndHC tests verify orchestration only.
        Mock Update-MatrixCounterHC { return @{ Total = @{ Errors = 0; Warnings = 0 } } }
        Mock Initialize-HtmlStructureHC { return @{ Style = '<style></style>' } }
        Mock Build-MatrixEmailHtmlHC { return '<table>matrix</table>' }
        Mock Build-ErrorWarningTableHC { return '<table>errors</table>' }
        Mock Generate-MailBodyHtmlHC { return '<html><body>OK</body></html>' }
        Mock Export-FilesHC { return @{ HtmlOverview = 'TestDrive:\overview.html' } }
        Mock Generate-MailRecipientListHC { return @('test@example.com') }
        Mock Generate-MailSubjectHC { return 'Test Subject' }
        Mock Send-MailKitMessageHC { }
        Mock Save-MailBodyToLogHC { return 'TestDrive:\Logs\mail.html' }
        Mock Write-EventLogSafe { }
        Mock Cleanup-OldLogsHC { }
        Mock Write-MatrixExecutionReportHC { }
    }

    Context 'Phase 1: Build HTML body' {
        It 'calls Build-MatrixEmailHtmlHC when FileResults has entries' {
            $ctx = New-EndContext -FileResults @((New-EndFileResult))

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            Should -Invoke Build-MatrixEmailHtmlHC -Times 1
        }

        It 'skips Build-MatrixEmailHtmlHC when FileResults is empty' {
            $ctx = New-EndContext -FileResults @()

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            Should -Invoke Build-MatrixEmailHtmlHC -Times 0
        }

        It 'always calls Generate-MailBodyHtmlHC' {
            $ctx = New-EndContext

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            Should -Invoke Generate-MailBodyHtmlHC -Times 1
        }

        It 'records a Warning when HTML generation throws (does not abort pipeline)' {
            Mock Generate-MailBodyHtmlHC { throw 'html boom' }
            $ctx = New-EndContext

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            $htmlWarnings = $systemErrors.Where({
                    $_.Name -eq 'HTML Generation' -and $_.Type -eq 'Warning'
                })
            $htmlWarnings.Count | Should -Be 1

            # Pipeline continues: subsequent phases should still attempt to run
            Should -Invoke Send-MailKitMessageHC -Times 1
        }

        It 'sends mail when SendMail config is present' {
            $ctx = New-EndContext

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)
     
            Should -Invoke Send-MailKitMessageHC -Times 1
        }
    }

    Context 'Phase 2: Exports & ServiceNow' {
        It 'skips Export-FilesHC when fatal errors are present in SystemErrors' {
            $systemErrors.Add((New-FatalCheck))
            $ctx = New-EndContext -AllMatrices @((New-EndMatrix))

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            Should -Invoke Export-FilesHC -Times 0
        }

        It 'skips Export-FilesHC when AllMatrices is empty' {
            $ctx = New-EndContext -AllMatrices @()

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            Should -Invoke Export-FilesHC -Times 0
        }

        It 'calls Export-FilesHC on the happy path' {
            $ctx = New-EndContext -AllMatrices @((New-EndMatrix))

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            Should -Invoke Export-FilesHC -Times 1
        }

        It 'invokes the ServiceNow script only when both Excel path AND credentials are set' {
            $null = New-Item 'TestDrive:\Snow.ps1' -ItemType File -Force

            $ctx = New-EndContext `
                -AllMatrices @((New-EndMatrix)) `
                -Export @{ ServiceNowFormDataExcelFile = 'TestDrive:\snow.xlsx' } `
                -ServiceNow @{
                CredentialsFilePath = 'TestDrive:\creds.json'
                Environment         = 'Prod'
                TableName           = 'u_test'
            }

            # Mock the script invocation by stubbing it as a function the same name
            $script:snowCalled = $false
            $ctx.ScriptPath.UpdateServiceNow = 'TestDrive:\Snow.ps1'

            # We can't easily mock `& $path` direct script invocation in Pester,
            # so this test verifies through side effect: a real testscript on
            # TestDrive that touches a file when called.
            $marker = 'TestDrive:\snow-was-called.txt'
            Set-Content -Path $ctx.ScriptPath.UpdateServiceNow -Value @"
param(`$CredentialsFilePath, `$Environment, `$TableName, `$FormDataExcelFilePath, `$ExcelFileWorksheetName)
'called' | Set-Content -Path '$marker'
"@

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            Test-Path $marker | Should -Be $true
        }

        It 'skips the ServiceNow script when CredentialsFilePath is missing' {
            $ctx = New-EndContext `
                -AllMatrices @((New-EndMatrix)) `
                -Export @{ ServiceNowFormDataExcelFile = 'TestDrive:\snow.xlsx' } `
                -ServiceNow @{ CredentialsFilePath = $null }

            $marker = 'TestDrive:\snow-was-called.txt'
            $scriptPath = New-Item 'TestDrive:\Snow.ps1' -ItemType File -Force
            Set-Content -Path $scriptPath.FullName -Value "'called' | Set-Content '$marker'"
            $ctx.ScriptPath.UpdateServiceNow = $scriptPath.FullName

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            Test-Path $marker | Should -Be $false
        }

        It 'records a Warning when Export-FilesHC throws' {
            Mock Export-FilesHC { throw 'export boom' }
            $ctx = New-EndContext -AllMatrices @((New-EndMatrix))

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            $systemErrors.Where({ $_.Name -eq 'Exports/ServiceNow' }).Count | Should -Be 1
        }
    }

    Context 'Phase 3: Log files' {
        It 'creates a dated log folder when FoundMatrices is true' {
            $logRoot = (New-Item 'TestDrive:\Logs' -ItemType Directory -Force).FullName
            $ctx = New-EndContext -LogFolder $logRoot -FileResults @((New-EndFileResult))

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            (Get-ChildItem -Path $logRoot -Directory).Count | Should -BeGreaterThan 0
        }

        It 'skips dated log folder creation when FoundMatrices is false and no email is sent' {
            # The dated folder is created lazily — only when something writes
            # to it. With FoundMatrices=$false, no per-file logs run. With
            # SendMail=$null and no errors, the email block is gated off too.
            # Net result: nothing writes, nothing gets created.
            $logRoot = (New-Item 'TestDrive:\Logs' -ItemType Directory -Force).FullName
            $ctx = New-EndContext -LogFolder $logRoot -FoundMatrices $false -SendMail $null

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            (Get-ChildItem -Path $logRoot -Directory -ErrorAction Ignore).Count | Should -Be 0
        }

        It 'creates a dated log folder when FoundMatrices is false but errors occurred (email triggers it)' {
            # Regression guard: even without matrices, an error-only run still
            # sends an email (per the gating rule), and the email body save
            # triggers the lazy dated-folder creation.
            $logRoot = (New-Item 'TestDrive:\Logs' -ItemType Directory -Force).FullName
            $ctx = New-EndContext -LogFolder $logRoot -FoundMatrices $false
            $systemErrors.Add([pscustomobject]@{
                    Type    = 'FatalError'
                    Name    = 'Upstream Failure'
                    Message = 'something failed before we got here'
                })

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            (Get-ChildItem -Path $logRoot -Directory -ErrorAction Ignore).Count | Should -Be 1
        }

        It 'falls back to TEMP\PermissionMatrixLogs when configured folder cannot be created' {
            # Use a deliberately invalid path - colon in middle is invalid on Windows
            $ctx = New-EndContext -LogFolder 'C:\<invalid>\path' -FoundMatrices $true

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            $fallbackWarning = $systemErrors.Where({ $_.Name -eq 'Log Folder Fallback' })
            $fallbackWarning.Count | Should -Be 1
        }
    }

    It 'creates JSON files only for checks with a Value property' {
        $logRoot = (New-Item 'TestDrive:\Logs' -ItemType Directory -Force).FullName

        $checkWithValue = New-FatalCheck -Name 'WithValue' -Value 'some data'
        $checkWithoutValue = New-FatalCheck -Name 'NoValue' -Value $null

        $fileResult = New-EndFileResult -Check @($checkWithValue, $checkWithoutValue) -Name 'TestFile'
        $ctx = New-EndContext -LogFolder $logRoot -FileResults @($fileResult)

        Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

        $jsonFiles = Get-ChildItem -Path $logRoot -Recurse -Filter '*.json' -ErrorAction Ignore
        # Only the check WithValue should produce a JSON file
        $jsonFiles.Count | Should -Be 1
    }

    Context 'Phase 4: Send email' {
        It 'sends mail when SendMail is configured and matrices were found' {
            $ctx = New-EndContext
    
            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)
    
            Should -Invoke Send-MailKitMessageHC -Times 1
        }
    
        It 'skips mail when SendMail config is missing' {
            $ctx = New-EndContext -SendMail $null
    
            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)
    
            Should -Invoke Send-MailKitMessageHC -Times 0
        }

        It 'skips mail when SendMail is configured but FoundMatrices is false and no errors occurred' {
            # The "silent run" case — script runs every 5 minutes, nothing to do,
            # don't spam recipients with empty reports.
            $ctx = New-EndContext -FoundMatrices $false

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            Should -Invoke Send-MailKitMessageHC -Times 0
        }
    
        It 'saves the mail body to log folder when log folder exists' {
            $logRoot = (New-Item 'TestDrive:\Logs' -ItemType Directory -Force).FullName
            $ctx = New-EndContext -LogFolder $logRoot
    
            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)
    
            Should -Invoke Save-MailBodyToLogHC -Times 1
        }
    
        It 'records a Warning when Send-MailKitMessageHC throws' {
            Mock Send-MailKitMessageHC { throw 'mail boom' }
            $ctx = New-EndContext
    
            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)
    
            $systemErrors.Where({ $_.Name -eq 'Email Failed' }).Count | Should -Be 1
        }
        
        It 'does not send mail when FoundMatrices is false and no errors occurred' {
            $ctx = New-EndContext -FoundMatrices $false  # SendMail defaults to populated

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            Should -Invoke Send-MailKitMessageHC -Times 0
        }
    }
    
    Context 'Phase 5: Event log & cleanup' {
        It 'writes to event log only when SaveInEventLog.Save is true' {
            $ctx = New-EndContext -SaveInEventLog @{ Save = $true; LogName = 'Application' }
    
            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)
    
            Should -Invoke Write-EventLogSafe -Times 1
        }
    
        It 'skips event log when SaveInEventLog.Save is false' {
            $ctx = New-EndContext -SaveInEventLog @{ Save = $false; LogName = 'Application' }
    
            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)
    
            Should -Invoke Write-EventLogSafe -Times 0
        }
    
        It 'falls back to default ScriptName when not set' {
            $ctx = New-EndContext `
                -ScriptName $null `
                -SaveInEventLog @{ Save = $true; LogName = 'Application' }
    
            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)
    
            Should -Invoke Write-EventLogSafe -ParameterFilter {
                $ScriptName -eq 'Permission Matrix'
            }
        }
    
        It 'cleans up old logs when DeleteLogsAfterDays > 0 and log folder exists' {
            $logRoot = (New-Item 'TestDrive:\Logs' -ItemType Directory -Force).FullName
            $ctx = New-EndContext `
                -LogFolder $logRoot `
                -SaveLogFiles @{
                Detailed            = $false
                DeleteLogsAfterDays = 30
                Where               = @{ Folder = $logRoot }
            }
    
            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)
    
            Should -Invoke Cleanup-OldLogsHC -Times 1
        }
    
        It 'skips cleanup when DeleteLogsAfterDays is 0' {
            $ctx = New-EndContext
    
            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)
    
            Should -Invoke Cleanup-OldLogsHC -Times 0
        }
    
        It 'does not throw when phase 5 itself fails (final catch is silent)' {
            Mock Write-EventLogSafe { throw 'eventlog boom' }
            $ctx = New-EndContext -SaveInEventLog @{ Save = $true; LogName = 'Application' }
    
            { Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors) } |
            Should -Not -Throw
        }
    }
    
    Context 'Integration: control flow' {
        It 'runs all phases on a happy path' {
            $ctx = New-EndContext `
                -AllMatrices @((New-EndMatrix)) `
                -FileResults @((New-EndFileResult)) `
                -SaveInEventLog @{ Save = $true; LogName = 'Application' }
    
            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)
    
            Should -Invoke Generate-MailBodyHtmlHC -Times 1
            Should -Invoke Export-FilesHC -Times 1
            Should -Invoke Send-MailKitMessageHC -Times 1
            Should -Invoke Write-EventLogSafe -Times 1
        }
    
        It 'continues through later phases when earlier phases fail' {
            Mock Build-MatrixEmailHtmlHC { throw 'phase 1 boom' }
            $ctx = New-EndContext -FileResults @((New-EndFileResult))
    
            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)
    
            # Email is still attempted
            Should -Invoke Send-MailKitMessageHC -Times 1
        }
    }
}
