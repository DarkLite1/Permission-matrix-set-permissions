#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

Describe 'Invoke-PermissionMatrixEndHC' {
    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        Get-ChildItem "$moduleRoot\Private" -Filter '*.ps1' -File |
        ForEach-Object { . $_.FullName }

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
                [hashtable]$SharePoint = @{},
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
                [hashtable]$ScriptPath = @{
                    UpdateServiceNow   = 'TestDrive:\Snow.ps1'
                    UploadToSharePoint = (Join-Path $TestDrive 'Upload.ps1')
                },
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
                    SharePoint = [PSCustomObject]$SharePoint
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

        function New-UploadStub {
            <#
            .SYNOPSIS
                Write a stand-in UploadToSharePoint.ps1 onto TestDrive.

            .DESCRIPTION
                EndHC invokes the operation script with the call operator
                (& $Context.ScriptPath.UploadToSharePoint @spParams), which
                Pester cannot mock. This writes a real script that records the
                arguments it received as JSON, so tests can assert WHICH
                arguments were passed rather than only that a call happened.

                Every parameter is optional on purpose: a mandatory one the
                caller forgot would make PowerShell prompt, hanging a
                non-interactive run instead of failing it. Missing arguments are
                detected from the captured JSON instead.

            .PARAMETER Path
                Where to write the stub script.

            .PARAMETER CapturePath
                Where the stub writes the JSON record of its arguments.

            .PARAMETER Throw
                Make the stub throw after recording, to exercise the caller's
                error handling.
            #>
            param(
                [Parameter(Mandatory)][string]$Path,
                [Parameter(Mandatory)][string]$CapturePath,
                [switch]$Throw
            )

            $body = @"
param(
    [String]`$FilePath,
    [String]`$SiteUrl,
    [String]`$DocumentLibraryName,
    [String]`$FolderPath,
    [String]`$FileName,
    [String]`$ClientId,
    [String]`$TenantId,
    [String]`$CertificateThumbprint,
    [int]`$MaxRetries,
    [int]`$ChunkSizeMB
)

`$PSBoundParameters | ConvertTo-Json -Depth 3 |
    Set-Content -LiteralPath '$CapturePath' -Encoding UTF8
"@

            if ($Throw) {
                $body += "`n`nthrow 'upload boom'"
            }

            $null = New-Item -Path $Path -ItemType File -Force

            Set-Content -Path $Path -Value $body

            # Set-Content has been observed to report success without producing
            # the file when handed a TestDrive: qualified path. Fail here rather
            # than let every assertion downstream report a confusing absence.
            if (-not (Test-Path -LiteralPath $Path)) {
                throw "New-UploadStub could not create the stub at '$Path'."
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
        Mock Get-MailBodyHtmlHC { return '<html><body>OK</body></html>' }
        # Matches the shape Export-FilesHC really returns, with the keys EndHC
        # guards on present. Those guards read ExportedFiles.FormData and
        # ExportedFiles.OverviewHtml, so a mock missing a key silently skips the
        # branch a test means to exercise.
        Mock Export-FilesHC {
            return [ordered]@{
                Permissions  = $null
                FormData     = $null
                OverviewHtml = 'TestDrive:\overview.html'
            }
        }
        Mock Get-MailRecipientListHC { return @('test@example.com') }
        Mock Get-MailSubjectHC { return 'Test Subject' }
        Mock Send-MailKitMessageHC { }
        Mock Save-MailBodyToLogHC { return 'TestDrive:\Logs\mail.html' }
        Mock Write-EventLogSafeHC { }
        Mock Remove-OldLogsHC { }
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

        It 'always calls Get-MailBodyHtmlHC' {
            $ctx = New-EndContext

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            Should -Invoke Get-MailBodyHtmlHC -Times 1
        }

        It 'passes exported files and the planned mail log path to Get-MailBodyHtmlHC' {
            Mock Export-FilesHC { return [ordered]@{ Permissions = 'TestDrive:\Permissions.xlsx' } }
            $ctx = New-EndContext -AllMatrices @((New-EndMatrix))

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            Should -Invoke Get-MailBodyHtmlHC -Times 1 -ParameterFilter {
                $ExportedFiles.Permissions -eq 'TestDrive:\Permissions.xlsx' -and
                $BrowserViewFilePath -like '*Mail - Test Subject.html'
            }
        }

        It 'records a Warning when HTML generation throws (does not abort pipeline)' {
            Mock Get-MailBodyHtmlHC { throw 'html boom' }
            $ctx = New-EndContext

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            $htmlWarnings = $systemErrors.Where({
                    $_.Name -eq 'HTML Generation' -and $_.Type -eq 'Warning'
                })
            $htmlWarnings.Count | Should -Be 1

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

            # The guard now requires Export-FilesHC to have produced the workbook, not
            # merely for it to be configured.
            Mock Export-FilesHC {
                return @{
                    Permissions  = $null
                    FormData     = 'TestDrive:\snow.xlsx'
                    OverviewHtml = $null
                }
            }

            $ctx = New-EndContext `
                -AllMatrices @((New-EndMatrix)) `
                -Export @{ ServiceNowFormDataExcelFile = 'TestDrive:\snow.xlsx' } `
                -ServiceNow @{
                CredentialsFilePath = 'TestDrive:\creds.json'
                Environment         = 'Prod'
                TableName           = 'u_test'
            }

            $ctx.ScriptPath.UpdateServiceNow = 'TestDrive:\Snow.ps1'

            # We can't mock `& $path` script invocation in Pester, so this verifies
            # through side effect: a real script on TestDrive that touches a file.
            $marker = 'TestDrive:\snow-was-called.txt'
            Set-Content -Path $ctx.ScriptPath.UpdateServiceNow -Value @"
param(`$CredentialsFilePath, `$Environment, `$TableName, `$FormDataExcelFilePath, `$ExcelFileWorksheetName)
'called' | Set-Content -Path '$marker'
"@

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            Test-Path $marker | Should -Be $true
        }

        It 'skips the ServiceNow script when CredentialsFilePath is missing' {
            # The export succeeded and produced the workbook, so the missing
            # credentials are the only reason to skip. Without this the guard
            # would fail on ExportedFiles.FormData first and the test would pass
            # even if the credentials check were removed.
            Mock Export-FilesHC {
                return [ordered]@{
                    Permissions  = $null
                    FormData     = 'TestDrive:\snow.xlsx'
                    OverviewHtml = $null
                }
            }

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

            $systemErrors.Where({ $_.Name -eq 'Exports' }).Count | Should -Be 1
        }
        It 'does not update ServiceNow when Export-FilesHC threw' {
            Mock Export-FilesHC { throw 'export boom' }

            $marker = Join-Path $TestDrive 'snow-was-called.txt'
            $snowStub = Join-Path $TestDrive 'Snow.ps1'
            Set-Content -Path $snowStub -Value "'called' | Set-Content '$marker'"

            $ctx = New-EndContext `
                -AllMatrices @((New-EndMatrix)) `
                -Export @{ ServiceNowFormDataExcelFile = 'TestDrive:\snow.xlsx' } `
                -ServiceNow @{ CredentialsFilePath = 'TestDrive:\creds.json' }
            $ctx.ScriptPath.UpdateServiceNow = $snowStub

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            Test-Path $marker | Should -BeFalse
        }
    }

    Context 'Phase 2b: SharePoint upload' {
        BeforeEach {
            $uploadStub = Join-Path $TestDrive 'Upload.ps1'
            $capture = Join-Path $TestDrive 'upload-args.json'

            New-UploadStub -Path $uploadStub -CapturePath $capture

            # The file-level default already supplies OverviewHtml, but these
            # tests assert on the value, so it is pinned here rather than left
            # coupled to whatever the shared default happens to hold.
            Mock Export-FilesHC {
                return [ordered]@{
                    Permissions  = $null
                    FormData     = $null
                    OverviewHtml = 'TestDrive:\overview.html'
                }
            }

            $sharePointConfig = @{
                SiteUrl               = 'https://contoso.sharepoint.com/sites/IT'
                DocumentLibraryName   = 'Documents'
                ClientId              = 'client-id-1'
                TenantId              = 'tenant-id-1'
                CertificateThumbprint = 'THUMBPRINT1'
            }
        }

        It 'uploads when an overview html was exported and a site is configured' {
            $ctx = New-EndContext `
                -AllMatrices @((New-EndMatrix)) `
                -SharePoint $sharePointConfig

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            Test-Path $capture | Should -BeTrue
        }

        It 'does not upload when no site is configured' {
            $ctx = New-EndContext -AllMatrices @((New-EndMatrix)) -SharePoint @{}

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            Test-Path $capture | Should -BeFalse
        }

        It 'does not upload when no overview html was exported' {
            # Export-FilesHC ran but OverviewHtmlFile was not configured, so
            # there is nothing to upload even though SharePoint is set up.
            Mock Export-FilesHC {
                return [ordered]@{
                    Permissions  = $null
                    FormData     = $null
                    OverviewHtml = $null
                }
            }

            $ctx = New-EndContext `
                -AllMatrices @((New-EndMatrix)) `
                -SharePoint $sharePointConfig

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            Test-Path $capture | Should -BeFalse
        }

        It 'does not upload when there are no matrices to report on' {
            $ctx = New-EndContext -AllMatrices @() -SharePoint $sharePointConfig

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            Test-Path $capture | Should -BeFalse
        }

        It 'passes the exported html file and the configured site settings' {
            $ctx = New-EndContext `
                -AllMatrices @((New-EndMatrix)) `
                -SharePoint $sharePointConfig

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            $sent = Get-Content $capture -Raw | ConvertFrom-Json

            $sent.FilePath | Should -BeExactly 'TestDrive:\overview.html'
            $sent.SiteUrl | Should -BeExactly 'https://contoso.sharepoint.com/sites/IT'
            $sent.DocumentLibraryName | Should -BeExactly 'Documents'
            $sent.ClientId | Should -BeExactly 'client-id-1'
            $sent.TenantId | Should -BeExactly 'tenant-id-1'
            $sent.CertificateThumbprint | Should -BeExactly 'THUMBPRINT1'
        }

        It 'forwards the file path produced by Export-FilesHC, not the configured one' {
            # The config holds the path the html SHOULD be written to;
            # Export-FilesHC returns where it actually landed. The upload must
            # follow the latter.
            Mock Export-FilesHC {
                return [ordered]@{
                    Permissions  = $null
                    FormData     = $null
                    OverviewHtml = 'TestDrive:\actual-overview.html'
                }
            }

            $ctx = New-EndContext `
                -AllMatrices @((New-EndMatrix)) `
                -Export @{ OverviewHtmlFile = 'TestDrive:\configured-overview.html' } `
                -SharePoint $sharePointConfig

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            $sent = Get-Content $capture -Raw | ConvertFrom-Json

            $sent.FilePath | Should -BeExactly 'TestDrive:\actual-overview.html'
        }

        It 'omits FolderPath and FileName when they are not configured' {
            # These are optional on the script and must not be sent as empty
            # strings, which would upload to a folder literally named ''.
            $ctx = New-EndContext `
                -AllMatrices @((New-EndMatrix)) `
                -SharePoint $sharePointConfig

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            # Without this the assertions below pass on a $null when the upload
            # never ran, which is a false green.
            Test-Path $capture | Should -BeTrue

            $sentNames = (Get-Content $capture -Raw |
                ConvertFrom-Json).PSObject.Properties.Name

            $sentNames | Should -Not -Contain 'FolderPath'
            $sentNames | Should -Not -Contain 'FileName'
        }

        It 'passes FolderPath and FileName when they are configured' {
            $ctx = New-EndContext `
                -AllMatrices @((New-EndMatrix)) `
                -SharePoint ($sharePointConfig + @{
                    FolderPath = 'Reports/Permission matrix'
                    FileName   = 'Permission matrix overview.html'
                })

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            $sent = Get-Content $capture -Raw | ConvertFrom-Json

            $sent.FolderPath | Should -BeExactly 'Reports/Permission matrix'
            $sent.FileName | Should -BeExactly 'Permission matrix overview.html'
        }

        It 'sends only parameters that UploadToSharePoint.ps1 declares' {
            # Guards against caller/script drift: renaming a parameter on the
            # script without updating EndHC would fail here rather than at 02:00
            # in production, where the splat would throw a binding error.
            $ctx = New-EndContext `
                -AllMatrices @((New-EndMatrix)) `
                -SharePoint ($sharePointConfig + @{
                    FolderPath = 'Reports'
                    FileName   = 'Overview.html'
                })

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            # Without this the assertion below passes on a $null when the upload
            # never ran, which is a false green.
            Test-Path $capture |
            Should -BeTrue -Because 'the upload must have run for this assertion to mean anything'

            $sentNames = (Get-Content $capture -Raw |
                ConvertFrom-Json).PSObject.Properties.Name

            $realScript = Join-Path $root 'Scripts\Operations\UploadToSharePoint.ps1'
            $declared = (Get-Command $realScript).Parameters.Keys

            $unknown = $sentNames | Where-Object { $_ -notin $declared }

            $unknown | Should -BeNullOrEmpty -Because "EndHC must not splat parameters the script does not accept, found: $($unknown -join ', ')"
        }

        It 'sends every mandatory parameter of UploadToSharePoint.ps1' {
            $ctx = New-EndContext `
                -AllMatrices @((New-EndMatrix)) `
                -SharePoint $sharePointConfig

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            $sentNames = (Get-Content $capture -Raw |
                ConvertFrom-Json).PSObject.Properties.Name

            $realScript = Join-Path $root 'Scripts\Operations\UploadToSharePoint.ps1'

            $mandatory = (Get-Command $realScript).Parameters.Values |
            Where-Object {
                $_.Attributes.Where({
                        $_ -is [System.Management.Automation.ParameterAttribute] -and $_.Mandatory
                    })
            } |
            Select-Object -ExpandProperty Name

            $missing = $mandatory | Where-Object { $_ -notin $sentNames }

            $missing | Should -BeNullOrEmpty -Because "EndHC must supply every mandatory parameter, missing: $($missing -join ', ')"
        }

        It 'records a Warning and still sends the mail when the upload throws' {
            New-UploadStub -Path $uploadStub -CapturePath $capture -Throw

            $ctx = New-EndContext `
                -AllMatrices @((New-EndMatrix)) `
                -SharePoint $sharePointConfig

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            # The stub records its arguments before throwing, so this proves the
            # warning came from the upload and not from something earlier in the
            # export phase that the same catch block also covers.
            Test-Path $capture | Should -BeTrue

            # A SharePoint outage must not cost us the notification mail.
            $systemErrors.Where({
                    $_.Name -eq 'SharePoint' -and $_.Type -eq 'Warning'
                }).Count | Should -Be 1

            Should -Invoke Send-MailKitMessageHC -Times 1
        }

        It 'does not upload when Export-FilesHC threw' {
            Mock Export-FilesHC { throw 'export boom' }

            $ctx = New-EndContext `
                -AllMatrices @((New-EndMatrix)) `
                -SharePoint $sharePointConfig

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            Test-Path $capture | Should -BeFalse
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
    
        It 'sends mail when FoundMatrices is false but a system error occurred' {
            # Even with no matrices, an upstream failure should be reported.
            $ctx = New-EndContext -FoundMatrices $false
            $systemErrors.Add([pscustomobject]@{
                    Type    = 'FatalError'
                    Name    = 'Upstream Failure'
                    Message = 'something failed before we got here'
                })

            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)

            Should -Invoke Send-MailKitMessageHC -Times 1
        }

        It 'does not send mail when FoundMatrices is false and no errors occurred' {
            $ctx = New-EndContext -FoundMatrices $false  # SendMail defaults to populated

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
    }
    
    Context 'Phase 5: Event log & cleanup' {
        It 'writes to event log only when SaveInEventLog.Save is true' {
            $ctx = New-EndContext -SaveInEventLog @{ Save = $true; LogName = 'Application' }
    
            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)
    
            Should -Invoke Write-EventLogSafeHC -Times 1
        }
    
        It 'skips event log when SaveInEventLog.Save is false' {
            $ctx = New-EndContext -SaveInEventLog @{ Save = $false; LogName = 'Application' }
    
            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)
    
            Should -Invoke Write-EventLogSafeHC -Times 0
        }
    
        It 'falls back to default ScriptName when not set' {
            $ctx = New-EndContext `
                -ScriptName $null `
                -SaveInEventLog @{ Save = $true; LogName = 'Application' }
    
            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)
    
            Should -Invoke Write-EventLogSafeHC -ParameterFilter {
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
    
            Should -Invoke Remove-OldLogsHC -Times 1
        }
    
        It 'skips cleanup when DeleteLogsAfterDays is 0' {
            $ctx = New-EndContext
    
            Invoke-PermissionMatrixEndHC -Context $ctx -SystemErrors ([ref]$systemErrors)
    
            Should -Invoke Remove-OldLogsHC -Times 0
        }
    
        It 'does not throw when phase 5 itself fails (final catch is silent)' {
            Mock Write-EventLogSafeHC { throw 'eventlog boom' }
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
    
            Should -Invoke Get-MailBodyHtmlHC -Times 1
            Should -Invoke Export-FilesHC -Times 1
            Should -Invoke Send-MailKitMessageHC -Times 1
            Should -Invoke Write-EventLogSafeHC -Times 1
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