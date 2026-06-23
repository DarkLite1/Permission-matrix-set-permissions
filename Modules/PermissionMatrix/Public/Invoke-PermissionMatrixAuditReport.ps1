function Invoke-PermissionMatrixAuditReport {
    <#
    .SYNOPSIS
        Public orchestrator for the monthly permission audit report.

    .DESCRIPTION
        Counterpart of Invoke-PermissionMatrix for the audit report. It REUSES
        the Begin stage (Invoke-PermissionMatrixBeginHC) to read the matrices,
        resolve Active Directory and rewrite names to SIDs, then - WITHOUT
        running the Process or End stages, so no ACL is ever changed - it:

          - disregards matrices that came out of Begin with fatal errors and
            reports those to the admin in a single summary mail,
          - writes the per-matrix Excel log file (Copy-MatrixFileToLogFolderHC,
            the same file the matrix run produces), and
          - e-mails each MatrixResponsible that log file, using the subject and
            body templates from the audit configuration.

        Mail is sent with the same private transport the End stage uses
        (Send-MailKitMessageHC), driven by 'Settings.SendMail' from the audit
        config. Because this is a public module function it runs in module
        scope and can call the private *HC helpers directly.

    .PARAMETER ConfigurationJsonFile
        Path to the audit JSON configuration file. The file only needs the
        fields the audit uses (Matrix.FolderPath/DefaultsFile, the AuditReport
        section, and Settings.SendMail/SaveLogFiles/SaveInEventLog/ScriptName).
        The schema-only fields the shared validator expects (Export, ServiceNow,
        PSSessionConfiguration, MaxConcurrent, Matrix.Archive,
        Settings.SaveLogFiles.Detailed) are filled with defaults in memory
        before validation, so they may be omitted from the file.

    .PARAMETER ScriptPath
        Same hashtable contract as Invoke-PermissionMatrix. For the audit only
        'PermissionMatrixModule' is required: the Begin stage uses it to
        dot-source the private functions inside its runspaces. The
        SetPermissions / UpdateServiceNow / TestRequirements paths are not
        needed because the Process and End stages never run.

    .PARAMETER SystemErrors
        [ref] collection the Begin stage appends initialization errors to.

    .PARAMETER ScriptStartTime
        The script start time, used as the single timestamp in the per-matrix
        log file names ('yyyy-MM-dd HHmm (dddd) - ScriptName - Matrix.xlsx').
        Defaults to the current time; the entry-point passes the time captured
        at the start of the run so every log file from one run shares it.

    .NOTES
        The admin notification recipients (BCC'd on every audit mail and
        notified about skipped matrices and initialization failures) are read
        from 'AuditReport.ScriptAdmin' in the configuration file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ConfigurationJsonFile,
        [Parameter(Mandatory)][hashtable]$ScriptPath,
        [Parameter(Mandatory)][ref]$SystemErrors,
        [datetime]$ScriptStartTime = (Get-Date)
    )

    #region Fill the schema-only fields the shared validator still expects
    # The audit input file only carries what the audit actually uses. The Begin
    # stage runs Test-ConfigurationStructureHC, which also expects a handful of
    # fields the audit never touches: the 'Export', 'ServiceNow',
    # 'PSSessionConfiguration' and 'MaxConcurrent' top-level blocks (Process/End
    # concerns), 'Matrix.Archive' and 'Settings.SaveLogFiles.Detailed'. Rather
    # than repeat that boilerplate in every audit config, fill sensible defaults
    # here, in memory, before validation. Anything already present in the file
    # is left untouched. If the file cannot be read/parsed we fall back to the
    # original path so Begin reports the problem exactly as it normally would.
    $auditConfig = $null
    try {
        $auditConfig = Get-Content -LiteralPath $ConfigurationJsonFile -Raw -ErrorAction Stop |
            ConvertFrom-Json -ErrorAction Stop
    }
    catch { $auditConfig = $null }

    $configFileForBegin = $ConfigurationJsonFile
    $mergedConfigFile = $null

    if ($auditConfig) {
        $topLevelDefaults = [ordered]@{
            Export                 = [ordered]@{
                PermissionsExcelFile        = $null
                OverviewHtmlFile            = $null
                ServiceNowFormDataExcelFile = $null
            }
            ServiceNow             = [ordered]@{
                CredentialsFilePath = $null
                Environment         = $null
                TableName           = $null
            }
            PSSessionConfiguration = 'PowerShell.7'
            MaxConcurrent          = [ordered]@{
                Computers             = 1
                FoldersPerMatrix      = 3
                JobsPerRemoteComputer = 1
            }
        }
        foreach ($key in $topLevelDefaults.Keys) {
            if ($null -eq $auditConfig.$key) {
                $auditConfig | Add-Member -NotePropertyName $key `
                    -NotePropertyValue $topLevelDefaults[$key] -Force
            }
        }
        # The audit only reads matrices: it must never archive them.
        if ($auditConfig.Matrix -and $null -eq $auditConfig.Matrix.Archive) {
            $auditConfig.Matrix | Add-Member -NotePropertyName 'Archive' `
                -NotePropertyValue $false -Force
        }
        # 'Detailed' log sheets are a Process-stage concern, unused by the audit.
        if ($auditConfig.Settings -and $auditConfig.Settings.SaveLogFiles -and
            $null -eq $auditConfig.Settings.SaveLogFiles.Detailed) {
            $auditConfig.Settings.SaveLogFiles | Add-Member -NotePropertyName 'Detailed' `
                -NotePropertyValue $false -Force
        }

        $mergedConfigFile = Join-Path ([System.IO.Path]::GetTempPath()) (
            'PermissionMatrixAudit_{0}.json' -f [guid]::NewGuid()
        )
        $auditConfig | ConvertTo-Json -Depth 25 |
            Set-Content -LiteralPath $mergedConfigFile -Encoding utf8
        $configFileForBegin = $mergedConfigFile
    }
    #endregion

    try {
        $context = Invoke-PermissionMatrixBeginHC `
            -ConfigurationJsonFile $configFileForBegin `
            -ScriptPath $ScriptPath `
            -SystemErrors $SystemErrors
    }
    finally {
        if ($mergedConfigFile) {
            Remove-Item -LiteralPath $mergedConfigFile -ErrorAction SilentlyContinue
        }
    }

    $config = if ($context) { $context.Config } else { $null }
    $sendMail = $config.Settings.SendMail

    # Admin notification recipients come from the audit config
    $scriptAdmin = @($config.AuditReport.ScriptAdmin | Where-Object { $_ })

    # Local helper: assemble Send-MailKitMessageHC parameters from the audit
    # config transport plus the supplied message, mirroring the End stage.
    $newMailKitParams = {
        param($To, $Subject, $Body, $Attachments, $Bcc, $Priority)

        $p = @{
            To                  = $To
            From                = Get-StringValueHC $sendMail.From
            FromDisplayName     = Get-StringValueHC $sendMail.FromDisplayName
            SmtpServerName      = Get-StringValueHC $sendMail.Smtp.ServerName
            SmtpPort            = Get-StringValueHC $sendMail.Smtp.Port
            SmtpConnectionType  = Get-StringValueHC $sendMail.Smtp.ConnectionType
            MailKitAssemblyPath = Get-StringValueHC $sendMail.AssemblyPath.MailKit
            MimeKitAssemblyPath = Get-StringValueHC $sendMail.AssemblyPath.MimeKit
            Subject             = $Subject
            Body                = $Body
            Priority            = $Priority
        }

        if ($Attachments) { $p.Attachments = $Attachments }
        if ($Bcc) { $p.Bcc = $Bcc }

        $smtpUser = Get-StringValueHC $sendMail.Smtp.UserName
        $smtpPass = Get-StringValueHC $sendMail.Smtp.Password
        if ($smtpUser -and $smtpPass) {
            $secure = ConvertTo-SecureString -String $smtpPass -AsPlainText -Force
            $p.Credential = New-Object System.Management.Automation.PSCredential($smtpUser, $secure)
        }

        Remove-BlankValueHC -Hashtable $p
    }

    #region Stop early on initialization failure (report it to the admin)
    $fatalInit = @($SystemErrors.Value | Where-Object { $_.Type -eq 'FatalError' })
    if ($fatalInit.Count -gt 0) {
        if ($scriptAdmin) {
            try {
                $errHtml = @(
                    $fatalInit | ForEach-Object {
                        '<li>{0}: {1}</li>' -f $_.Name, $_.Message
                    }
                ) -join ''

                $p = & $newMailKitParams `
                    -To $scriptAdmin `
                    -Subject 'Permission matrix audit report - initialization failed' `
                    -Body "<p>The audit could not run because of the following errors:</p><ul>$errHtml</ul>" `
                    -Attachments $null -Bcc $null -Priority 'High'

                Send-MailKitMessageHC @p
            }
            catch {
                Write-Warning "Failed to send initialization-failure mail: $_"
            }
        }
        return
    }
    #endregion

    if (-not $context.FoundMatrices -or -not $context.FileResults) {
        Write-Verbose 'No matrix files found, no audit e-mails to send'
        return
    }

    $requestTicketURL = $context.Config.AuditReport.RequestTicketURL
    $auditLogFolder = $context.Config.Settings.SaveLogFiles.Where.Folder

    # One timestamp for the whole run (the script start time), e.g.
    # '2026_04_13_110000'.
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $runStamp = '{0}' -f $ScriptStartTime.ToString('yyyy_MM_dd_HHmmss', $inv)
    $scriptName = $config.Settings.ScriptName

    # Placeholder accounts (Matrix.AdGroupPlaceHolders) never receive audit mail.
    $placeHolders = @($config.Matrix.AdGroupPlaceHolders | Where-Object { $_ })

    # Test override: when 'Settings.SendMail.To' is set, every audit mail is
    # redirected to these address(es) and the per-matrix responsible resolved
    # from the matrix 'FormData' worksheet is ignored entirely. No Bcc is added
    # in this mode (To only). Leave 'To' empty for normal, per-matrix routing.
    $overrideTo = @($sendMail.To | Where-Object { $_ })
    if ($overrideTo) {
        Write-Verbose "Settings.SendMail.To is set: all audit mail redirected to $($overrideTo -join ', ') (responsible ignored, no Bcc)"
    }

    # Matrices skipped because of fatal errors; reported to the admin at the end
    $skipped = [System.Collections.Generic.List[pscustomobject]]::new()
    # Responsibles / members that could not be resolved to an e-mail address
    $recipientIssues = [System.Collections.Generic.List[pscustomobject]]::new()
    # Matrices whose responsible resolved to nobody (no audit mail sent)
    $noRecipient = [System.Collections.Generic.List[pscustomobject]]::new()

    foreach ($fileResult in $context.FileResults) {

        $matrixName = if ($fileResult.Item) { $fileResult.Item.Name }
        elseif ($fileResult.File) { $fileResult.File.Name }
        else { '(unknown)' }

        $formData = $fileResult.Sheets.FormData.Formatted

        #region Disregard matrices with fatal errors (reported to admin below)
        $fatalChecks = @($fileResult.Check | Where-Object { $_.Type -eq 'FatalError' })
        foreach ($m in $fileResult.Matrices) {
            $fatalChecks += @($m.Check | Where-Object { $_.Type -eq 'FatalError' })
        }

        if ($fatalChecks) {
            $skipped.Add(
                [pscustomobject]@{
                    MatrixFile  = $matrixName
                    Responsible = if ($formData) { @($formData.MatrixResponsible) -join ', ' } else { '' }
                    Errors      = (
                        @($fatalChecks | ForEach-Object { $_.Name }) |
                        Sort-Object -Unique
                    ) -join '; '
                }
            )
            Write-Verbose "Matrix '$matrixName' has fatal errors, skipped (reported to admin)"
            continue
        }
        #endregion

        if ($overrideTo) {
            #region Test override: send to Settings.SendMail.To, ignore responsible
            # Every matrix that produced a FormData row is mailed to the override
            # address(es). The responsible resolved from Excel is not used, the
            # 'no responsible' / 'no e-mail' skips below do not apply, and no Bcc
            # is added (handled after the mail is built).
            if (-not $formData) {
                Write-Verbose "Matrix '$matrixName' has no FormData, skipped (override)"
                continue
            }
            $recipientEmail = $overrideTo
            $mailBcc        = @()
            #endregion
        }
        else {
            #region Skip matrices without a responsible (no one to send to)
            if (-not $formData -or -not $formData.MatrixResponsible) {
                Write-Verbose "Matrix '$matrixName' has no MatrixResponsible, skipped"
                continue
            }
            #endregion

            #region Resolve the responsible (e-mail / user / group) to addresses
            $resolved = Resolve-ResponsibleEmailHC `
                -Responsible $formData.MatrixResponsible `
                -ExcludeSamAccountName $placeHolders

            if ($resolved.Unresolved) {
                $recipientIssues.Add(
                    [pscustomobject]@{
                        MatrixFile = $matrixName
                        Entries    = @($resolved.Unresolved) -join '; '
                    }
                )
                Write-Verbose "Matrix '$matrixName': unresolved recipient(s): $(@($resolved.Unresolved) -join '; ')"
            }

            if (-not $resolved.Emails) {
                $noRecipient.Add(
                    [pscustomobject]@{
                        MatrixFile  = $matrixName
                        Responsible = @($formData.MatrixResponsible) -join ', '
                    }
                )
                Write-Verbose "Matrix '$matrixName': responsible resolved to no e-mail address, no mail sent"
                continue
            }
            #endregion

            $recipientEmail = $resolved.Emails
            $mailBcc        = $scriptAdmin
        }

        #region Build the log sheet rows and write the per-matrix Excel log file
        $logSheets = Build-MatrixLogSheetRowsHC `
            -FileResult $fileResult `
            -AdObjectDetails $context.AdObjectDetails

        $matrixLogFolder = if (
            $fileResult.PSObject.Properties['LogFolder'] -and $fileResult.LogFolder
        ) {
            $fileResult.LogFolder
        }
        else {
            Join-Path $auditLogFolder $fileResult.Item.BaseName
        }

        $null = New-Item -Path $matrixLogFolder -ItemType Directory -Force -ErrorAction Stop

        # Date-stamped base name, e.g.
        # '2026-04-13 1100 (Monday) - Permission matrix audit report (BNL) - BEL-MTX-...'
        # so every run of the same matrix file accumulates in its folder.
        $logBaseName = '{0} - {1} - {2}' -f $runStamp, $scriptName, $fileResult.Item.BaseName

        $attachmentPath = Copy-MatrixFileToLogFolderHC `
            -SourceFilePath $fileResult.Item.FullName `
            -LogFolder $matrixLogFolder `
            -DestinationFileName "$logBaseName.xlsx" `
            -AccessListRows $logSheets.AccessList `
            -GroupManagerRows $logSheets.GroupManagers `
            -AdObjectRows $logSheets.AdObjects
        #endregion

        #region Build and send the mail (subject + body from config, log attached)
        $msg = Build-AuditReportMailHC `
            -FormData $formData `
            -AccessList $logSheets.AccessList `
            -AttachmentPath $attachmentPath `
            -MailSettings $sendMail `
            -RequestTicketURL $requestTicketURL `
            -RecipientEmail $recipientEmail `
            -Bcc $mailBcc

        # 'To only' override: Build-AuditReportMailHC always folds in
        # 'Settings.SendMail.Bcc'; drop it so a test run reaches no one but the
        # override address(es).
        if ($overrideTo) { $msg.Bcc = @() }

        $mailKit = & $newMailKitParams `
            -To $msg.To -Subject $msg.Subject -Body $msg.Body `
            -Attachments $msg.Attachments -Bcc $msg.Bcc -Priority 'Normal'

        # Save the rendered mail body next to the Excel log, e.g.
        # '... - BEL-MTX-... - Mail.html'
        $mailHtmlPath = Join-Path $matrixLogFolder "$logBaseName - Mail.html"
        try {
            $msg.Body | Set-Content -LiteralPath $mailHtmlPath -Encoding utf8 -ErrorAction Stop
        }
        catch { Write-Warning "Failed saving mail body for '$matrixName': $_" }

        Write-Verbose "Matrix '$($formData.MatrixFileName)': mailing $($msg.To -join ', ')"

        try { Send-MailKitMessageHC @mailKit }
        catch { Write-Warning "Failed sending audit mail for '$matrixName': $_" }
        #endregion
    }

    #region Report matrices skipped for fatal errors to the admin
    if ($skipped.Count -gt 0 -and $scriptAdmin) {
        $rows = foreach ($s in $skipped) {
            '<tr><td>{0}</td><td>{1}</td><td>{2}</td></tr>' -f
            $s.MatrixFile, $s.Responsible, $s.Errors
        }

        $adminBody = @"
<p>The following $($skipped.Count) matrix file(s) were skipped during the audit
because they contain fatal errors. No audit e-mail was sent to their
responsible(s).</p>
<table border="1" cellpadding="4" cellspacing="0">
<tr><th>Matrix file</th><th>Responsible</th><th>Fatal errors</th></tr>
$($rows -join "`n")
</table>
"@

        $adminKit = & $newMailKitParams `
            -To $scriptAdmin `
            -Subject "Permission matrix audit report - $($skipped.Count) matrix file(s) skipped (fatal errors)" `
            -Body $adminBody -Attachments $null -Bcc $null -Priority 'High'

        try { Send-MailKitMessageHC @adminKit }
        catch { Write-Warning "Failed sending admin skip-report: $_" }
    }
    #endregion

    #region Report recipients that could not be resolved to an e-mail address
    if (($recipientIssues.Count -gt 0 -or $noRecipient.Count -gt 0) -and $scriptAdmin) {
        $noRecipientRows = foreach ($n in $noRecipient) {
            '<tr><td>{0}</td><td>{1}</td></tr>' -f $n.MatrixFile, $n.Responsible
        }
        $issueRows = foreach ($i in $recipientIssues) {
            '<tr><td>{0}</td><td>{1}</td></tr>' -f $i.MatrixFile, $i.Entries
        }

        $adminBody = ''
        if ($noRecipient.Count -gt 0) {
            $adminBody += @"
<p>The following $($noRecipient.Count) matrix file(s) had a responsible that
could not be resolved to any e-mail address. No audit e-mail was sent.</p>
<table border="1" cellpadding="4" cellspacing="0">
<tr><th>Matrix file</th><th>Responsible</th></tr>
$($noRecipientRows -join "`n")
</table>
"@
        }
        if ($recipientIssues.Count -gt 0) {
            $adminBody += @"
<p>The following responsible(s) or group member(s) had no e-mail address and
were skipped (the audit mail was still sent to the remaining recipients).</p>
<table border="1" cellpadding="4" cellspacing="0">
<tr><th>Matrix file</th><th>Skipped recipient(s)</th></tr>
$($issueRows -join "`n")
</table>
"@
        }

        $issueKit = & $newMailKitParams `
            -To $scriptAdmin `
            -Subject 'Permission matrix audit report - recipients without an e-mail address' `
            -Body $adminBody -Attachments $null -Bcc $null -Priority 'High'

        try { Send-MailKitMessageHC @issueKit }
        catch { Write-Warning "Failed sending admin recipient-report: $_" }
    }
    #endregion
}