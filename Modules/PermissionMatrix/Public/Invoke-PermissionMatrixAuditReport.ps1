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
        Path to the audit JSON configuration file.

    .PARAMETER ScriptPath
        Same hashtable contract as Invoke-PermissionMatrix. For the audit only
        'PermissionMatrixModule' is required: the Begin stage uses it to
        dot-source the private functions inside its runspaces. The
        SetPermissions / UpdateServiceNow / TestRequirements paths are not
        needed because the Process and End stages never run.

    .PARAMETER SystemErrors
        [ref] collection the Begin stage appends initialization errors to.

    .NOTES
        The admin notification recipients (BCC'd on every audit mail and
        notified about skipped matrices and initialization failures) are read
        from 'AuditReport.ScriptAdmin' in the configuration file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ConfigurationJsonFile,
        [Parameter(Mandatory)][hashtable]$ScriptPath,
        [Parameter(Mandatory)][ref]$SystemErrors
    )

    $context = Invoke-PermissionMatrixBeginHC `
        -ConfigurationJsonFile $ConfigurationJsonFile `
        -ScriptPath $ScriptPath `
        -SystemErrors $SystemErrors

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

    # Matrices skipped because of fatal errors; reported to the admin at the end
    $skipped = [System.Collections.Generic.List[pscustomobject]]::new()

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

        #region Skip matrices without a responsible (no one to send to)
        if (-not $formData -or -not $formData.MatrixResponsible) {
            Write-Verbose "Matrix '$matrixName' has no MatrixResponsible, skipped"
            continue
        }
        #endregion

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

        $attachmentPath = Copy-MatrixFileToLogFolderHC `
            -SourceFilePath $fileResult.Item.FullName `
            -LogFolder $matrixLogFolder `
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
            -Bcc $scriptAdmin

        $mailKit = & $newMailKitParams `
            -To $msg.To -Subject $msg.Subject -Body $msg.Body `
            -Attachments $msg.Attachments -Bcc $msg.Bcc -Priority 'Normal'

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
}