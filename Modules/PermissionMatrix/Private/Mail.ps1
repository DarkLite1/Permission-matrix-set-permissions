function Get-MailRecipientListHC {
    <#
    .SYNOPSIS
        Build a clean, de-duplicated list of e-mail recipients.

    .DESCRIPTION
        Combines the recipients defined in a mail settings object with an
        optional set of default recipients, then returns a single tidied list.

        The combined addresses are processed as follows:
        - $null and empty/blank entries are removed
        - leading and trailing whitespace is trimmed from each address
        - the result is sorted and de-duplicated

        When no valid recipients remain, nothing is returned.

    .PARAMETER SendMailSettings
        An object that exposes a 'To' property holding one or more recipient
        e-mail addresses, typically the mail configuration for a single
        notification or report. Only the 'To' property is read; any other
        properties on the object are ignored.

    .PARAMETER DefaultsMailTo
        One or more default recipient e-mail addresses that should always be
        included, regardless of what is defined in SendMailSettings.To. Useful
        for ensuring a central mailbox or administrator is always added.

    .EXAMPLE
        $settings = [PSCustomObject]@{ 
            To = 'bob@contoso.com', ' jane@contoso.com ' 
        }
        Get-MailRecipientListHC -SendMailSettings $settings

        Returns 'bob@contoso.com' and 'jane@contoso.com'. The whitespace around
        the second address is trimmed.

    .EXAMPLE
        $settings = [PSCustomObject]@{ To = 'bob@contoso.com' }
        Get-MailRecipientListHC `
            -SendMailSettings $settings `
            -DefaultsMailTo 'admin@contoso.com', 'bob@contoso.com'

        Returns 'admin@contoso.com' and 'bob@contoso.com'. The duplicate
        'bob@contoso.com' is collapsed to a single entry.

    .EXAMPLE
        $settings = [PSCustomObject]@{ To = $null }
        Get-MailRecipientListHC -SendMailSettings $settings

        Returns nothing, because there are no valid recipients to list.

    .OUTPUTS
        System.String
        Zero or more unique recipient e-mail addresses, sorted alphabetically.

    .NOTES
        De-duplication is performed by Sort-Object -Unique, which compares
        strings case-insensitively. Addresses differing only in casing are
        therefore collapsed into one entry.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $SendMailSettings,

        $DefaultsMailTo
    )

    $list = @()
    if ($SendMailSettings.To) { $list += $SendMailSettings.To }
    if ($DefaultsMailTo) { $list += $DefaultsMailTo }

    # First Where-Object drops $null / empty entries before .Trim() is called,
    # so a null in the array can no longer throw.
    return (
        $list |
        Where-Object { $_ } |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ }
    ) | Sort-Object -Unique
}

function Get-MailSubjectHC {
    <#
    .SYNOPSIS
        Build the subject line for a matrix-processing notification e-mail.

    .DESCRIPTION
        Produces a single, human-readable subject line that summarises the
        outcome of a matrix-processing run. The wording adapts automatically,
        including correct singular/plural forms.

        The subject is built in one of two mutually exclusive modes:

        - System errors present (SystemErrors.Count greater than 0):
          the subject reports the matrix file count and the number of system
          errors. The per-matrix error and warning counts from Counter are
          intentionally omitted, as the system errors take priority.

        - No system errors:
          the subject reports the matrix file count followed by the total
          error and warning counts from Counter. Each part is only added when
          its count is greater than zero, so a clean run shows just the matrix
          file count.

        In both modes, an optional CustomSubject is appended at the end.

    .PARAMETER SystemErrors
        A collection of system-level errors. Only its Count is used. When the
        count is greater than zero, the subject switches to system-error mode
        and the Counter totals are not included.

    .PARAMETER Counter
        An object exposing the properties TotalErrors and TotalWarnings. These
        are only used when no system errors are present.

    .PARAMETER MatrixCount
        The number of matrix files processed. Always reported in the subject
        and drives the 'matrix file' / 'matrix files' wording.

    .PARAMETER CustomSubject
        Optional free text appended to the end of the subject, prefixed with a
        comma and space (for example ', Nightly run'). When omitted, nothing
        is appended.

    .EXAMPLE
        $counter = [PSCustomObject]@{ TotalErrors = 5; TotalWarnings = 2 }
        Get-MailSubjectHC `
            -SystemErrors @('disk full', 'timeout') `
            -Counter $counter `
            -MatrixCount 3

        Returns '3 matrix files, 2 system errors'. Because system errors are
        present, the error and warning counts from Counter are omitted.

    .EXAMPLE
        $counter = [PSCustomObject]@{ TotalErrors = 1; TotalWarnings = 4 }
        Get-MailSubjectHC -SystemErrors @() -Counter $counter -MatrixCount 1

        Returns '1 matrix file, 1 error, 4 warnings'. No system errors, so the
        Counter totals are included with correct pluralisation.

    .EXAMPLE
        $counter = [PSCustomObject]@{ TotalErrors = 0; TotalWarnings = 0 }
        Get-MailSubjectHC -SystemErrors @() -Counter $counter -MatrixCount 2 -CustomSubject 'Nightly run'

        Returns '2 matrix files, Nightly run'. A clean run shows only the matrix
        file count plus the custom suffix.

    .OUTPUTS
        System.String
        A single subject line describing the run outcome.

    .NOTES
        System errors take priority: when SystemErrors contains any items, the
        per-matrix error and warning counts are not reported, even if Counter
        holds non-zero totals.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $SystemErrors,

        [Parameter(Mandatory)]
        $Counter,

        [Parameter(Mandatory)]
        $MatrixCount,

        [string]$CustomSubject
    )

    $matrixPlural = if ($MatrixCount -ne 1) { 's' } else { '' }
    $cSuffix = if ($CustomSubject) { ", $CustomSubject" } else { '' }

    # If system errors exist
    if ($SystemErrors.Count -gt 0) {
        $sysPlural = if ($SystemErrors.Count -ne 1) { 's' } else { '' }
        return "$MatrixCount matrix file$matrixPlural, $($SystemErrors.Count) system error$sysPlural$cSuffix"
    }

    # No system errors: embed matrix counts + warnings/errors
    $err = $Counter.TotalErrors
    $warn = $Counter.TotalWarnings

    $errPart = if ($err -gt 0) { ", $err error$(if ($err -ne 1) {'s'})" } else { '' }
    $warnPart = if ($warn -gt 0) { ", $warn warning$(if ($warn -ne 1) {'s'})" } else { '' }

    return "$MatrixCount matrix file$matrixPlural$errPart$warnPart$cSuffix"
}

function Send-MailKitMessageHC {
    <#
    .SYNOPSIS
        Send an e-mail message through an SMTP server using MailKit/MimeKit.

    .DESCRIPTION
        Loads the MimeKit and MailKit assemblies, builds an HTML e-mail message
        and sends it through the specified SMTP server.

        The function handles the full lifecycle: it validates and loads the
        required assemblies with clear, actionable errors, constructs the
        message (sender, recipients, subject, priority header, HTML body and
        attachments), connects to the SMTP server, optionally authenticates,
        sends the message, and disposes of the SMTP client and any open
        attachment file streams.

        The message body is always sent as HTML. Authentication is performed
        only when a Credential is supplied; otherwise the message is sent
        unauthenticated.

    .PARAMETER MailKitAssemblyPath
        Full path to MailKit.dll. Required. MailKit depends on MimeKit, which
        is loaded first. A missing or invalid path throws a descriptive error.

    .PARAMETER MimeKitAssemblyPath
        Full path to MimeKit.dll. Required. Loaded before MailKit. A missing or
        invalid path throws a descriptive error.

    .PARAMETER SmtpServerName
        Host name or IP address of the SMTP server.

    .PARAMETER SmtpPort
        TCP port of the SMTP server (for example 25, 587 or 465).

    .PARAMETER Body
        The message body, interpreted as HTML.

    .PARAMETER Subject
        The message subject line.

    .PARAMETER From
        The sender's e-mail address.

    .PARAMETER FromDisplayName
        Optional display name shown for the sender (for example 'IT Support').
        When omitted, only the address is used.

    .PARAMETER To
        One or more recipient e-mail addresses.

    .PARAMETER Bcc
        One or more blind-carbon-copy recipient e-mail addresses.

    .PARAMETER Priority
        Message priority, mapped to the X-Priority header:
        High maps to '1 (Highest)', Normal to '3 (Normal)' and Low to
        '5 (Lowest)'. Defaults to 'Normal'.

    .PARAMETER Attachments
        One or more file paths to attach. Each path is checked with Test-Path;
        paths that do not exist are skipped without raising an error. Existing
        files are attached using Base64 content-transfer encoding.

    .PARAMETER SmtpConnectionType
        The secure-socket option passed to MailKit when connecting. One of
        'None', 'Auto', 'SslOnConnect', 'StartTls' or
        'StartTlsWhenAvailable'. Defaults to 'None'.

    .PARAMETER Credential
        Optional PSCredential used to authenticate with the SMTP server. When
        omitted, the message is sent without authentication.

    .EXAMPLE
        Send-MailKitMessageHC `
            -MailKitAssemblyPath 'C:\lib\MailKit.dll' `
            -MimeKitAssemblyPath 'C:\lib\MimeKit.dll' `
            -SmtpServerName 'smtp.contoso.com' `
            -SmtpPort 25 `
            -From 'noreply@contoso.com' `
            -To 'bob@contoso.com' `
            -Subject 'Report ready' `
            -Body '<p>Your report is ready.</p>'

        Sends a basic unauthenticated HTML message over an unencrypted
        connection.

    .EXAMPLE
        $cred = Get-Credential
        Send-MailKitMessageHC `
            -MailKitAssemblyPath 'C:\lib\MailKit.dll' `
            -MimeKitAssemblyPath 'C:\lib\MimeKit.dll' `
            -SmtpServerName 'smtp.contoso.com' `
            -SmtpPort 587 `
            -SmtpConnectionType 'StartTls' `
            -Credential $cred `
            -From 'noreply@contoso.com' `
            -FromDisplayName 'Contoso Alerts' `
            -To 'bob@contoso.com', 'jane@contoso.com' `
            -Subject 'Nightly run' `
            -Body '<h1>Done</h1>' `
            -Priority 'High'

        Sends an authenticated, StartTLS-encrypted message to two recipients
        with a sender display name and high priority.

    .EXAMPLE
        Send-MailKitMessageHC `
            -MailKitAssemblyPath 'C:\lib\MailKit.dll' `
            -MimeKitAssemblyPath 'C:\lib\MimeKit.dll' `
            -SmtpServerName 'smtp.contoso.com' `
            -SmtpPort 25 `
            -From 'noreply@contoso.com' `
            -To 'bob@contoso.com' `
            -Subject 'Logs' `
            -Body '<p>See attached.</p>' `
            -Attachments 'C:\logs\run.log', 'C:\logs\missing.log'

        Sends a message attaching run.log. The non-existent missing.log is
        silently skipped.

    .OUTPUTS
        None. This function sends an e-mail and returns no output.

    .NOTES
        - Attachment paths that fail Test-Path are skipped silently; no error
          or warning is raised for a missing attachment.
        - To and Bcc are both optional. If neither is supplied the message is
          built and sent with no recipients.
        - Requires the MimeKit and MailKit NuGet packages; provide the paths to
          their DLLs via the assembly-path parameters.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$MailKitAssemblyPath,
        [Parameter(Mandatory)][string]$MimeKitAssemblyPath,
        [Parameter(Mandatory)][string]$SmtpServerName,
        [Parameter(Mandatory)][int]$SmtpPort,
        [Parameter(Mandatory)][string]$Body,
        [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][string]$From,
        [string]$FromDisplayName,
        [string[]]$To,
        [string[]]$Bcc,
        [ValidateSet('High', 'Normal', 'Low')]
        [string]$Priority = 'Normal',
        [string[]]$Attachments,
        [ValidateSet('None', 'Auto', 'SslOnConnect', 'StartTls', 'StartTlsWhenAvailable')]
        [string]$SmtpConnectionType = 'None',
        [PSCredential]$Credential
    )

    # Load assemblies (MimeKit first; MailKit depends on it). Wrap Add-Type so a
    # missing, unset, or invalid DLL produces a clear, actionable error instead
    # of a cryptic assembly-load failure.
    foreach ($assembly in @(
            [PSCustomObject]@{ Name = 'MimeKit'; Path = $MimeKitAssemblyPath },
            [PSCustomObject]@{ Name = 'MailKit'; Path = $MailKitAssemblyPath }
        )) {
        if ([string]::IsNullOrWhiteSpace($assembly.Path)) {
            throw "The $($assembly.Name) assembly path is not set. Provide the full path to $($assembly.Name).dll (e.g. via the configuration file or its environment variable)."
        }

        try {
            Add-Type -Path $assembly.Path -ErrorAction Stop
        }
        catch {
            throw "Failed to load the $($assembly.Name) assembly from '$($assembly.Path)'. $($_.Exception.Message) Verify the path is correct and that the MimeKit and MailKit NuGet packages are installed."
        }
    }

    # Streams opened for attachments must stay open until after Send (MimeKit
    # reads them during Send), then be disposed in the finally block.
    $attachmentStreams = [System.Collections.Generic.List[System.IO.Stream]]::new()
    $smtp = $null

    try {
        # Create message
        $message = New-Object MimeKit.MimeMessage
        $fromAddress = New-Object MimeKit.MailboxAddress($FromDisplayName, $From)
        $message.From.Add($fromAddress)

        # InternetAddressList.Add takes an InternetAddress, not a string, so
        # parse each recipient into a MailboxAddress.
        foreach ($t in $To) {
            $message.To.Add([MimeKit.MailboxAddress]::Parse($t))
        }
        foreach ($b in $Bcc) {
            $message.Bcc.Add([MimeKit.MailboxAddress]::Parse($b))
        }

        $message.Subject = $Subject

        # Priority header
        switch ($Priority) {
            'High' { $message.Headers.Add('X-Priority', '1 (Highest)') }
            'Normal' { $message.Headers.Add('X-Priority', '3 (Normal)') }
            'Low' { $message.Headers.Add('X-Priority', '5 (Lowest)') }
        }

        # HTML body part
        $bodyPart = New-Object MimeKit.TextPart('html')
        $bodyPart.Text = $Body

        $attachmentParts = @(
            foreach ($path in $Attachments) {
                if (Test-Path $path) {
                    $file = New-Object MimeKit.MimePart
                    $stream = [System.IO.File]::OpenRead($path)
                    $attachmentStreams.Add($stream)
                    $file.Content = New-Object MimeKit.MimeContent($stream)
                    $file.FileName = [System.IO.Path]::GetFileName($path)
                    $file.ContentDisposition = New-Object MimeKit.ContentDisposition
                    $file.ContentTransferEncoding = [MimeKit.ContentEncoding]::Base64
                    $file
                }
            }
        )

        if ($attachmentParts.Count -gt 0) {
            $bodyContainer = New-Object MimeKit.Multipart 'mixed'
            $bodyContainer.Add($bodyPart)
            foreach ($file in $attachmentParts) {
                $bodyContainer.Add($file)
            }
            $message.Body = $bodyContainer
        }
        else {
            $message.Body = $bodyPart
        }

        # SMTP client
        $smtp = New-Object MailKit.Net.Smtp.SmtpClient
        $smtp.Connect($SmtpServerName, $SmtpPort, [MailKit.Security.SecureSocketOptions]::$SmtpConnectionType)

        if ($Credential) {
            $smtp.Authenticate($Credential.UserName, $Credential.GetNetworkCredential().Password)
        }

        $smtp.Send($message)
    }
    finally {
        if ($smtp) {
            $smtp.Disconnect($true)
            $smtp.Dispose()
        }

        foreach ($stream in $attachmentStreams) {
            if ($stream) { $stream.Dispose() }
        }
    }
}

function Save-MailBodyToLogHC {
    <#
    .SYNOPSIS
        Save an e-mail's HTML body to a log file named after its subject.

    .DESCRIPTION
        Writes the Body of a mail-parameters object to an .html file in the
        specified log folder. The file name is derived from the mail Subject,
        prefixed with 'Mail - ' and suffixed with '.html'.

        Any character that is invalid in a file name is replaced with a space,
        so subjects containing characters such as ':' or '/' still produce a
        valid file name. If the subject is missing, empty, or consists only of
        characters that get stripped, a timestamp is used in place of the
        subject so a valid, non-empty file name is always produced.

        If the log folder does not exist, the function does nothing: no file is
        written and nothing is returned. When the body is written, the full
        path to the created file is returned. An existing file with the same
        name is overwritten.

    .PARAMETER MailParams
        An object exposing Subject and Body properties. Subject is used to
        build the file name; Body is the content written to the file. Other
        properties on the object are ignored.

    .PARAMETER LogFolder
        Path to the folder in which the log file is created. The folder must
        already exist as a directory; if it does not, the function returns
        without writing anything.

    .EXAMPLE
        $mail = @{ Subject = 'Daily report'; Body = '<p>All good</p>' }
        Save-MailBodyToLogHC -MailParams $mail -LogFolder 'C:\Logs'

        Writes the HTML body to 'C:\Logs\Mail - Daily report.html' and returns
        that path.

    .EXAMPLE
        $mail = @{ Subject = 'Results Q1/Q2'; Body = '<p>...</p>' }
        Save-MailBodyToLogHC -MailParams $mail -LogFolder 'C:\Logs'

        Writes to 'C:\Logs\Mail - Results Q1 Q2.html'. The '/' in the subject,
        which is invalid in a file name, is replaced with a space.

    .EXAMPLE
        $mail = @{ Subject = $null; Body = '<p>...</p>' }
        Save-MailBodyToLogHC -MailParams $mail -LogFolder 'C:\Logs'

        Writes to a file such as 'C:\Logs\Mail - 2026-06-03 142530.html'.
        Because the subject is missing, a timestamp is used in its place.

    .EXAMPLE
        $mail = @{ Subject = 'Daily report'; Body = '<p>All good</p>' }
        Save-MailBodyToLogHC -MailParams $mail -LogFolder 'C:\DoesNotExist'

        Returns nothing and writes nothing, because the log folder does not
        exist.

    .OUTPUTS
        System.String
        The full path to the written log file, or nothing when the log folder
        does not exist.

    .NOTES
        - When the log folder is missing the function is a silent no-op: it
          neither creates the folder nor raises an error.
        - When Subject is null, empty, or reduced to only whitespace after
          invalid characters are stripped, the file name uses a timestamp
          (format 'yyyy-MM-dd HHmmss') in place of the subject.
        - An existing file with the same name is overwritten (Out-File -Force).
        - The file is written as UTF-8.
        - Only the Subject and Body properties of MailParams are used.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $MailParams,

        [Parameter(Mandatory)]
        $LogFolder
    )

    if (-not (Test-Path -LiteralPath $LogFolder -PathType Container)) { return }

    # Replace any character that is invalid in a file name with a space.
    # Splitting on the invalid-char set avoids the $OFS-dependent string cast.
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $safeSubject = ([string]$MailParams.Subject).Split($invalid) -join ' '

    if ([string]::IsNullOrWhiteSpace($safeSubject)) {
        $safeSubject = Get-Date -Format 'yyyy-MM-dd HHmmss'
    }

    $path = Join-Path $LogFolder ('Mail - {0}.html' -f $safeSubject)

    $MailParams.Body | Out-File -LiteralPath $path -Encoding utf8 -Force

    return $path
}