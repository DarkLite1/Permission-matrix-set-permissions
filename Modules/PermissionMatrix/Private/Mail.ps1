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
        Sends email using MailKit.
        Expected parameters:
            - MailKitAssemblyPath
            - MimeKitAssemblyPath
            - SmtpServerName
            - SmtpPort
            - SmtpConnectionType
            - Credential (optional)
            - From
            - FromDisplayName
            - To[]
            - Bcc[]
            - Body (HTML)
            - Subject
            - Attachments[]
            - Priority
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

        # Body
        $bodyPart = New-Object MimeKit.TextPart('html')
        $bodyPart.Text = $Body
        $bodyContainer = New-Object MimeKit.Multipart 'mixed'
        $bodyContainer.Add($bodyPart)

        # Attachments
        foreach ($path in $Attachments) {
            if (Test-Path $path) {
                $file = New-Object MimeKit.MimePart
                $stream = [System.IO.File]::OpenRead($path)
                $attachmentStreams.Add($stream)
                $content = New-Object MimeKit.MimeContent($stream)
                $file.Content = $content
                $file.FileName = [System.IO.Path]::GetFileName($path)
                $file.ContentDisposition = New-Object MimeKit.ContentDisposition
                $file.ContentTransferEncoding = [MimeKit.ContentEncoding]::Base64
                $bodyContainer.Add($file)
            }
        }

        $message.Body = $bodyContainer

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

    $path = Join-Path $LogFolder ('Mail - {0}.html' -f $safeSubject)

    $MailParams.Body | Out-File -LiteralPath $path -Encoding utf8 -Force

    return $path
}