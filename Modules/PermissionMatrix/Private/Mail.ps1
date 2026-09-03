function Get-MailRecipientListHC {
    <#
    .SYNOPSIS
        Build a clean, de-duplicated list of e-mail recipients.

    .DESCRIPTION
        Combines SendMailSettings.To with the optional DefaultsMailTo, drops
        blank entries, trims each address, then sorts and de-duplicates.
        Returns nothing when no valid recipients remain.

    .NOTES
        De-duplication uses Sort-Object -Unique, which compares strings
        case-insensitively, so addresses differing only in casing collapse
        into one entry.

    .EXAMPLE
        $settings = [PSCustomObject]@{ To = 'bob@contoso.com' }
        Get-MailRecipientListHC `
            -SendMailSettings $settings `
            -DefaultsMailTo 'admin@contoso.com', 'bob@contoso.com'

        Returns 'admin@contoso.com' and 'bob@contoso.com'; the duplicate
        collapses.
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
        Build the subject line for a matrix-processing notification e-mail,
        with correct singular/plural wording.

    .DESCRIPTION
        Two mutually exclusive modes:

        - System errors present: reports the matrix file count and the number
          of system errors.
        - No system errors: reports the matrix file count followed by the
          Counter totals, each part added only when its count is above zero,
          so a clean run shows just the matrix file count.

        CustomSubject, when given, is appended to either form.

    .NOTES
        System errors take priority: when SystemErrors holds any items the
        per-matrix error and warning counts are NOT reported, even if Counter
        holds non-zero totals.

    .EXAMPLE
        $counter = [PSCustomObject]@{ TotalErrors = 1; TotalWarnings = 4 }
        Get-MailSubjectHC -SystemErrors @() -Counter $counter -MatrixCount 1

        Returns '1 matrix, 1 error, 4 warnings'.
    #>    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $SystemErrors,

        [Parameter(Mandatory)]
        $Counter,

        [Parameter(Mandatory)]
        $MatrixCount,

        [string]$CustomSubject
    )

    $matrixPlural = if ($MatrixCount -ne 1) { 'es' } else { '' }
    $cSuffix = if ($CustomSubject) { ", $CustomSubject" } else { '' }

    # If system errors exist
    if ($SystemErrors.Count -gt 0) {
        $sysPlural = if ($SystemErrors.Count -ne 1) { 's' } else { '' }
        return "$MatrixCount matrix$matrixPlural, $($SystemErrors.Count) system error$sysPlural$cSuffix"
    }

    # No system errors: embed matrix counts + warnings/errors
    $err = $Counter.TotalErrors
    $warn = $Counter.TotalWarnings
    $fixed = $Counter.TotalFixed

    $errPart = if ($err -gt 0) { ", $err error$(if ($err -ne 1) {'s'})" } else { '' }
    $warnPart = if ($warn -gt 0) { ", $warn warning$(if ($warn -ne 1) {'s'})" } else { '' }
    $fixedPart = if ($fixed -gt 0) { ", $fixed fixed" } else { '' }

    return "$MatrixCount matrix$matrixPlural$errPart$warnPart$fixedPart$cSuffix"
}

function Send-MailKitMessageHC {
    <#
    .SYNOPSIS
        Send an HTML e-mail message through an SMTP server using MailKit/MimeKit.

    .DESCRIPTION
        Loads the MimeKit and MailKit assemblies, builds the message (sender,
        recipients, subject, priority header, HTML body, attachments), connects,
        optionally authenticates, sends, then disposes the SMTP client and any
        open attachment streams.

        The body is always sent as HTML. Authentication happens only when a
        Credential is supplied.

    .PARAMETER Priority
        Mapped to the X-Priority header: High to '1 (Highest)', Normal to
        '3 (Normal)', Low to '5 (Lowest)'.

    .NOTES
        - Attachment paths that fail Test-Path are skipped SILENTLY; a missing
          attachment raises neither an error nor a warning.
        - To and Bcc are both optional. If neither is supplied the message is
          built and sent with no recipients.
        - Requires the MimeKit and MailKit NuGet packages; MimeKit is loaded
          first because MailKit depends on it.

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

function Get-MailBodyLogPathHC {
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

    return Join-Path $LogFolder ('Mail - {0}.html' -f $safeSubject)
}

function Save-MailBodyToLogHC {
    <#
    .SYNOPSIS
        Save an e-mail's HTML body to a log file named after its subject.

    .DESCRIPTION
        Writes MailParams.Body as UTF-8 to '<LogFolder>\Mail - <subject>.html'
        and returns that path. File-name handling is delegated to
        Get-MailBodyLogPathHC.

    .NOTES
        - When the log folder does not exist the function is a SILENT no-op: it
          neither creates the folder nor raises an error, and returns nothing.
        - An existing file with the same name is overwritten (Out-File -Force).
        - Only the Subject and Body properties of MailParams are used.

    .EXAMPLE
        $mail = @{ Subject = 'Results Q1/Q2'; Body = '<p>...</p>' }
        Save-MailBodyToLogHC -MailParams $mail -LogFolder 'C:\Logs'

        Writes to 'C:\Logs\Mail - Results Q1 Q2.html'; the '/' is replaced.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $MailParams,

        [Parameter(Mandatory)]
        $LogFolder
    )

    $path = Get-MailBodyLogPathHC -MailParams $MailParams -LogFolder $LogFolder
    if (-not $path) { return }

    $MailParams.Body | Out-File -LiteralPath $path -Encoding utf8 -Force

    return $path
}