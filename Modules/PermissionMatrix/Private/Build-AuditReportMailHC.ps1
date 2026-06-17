function Build-AuditReportMailHC {
    <#
    .SYNOPSIS
        Builds the e-mail parameters for one matrix audit report mail.

    .DESCRIPTION
        Assembles the recipient list, subject, HTML body and attachment for the
        audit e-mail that is sent to the responsible(s) of a single matrix.

        The subject and the HTML body are taken from the audit configuration
        ('Settings.SendMail.Subject' and 'Settings.SendMail.Body'), so the
        message text lives in the input file instead of in the code. Both
        support {{Token}} placeholders that are replaced with the values of the
        current matrix. Supported tokens:

            {{MatrixFileName}}          {{MatrixFilePath}}
            {{MatrixCategoryName}}      {{MatrixSubCategoryName}}
            {{MatrixFolderPath}}        {{MatrixFolderDisplayName}}
            {{MatrixResponsible}}       {{UniqueUserCount}}
            {{UniqueGroupCount}}        {{RequestTicketURL}}

        The function only assembles parameters; it does not send anything. The
        caller splats the result into Send-MailHC together with the transport
        settings. Keeping the sending out of this function makes it fully
        unit-testable without a mail server.

    .PARAMETER FormData
        The matrix' formatted FormData row
        (e.g. $fileResult.Sheets.FormData.Formatted). Provides MatrixFileName,
        MatrixResponsible, MatrixCategoryName, MatrixFolderPath, ...

    .PARAMETER AccessList
        The matrix' AccessList rows (from Build-MatrixLogSheetRowsHC), used to
        compute the unique user and group counts shown in the mail.

    .PARAMETER AttachmentPath
        Path to the Excel log file to attach. This is the per-matrix copy in
        the log folder created by Copy-MatrixFileToLogFolderHC (the same log
        file the Permission Matrix script produces).

    .PARAMETER MailSettings
        The 'Settings.SendMail' object from the audit config (From,
        FromDisplayName, Subject, Body, Bcc, ...).

    .PARAMETER RequestTicketURL
        Value used for the {{RequestTicketURL}} token.

    .PARAMETER Bcc
        Extra Bcc addresses (e.g. the script admins) merged with the Bcc
        configured in MailSettings.

    .OUTPUTS
        Hashtable with the keys From, FromDisplayName, To, Bcc, Subject, Body
        and Attachments.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$FormData,
        [array]$AccessList = @(),
        [Parameter(Mandatory)][string]$AttachmentPath,
        [Parameter(Mandatory)]$MailSettings,
        [string]$RequestTicketURL,
        [string[]]$Bcc = @()
    )

    #region Unique user / group counts
    $userAccounts = @(
        $AccessList | Where-Object { $_.Type -eq 'user' } |
        Select-Object -ExpandProperty SamAccountName
    )
    $memberAccounts = @(
        $AccessList | Where-Object { $_.MemberSamAccountName } |
        Select-Object -ExpandProperty MemberSamAccountName
    )

    $uniqueUserCount = @(
        $userAccounts + $memberAccounts |
        Where-Object { $_ } | Sort-Object -Unique
    ).Count

    $uniqueGroupCount = @(
        $AccessList | Where-Object { $_.Type -eq 'group' } |
        Select-Object -ExpandProperty Name |
        Where-Object { $_ } | Sort-Object -Unique
    ).Count
    #endregion

    #region Token replacement (literal, so values with '$' are safe)
    $responsible = @(
        "$($FormData.MatrixResponsible)".Split(',') |
        ForEach-Object { $_.Trim() } | Where-Object { $_ }
    )

    $tokens = [ordered]@{
        'MatrixFileName'          = $FormData.MatrixFileName
        'MatrixFilePath'          = $FormData.MatrixFilePath
        'MatrixCategoryName'      = $FormData.MatrixCategoryName
        'MatrixSubCategoryName'   = $FormData.MatrixSubCategoryName
        'MatrixFolderPath'        = $FormData.MatrixFolderPath
        'MatrixFolderDisplayName' = $FormData.MatrixFolderDisplayName
        'MatrixResponsible'       = $responsible -join ', '
        'UniqueUserCount'         = $uniqueUserCount
        'UniqueGroupCount'        = $uniqueGroupCount
        'RequestTicketURL'        = $RequestTicketURL
    }

    $subject = $MailSettings.Subject
    $body = $MailSettings.Body

    foreach ($key in $tokens.Keys) {
        $token = '{{' + $key + '}}'
        $value = [string]$tokens[$key]

        if ($subject) { $subject = $subject.Replace($token, $value) }
        if ($body) { $body = $body.Replace($token, $value) }
    }
    #endregion

    #region Recipients
    $bccAll = @(
        @($MailSettings.Bcc) + @($Bcc) |
        Where-Object { $_ } | Sort-Object -Unique
    )
    #endregion

    return @{
        From            = $MailSettings.From
        FromDisplayName = $MailSettings.FromDisplayName
        To              = $responsible
        Bcc             = $bccAll
        Subject         = $subject
        Body            = $body
        Attachments     = $AttachmentPath
    }
}
