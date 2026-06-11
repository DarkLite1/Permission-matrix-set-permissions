# HtmlMailBody.ps1
# Builds the summary email body. Depends on HtmlCommon.ps1 for $Script:Theme
# and shared primitives/builders (New-PillHtmlHC, Build-FileLevelCheckRowHC,
# Format-IssueCountLabelHC, ConvertTo-FileUrlHC). Load HtmlCommon.ps1 first.

function Build-SystemErrorsBlockHC {
    <#
        .DESCRIPTION
            Renders the detailed cards for script-level system errors and
            warnings — the items collected in $SystemErrors throughout the
            run. Errors get a red stripe and ✖ glyph; warnings get an amber
            stripe and ⚠ glyph. Anything that isn't a 'FatalError' or
            'Warning' is ignored.
    #>
    param([array]$SystemErrors)

    if (-not $SystemErrors -or $SystemErrors.Count -eq 0) { return '' }

    $items = @($SystemErrors | Where-Object {
            $_.Type -eq 'FatalError' -or $_.Type -eq 'Warning'
        })
    if ($items.Count -eq 0) { return '' }

    $rows = ''
    foreach ($item in $items) {
        $isFatal = ($item.Type -eq 'FatalError')

        if ($isFatal) {
            $bgColor = $Script:Theme.StatusError
            $accentColor = $Script:Theme.AccentError
            $glyph = '✖'
            $pillText = 'System Error'
            $pillBg = $Script:Theme.AccentSystem
        }
        else {
            $bgColor = $Script:Theme.StatusWarning
            $accentColor = $Script:Theme.AccentWarning
            $glyph = '⚠'
            $pillText = 'System Warning'
            $pillBg = $Script:Theme.AccentWarning
        }

        $name = [System.Net.WebUtility]::HtmlEncode(
            (Get-StringOrDefaultHC $item.Name 'Unnamed item')
        )
        $msg = [System.Net.WebUtility]::HtmlEncode(
            (Get-StringOrDefaultHC $item.Message (Get-StringOrDefaultHC $item.Description ''))
        )
        $category = [System.Net.WebUtility]::HtmlEncode(
            (Get-StringOrDefaultHC $item.Category '')
        )

        $catHtml = ''
        if ($category) {
            $catHtml = "<span style='display:inline-block; margin-right:8px; padding:1px 8px; background-color:$($Script:Theme.BgAlt); border:1px solid $($Script:Theme.BorderLight); border-radius:10px; font-size:10px; font-weight:600; color:$($Script:Theme.TextMuted); text-transform:uppercase; letter-spacing:0.5px;'>$category</span>"
        }

        $pill = New-PillHtmlHC -Text $pillText -Bg $pillBg

        $rows += @"
<tr>
    <td style='padding:0 0 8px 0;'>
        <div class="rr-syscard" style="display:flex; align-items:center; gap:16px; background-color:$bgColor; border-left:3px solid $accentColor; border-radius:6px; padding:10px 14px;">
            <span style='flex:0 0 auto; color:$accentColor; font-size:16px; padding-right:6px; font-weight:bold; line-height:1;'>$glyph</span>
            <span class="rr-syscard-body" style='flex:1 1 auto; min-width:0;'>
                <span style='display:block; margin-bottom:4px;'>$catHtml<span style='font-weight:700; color:$($Script:Theme.TextMain); font-size:13px;'>$name</span></span>
                <span style='display:block; color:$($Script:Theme.TextMuted); font-size:12px; line-height:1.5; font-family:$($Script:Theme.MonoStack); overflow-wrap:anywhere;'>$msg</span>
            </span>
            <span class="rr-syscard-status" style='flex:0 0 auto;'>$pill</span>
        </div>
    </td>
</tr>
"@
    }

    # Section header — pluralized and labeled to match what's actually rendered.
    $errCount = @($items | Where-Object Type -EQ 'FatalError').Count
    $warnCount = @($items | Where-Object Type -EQ 'Warning').Count
    $labelParts = @()
    if ($errCount -gt 0) { $labelParts += "$errCount Error" + $(if ($errCount -ne 1) { 's' }) }
    if ($warnCount -gt 0) { $labelParts += "$warnCount Warning" + $(if ($warnCount -ne 1) { 's' }) }
    $headerLabel = 'System Issues (' + ($labelParts -join ', ') + ')'

    return @"
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; margin:0 0 20px 0; table-layout:fixed; width:100%; max-width:100%;">
    <tr>
        <td style='padding:0 0 8px 0; font-size:11px; font-weight:700; color:$($Script:Theme.TextLight); letter-spacing:1.5px; text-transform:uppercase;'>$headerLabel</td>
    </tr>
    $rows
</table>
"@
}

function Build-SettingsRowHC {
    param([object]$MatrixItem)

    # Determine row status — first check type wins (sorted by severity in upstream code).
    $types = @($MatrixItem.Check.Type)
    $firstType = if ($types.Count -gt 0) { $types[0] } else { 'Info' }

    if ($firstType -eq 'FatalError') {
        $accent = $Script:Theme.AccentError
    }
    elseif ($firstType -eq 'Warning') {
        $accent = $Script:Theme.AccentWarning
    }
    else {
        $accent = $Script:Theme.AccentSuccess
    }

    $err = @($MatrixItem.Check | Where-Object Type -EQ 'FatalError').Count
    $warn = @($MatrixItem.Check | Where-Object Type -EQ 'Warning').Count

    $comp = [System.Net.WebUtility]::HtmlEncode((Get-StringOrDefaultHC $MatrixItem.Setting.Formatted.ComputerName ''))

    $pathRaw = Get-StringOrDefaultHC $MatrixItem.Setting.Formatted.Path ''
    # Show the full path — it sits on its own line and wraps via CSS
    # (overflow-wrap:anywhere) instead of being truncated server-side, so long
    # paths are fully visible rather than clipped by the viewport.
    $pathDisp = [System.Net.WebUtility]::HtmlEncode($pathRaw)
    $pathTitle = if ($pathRaw) {
        " title=`"$([System.Net.WebUtility]::HtmlEncode($pathRaw))`""
    }
    else { '' }

    $action = [System.Net.WebUtility]::HtmlEncode((Get-StringOrDefaultHC $MatrixItem.Setting.Formatted.Action ''))

    $dur = if ($MatrixItem.JobTime.Duration) {
        '{0:00}:{1:00}:{2:00}' -f $MatrixItem.JobTime.Duration.Hours, $MatrixItem.JobTime.Duration.Minutes, $MatrixItem.JobTime.Duration.Seconds
    }
    else { 'N/A' }

    $link = if ($MatrixItem.FileContext.ReportFilePath) {
        [System.Net.WebUtility]::HtmlEncode($MatrixItem.FileContext.ReportFilePath)
    }
    else { '#' }

    # Always reserve the pill cell so columns align even when this row has no issues.
    $pillHtml = if ($err -gt 0) {
        New-PillHtmlHC -Text 'Error' -Bg $Script:Theme.AccentError
    }
    elseif ($warn -gt 0) {
        New-PillHtmlHC -Text 'Warning' -Bg $Script:Theme.AccentWarning
    }
    else { '&nbsp;' }

    # Fluid flex card. The status pill sits on the RIGHT, but unlike the old
    # fixed-column table it can no longer clip: it is flex:0 0 auto (never
    # shrinks) and the PATH is the flexible element that absorbs horizontal
    # overflow via ellipsis. On very narrow panes the whole row wraps (see the
    # .rr-srow @media rule in the stylesheet). New Outlook / Outlook-on-the-web
    # render this with a modern Chromium engine, so flexbox and max-width
    # behave as in a browser.
    #
    # $pillCell is empty (no pill) for clean rows; we omit the element entirely
    # rather than reserving width, since flex handles alignment without it.
    $pillCell = if ($pillHtml -and $pillHtml -ne '&nbsp;') {
        "<span class='rr-srow-status' style='flex:0 0 auto; margin-left:6px;'>$pillHtml</span>"
    }
    else { '' }

    return @"
<a href='$link' target='_blank' rel='noopener noreferrer' class='rr-srow' style='display:flex; align-items:center; gap:16px; text-decoration:none; color:inherit; background-color:$($Script:Theme.BgWhite); border:1px solid $($Script:Theme.BorderLight); border-left:3px solid $accent; border-radius:6px; padding:10px 14px; margin:0 0 8px 0;'>
    <span style='flex:0 0 auto; width:8px; height:8px;  margin-right:6px;background-color:$accent; border-radius:50%;'></span>
    <span class='rr-srow-ident' style='flex:1 1 auto; min-width:0;'>
        <span style='display:block; font-weight:700; color:$($Script:Theme.TextMain); font-size:13px;'>$comp</span>
        <span class='rr-srow-path' style='display:block; font-family:$($Script:Theme.MonoStack); font-size:11px; color:$($Script:Theme.TextMuted); white-space:normal; overflow-wrap:anywhere; word-break:break-all;'$pathTitle>$pathDisp</span>
    </span>
    <span class='rr-srow-meta' style='flex:0 0 auto; color:$($Script:Theme.TextLight); font-size:11px; white-space:nowrap;'>
        <span style='margin-right:14px;'>$action</span>
        <span style='font-family:$($Script:Theme.MonoStack);'>$dur</span>
    </span>
    $pillCell
</a>
"@
}

function Build-MatrixFileCardHC {
    param([object]$FileContext)

    # File header info
    $fileName = [System.Net.WebUtility]::HtmlEncode($FileContext.Item.Name)

    $lastChangeInfo = Format-LastChangeHC `
        -LastModifiedBy $FileContext.ExcelInfo.LastModifiedBy `
        -Modified $FileContext.ExcelInfo.Modified

    <#
     Two distinct links live in this card:
        1. $matrixLink — opens the source .xlsx file directly. Used by the
        filename in the gradient header. When the file was archived
        (Matrix.Archive = true), it no longer exists at its original
        location, so we link to the archived copy instead.
        2. $reportLink — opens the standalone execution report HTML. Used by
        the "Open full report &rarr;" footer link.
    #>
    $matrixPath = if (
        $FileContext.PSObject.Properties.Match('ArchivedPath').Count -and
        -not [string]::IsNullOrWhiteSpace($FileContext.ArchivedPath)
    ) {
        $FileContext.ArchivedPath
    }
    else {
        Get-StringOrDefaultHC $FileContext.Item.FullName ''
    }
    $matrixLink = if ($matrixPath) {
        [System.Net.WebUtility]::HtmlEncode((ConvertTo-FileUrlHC $matrixPath))
    }
    else { '#' }

    # Tooltip text shown when hovering the filename
    $matrixTitle = if ($matrixPath) { [System.Net.WebUtility]::HtmlEncode($matrixPath) } else { '' }

    $reportLink = if ($FileContext.ReportFilePath) {
        [System.Net.WebUtility]::HtmlEncode((ConvertTo-FileUrlHC $FileContext.ReportFilePath))
    }
    elseif ($matrixPath) {
        # Fall back to the matrix file if no report was written
        [System.Net.WebUtility]::HtmlEncode((ConvertTo-FileUrlHC $matrixPath))
    }
    else { '#' }

    # Tally checks across all sources to decide header color and summary text
    $allChecks = @()
    if ($FileContext.Check) { $allChecks += $FileContext.Check }
    if ($FileContext.Sheets.FormData.Check) { $allChecks += $FileContext.Sheets.FormData.Check }
    if ($FileContext.Sheets.Permissions.Check) { $allChecks += $FileContext.Sheets.Permissions.Check }
    if ($FileContext.Matrices) {
        foreach ($m in $FileContext.Matrices) {
            if ($m.Check) { $allChecks += $m.Check }
        }
    }
    $fileErrs = @($allChecks | Where-Object Type -EQ 'FatalError').Count
    $fileWarns = @($allChecks | Where-Object Type -EQ 'Warning').Count

    if ($fileErrs -gt 0) {
        $headerSymbol = '✖'
        $gradFrom, $gradTo = $Script:Theme.GradError
    }
    elseif ($fileWarns -gt 0) {
        $headerSymbol = '⚠'
        $gradFrom, $gradTo = $Script:Theme.GradWarning
    }
    else {
        $headerSymbol = '✓'
        $gradFrom, $gradTo = $Script:Theme.GradSuccess
    }

    $headerLabel = Format-IssueCountLabelHC -Errors $fileErrs -Warnings $fileWarns
    $headerLabelHtml = "<span style=`"font-size:12px; font-weight:700; color:rgba(255,255,255,0.95); text-transform:uppercase; letter-spacing:0.5px;`">$headerLabel</span>"

    # ---- Body content: file-level issues + settings table ----
    $contentRows = ''

    # File-level check groups (Excel file / FormData / Permissions sheets)
    $fileLevelCount = 0
    $fileLevelGroups = @(
        @{ Label = 'Excel File'; Checks = $FileContext.Check }
        @{ Label = 'FormData Sheet'; Checks = $FileContext.Sheets.FormData.Check }
        @{ Label = 'Permissions Sheet'; Checks = $FileContext.Sheets.Permissions.Check }
    )
    foreach ($g in $fileLevelGroups) {
        if ($g.Checks) { $fileLevelCount += @($g.Checks).Count }
    }

    if ($fileLevelCount -gt 0) {
        $contentRows += @"
<tr>
    <td style='padding:14px 16px 6px 16px; font-size:11px; font-weight:700; color:$($Script:Theme.TextLight); letter-spacing:1.5px; text-transform:uppercase;'>File Issues ($fileLevelCount)</td>
</tr>
"@
        foreach ($g in $fileLevelGroups) {
            if ($g.Checks) {
                foreach ($c in $g.Checks) {
                    $contentRows += Build-FileLevelCheckRowHC -Check $c -SheetLabel $g.Label
                }
            }
        }
    }

    # Settings rows — each row is a self-contained fluid flex card
    if ($FileContext.Matrices -and $FileContext.Matrices.Count -gt 0) {
        $sortedMatrices = $FileContext.Matrices |
        Sort-Object { $_.Setting.Formatted.ComputerName }, { $_.Setting.Formatted.Path }, { $_.ID }

        $settingsRowsHtml = ''
        foreach ($m in $sortedMatrices) {
            $settingsRowsHtml += Build-SettingsRowHC -MatrixItem $m
        }

        $matrixCount = @($sortedMatrices).Count
        $contentRows += @"
<tr>
    <td style='padding:14px 16px 6px 16px; font-size:11px; font-weight:700; color:$($Script:Theme.TextLight); letter-spacing:1.5px; text-transform:uppercase;'>Settings ($matrixCount)</td>
</tr>
<tr>
    <td style='padding:0 16px;'>
        $settingsRowsHtml
    </td>
</tr>
"@
    }
    elseif ($fileLevelCount -eq 0) {
        # No file-level issues AND no settings rows — rare but possible
        $contentRows = @"
<tr>
    <td style='padding:14px 16px; font-size:12px; color:$($Script:Theme.TextLight); font-style:italic;'>
        No settings rows were processed for this file.
    </td>
</tr>
"@
    }

    return @"
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:separate; margin:0 0 24px 0; table-layout:fixed; width:100%; max-width:100%; background-color:$($Script:Theme.BgWhite); border:1px solid $($Script:Theme.BorderLight); border-radius:10px; overflow:hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.06);">
    <tr>
        <td bgcolor="$gradTo" style='padding:0; background-color:$gradTo; background-image: linear-gradient(135deg, $gradFrom 0%, $gradTo 100%); border-bottom:1px solid $($Script:Theme.BorderLight);'>
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;">
                <tr>
                    <td valign='middle' width='34' style='padding:14px 0 14px 18px; font-size:20px; font-weight:bold; color:#ffffff; line-height:1; text-align:left;'>$headerSymbol</td>
                    <td valign='middle' style='padding:14px 10px 14px 4px;'>
                        <div style='font-size:16px; font-weight:700; color:#ffffff; line-height:1.25;'>
                            <a href="$matrixLink" title="$matrixTitle" style="color:#ffffff; text-decoration:none;">$fileName</a>
                        </div>
                        <div style='font-size:12px; color:rgba(255,255,255,0.85); line-height:1.4; margin-top:2px;font-style:italic;'>
                            $lastChangeInfo
                        </div>
                    </td>
                    <td valign='middle' align='right' style='padding:14px 18px 14px 10px; white-space:nowrap;'>$headerLabelHtml</td>
                </tr>
            </table>
        </td>
    </tr>
    $contentRows
    <tr>
        <td style='padding:6px 16px 14px 16px; text-align:center; font-size:12px; color:$($Script:Theme.TextLight);'>
            <a href='$reportLink' target='_blank' rel='noopener noreferrer' style='color:$($Script:Theme.LinkColor); text-decoration:none; font-weight:600;'>Open full report &rarr;</a>
        </td>
    </tr>
</table>
"@
}

function Build-MatrixEmailHtmlHC {
    param(
        [Parameter(Mandatory)][array]$FileResults,
        [Parameter(Mandatory)][hashtable]$Html
    )

    $output = ''
    foreach ($fileContext in $FileResults) {
        $output += Build-MatrixFileCardHC -FileContext $fileContext
    }
    return $output
}

function Get-MailBodyHtmlHC {
    param(
        $Settings,
        $Html,
        $ExportedFiles,
        $AttNote,
        [datetime]$ScriptStartTime,
        [datetime]$ScriptEndTime = (Get-Date),
        $LogFolder
    )

    $scriptName = [System.Net.WebUtility]::HtmlEncode(
        (Get-StringOrDefaultHC $Settings.ScriptName 'Permission Matrix')
    )
    $userBody = Get-StringOrDefaultHC $Settings.SendMail.Body ''
    $bodyWidth = $Script:Theme.BodyWidth

    # Resolve system errors from $Html.SystemErrors if supplied. Accepts a
    # [ref] (e.g. $SystemErrors from Invoke-PermissionMatrixEndHC), a plain
    # array, or nothing. Absence is fine — block just isn't rendered.
    $sysErrArr = @()
    if ($Html.SystemErrors) {
        $sysErrArr = if ($Html.SystemErrors -is [System.Management.Automation.PSReference]) {
            @($Html.SystemErrors.Value)
        }
        else {
            @($Html.SystemErrors)
        }
    }
    $systemErrorsBlock = Build-SystemErrorsBlockHC -SystemErrors $sysErrArr

    # ---- Footer with run timing: Started · Ended · Duration ----
    # Compute duration here so callers don't have to format a TimeSpan themselves.
    # All three fields are rendered as label/value pairs, matching the
    # metadata grid style used elsewhere in the email and report.
    $footer = ''
    if ($ScriptStartTime) {
        $startStr = $ScriptStartTime.ToString('dd/MM/yyyy HH:mm')
        $endStr = $ScriptEndTime.ToString('dd/MM/yyyy HH:mm')
        $span = $ScriptEndTime - $ScriptStartTime
        $durStr = '{0:00}:{1:00}:{2:00}' -f $span.Hours, $span.Minutes, $span.Seconds

        # Small helper for footer label/value pairs (kept local — no other
        # caller needs this exact layout).
        $renderField = {
            param([string]$Label, [string]$Value)
            "<span style='display:inline-block; margin-right:18px;'>" +
            "<strong style='color:$($Script:Theme.TextLight); font-weight:700; " +
            "text-transform:uppercase; letter-spacing:0.5px; margin-right:5px; font-size:10px;'>" +
            "$Label</strong>" +
            "<span style='font-size:11px; color:$($Script:Theme.TextLight); " +
            "font-family:$($Script:Theme.MonoStack);'>$Value</span>" +
            '</span>'
        }

        $startedHtml = & $renderField 'Started' ([System.Net.WebUtility]::HtmlEncode($startStr))
        $endedHtml = & $renderField 'Ended' ([System.Net.WebUtility]::HtmlEncode($endStr))
        $durationHtml = & $renderField 'Duration' ([System.Net.WebUtility]::HtmlEncode($durStr))

        $footer = "<p style='margin:16px 0 0 0; text-align:center;'>$startedHtml$endedHtml$durationHtml</p>"
    }

    @"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
$($Html.Style)
</head>
<body>
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; background-color:$($Script:Theme.BgPage);">
    <tr>
        <td align="left" valign="top" style="padding:0;">
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; width:100%; max-width:${bodyWidth}px;">
                <tr><td style="padding:0 0 4px 0;"><h1>$scriptName</h1></td></tr>
                <tr><td style="padding:0 0 16px 0; color:$($Script:Theme.TextMuted); font-size:13px; line-height:1.6;">$userBody</td></tr>
                <tr><td style="padding:0;">$($Html.ErrorWarningTable)</td></tr>
                <tr><td style="padding:0;">$systemErrorsBlock</td></tr>
                <tr><td style="padding:0;">$($Html.MatrixTables)</td></tr>
                <tr><td style="padding:0;">$footer</td></tr>
            </table>
        </td>
    </tr>
</table>
</body>
</html>
"@
}

