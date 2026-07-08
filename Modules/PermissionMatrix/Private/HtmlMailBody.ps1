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
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" class="rr-syscard" bgcolor="$bgColor" style="border-collapse:separate; width:100%; max-width:100%; background-color:$bgColor; border-left:3px solid $accentColor; border-radius:6px;">
            <tr>
                <td valign="top" width="26" style='padding:12px 0 12px 14px; color:$accentColor; font-size:16px; font-weight:bold; line-height:1;'>$glyph</td>
                <td valign="middle" class="rr-syscard-body" style='padding:10px 12px;'>
                    <span style='display:block; margin-bottom:4px;'>$catHtml<span style='font-weight:700; color:$($Script:Theme.TextMain); font-size:13px;'>$name</span></span>
                    <span style='display:block; color:$($Script:Theme.TextMuted); font-size:12px; line-height:1.5; font-family:$($Script:Theme.MonoStack); overflow-wrap:anywhere; word-break:break-word;'>$msg</span>
                </td>
                <td valign="middle" align="right" class="rr-syscard-status" style='padding:10px 14px 10px 6px; white-space:nowrap;'>$pill</td>
            </tr>
        </table>
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

function Build-MailTopLinksBlockHC {
    param(
        [string]$BrowserViewFilePath,
        $ExportedFiles
    )

    $linkStyle = "color:$($Script:Theme.LinkColor); text-decoration:none; font-weight:600;"
    $mutedStyle = "color:$($Script:Theme.TextMuted); font-size:12px; line-height:1.45;"
    $rows = ''

    if (-not [string]::IsNullOrWhiteSpace($BrowserViewFilePath)) {
        $browserUrl = [System.Net.WebUtility]::HtmlEncode((ConvertTo-FileUrlHC $BrowserViewFilePath))
        $browserTitle = [System.Net.WebUtility]::HtmlEncode($BrowserViewFilePath)

        $rows += @"
<tr>
    <td style='padding:0 0 8px 0; $mutedStyle'>If this mail is not visible, please <a href='$browserUrl' title="$browserTitle" target='_blank' rel='noopener noreferrer' style='$linkStyle'>click here to view it in the browser</a>.</td>
</tr>
"@
    }

    $exportLinks = [System.Collections.Generic.List[string]]::new()
    $exportMap = @(
        @{ Property = 'Permissions'; Label = 'Permissions Excel' }
        @{ Property = 'FormData'; Label = 'ServiceNow FormData Excel' }
        @{ Property = 'OverviewHtml'; Label = 'Overview HTML' }
    )

    foreach ($item in $exportMap) {
        $path = $null
        if ($ExportedFiles -is [System.Collections.IDictionary] -and $ExportedFiles.Contains($item.Property)) {
            $path = Get-StringOrDefaultHC $ExportedFiles[$item.Property] ''
        }
        elseif ($ExportedFiles) {
            $prop = $ExportedFiles.PSObject.Properties[$item.Property]
            if ($prop) {
                $path = Get-StringOrDefaultHC $prop.Value ''
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($path)) {
            $url = [System.Net.WebUtility]::HtmlEncode((ConvertTo-FileUrlHC $path))
            $title = [System.Net.WebUtility]::HtmlEncode($path)
            $label = [System.Net.WebUtility]::HtmlEncode($item.Label)
            $exportLinks.Add("<a href='$url' title=`"$title`" target='_blank' rel='noopener noreferrer' style='$linkStyle'>$label</a>")
        }
    }

    if ($exportLinks.Count -gt 0) {
        $linksHtml = $exportLinks -join "<span style='color:$($Script:Theme.TextLight); padding:0 8px;'>&middot;</span>"
        $rows += @"
<tr>
    <td style='padding:0; $mutedStyle'>Export files: $linksHtml</td>
</tr>
"@
    }

    if (-not $rows) { return '' }

    return @"
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; margin:0 0 14px 0; table-layout:fixed; width:100%; max-width:100%;">
    $rows
</table>
"@
}

function Build-SettingsRowHC {
    param(
        [object]$MatrixItem,
        # When the parent matrix file hit a file-level FatalError its settings
        # never executed, so such rows must not read as green "success". They
        # are shown as grey "Skipped" instead — unless the row carries its own
        # error/warning, which still wins.
        [bool]$FileHasError = $false
    )

    $err = @($MatrixItem.Check | Where-Object Type -EQ 'FatalError').Count
    $warn = @($MatrixItem.Check | Where-Object Type -EQ 'Warning').Count

    # Determine row status — a row's own error/warning wins; otherwise a
    # file-level error downgrades the row to "Skipped" (grey); only a clean row
    # in a successfully processed file stays green.
    $isSkipped = $false
    if ($err -gt 0) {
        $accent = $Script:Theme.AccentError
    }
    elseif ($warn -gt 0) {
        $accent = $Script:Theme.AccentWarning
    }
    elseif ($FileHasError) {
        $accent = $Script:Theme.AccentSkipped
        $isSkipped = $true
    }
    else {
        $accent = $Script:Theme.AccentSuccess
    }

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
    $pillText = $null
    $pillBg = $null
    if ($err -gt 0) {
        $pillText = 'Error'; $pillBg = $Script:Theme.AccentError
    }
    elseif ($warn -gt 0) {
        $pillText = 'Warning'; $pillBg = $Script:Theme.AccentWarning
    }
    elseif ($isSkipped) {
        $pillText = 'Skipped'; $pillBg = $Script:Theme.AccentSkipped
    }

    # Table-based card so Outlook Classic (Word rendering engine, which has no
    # flexbox support) shows the identifier, metadata and status pill as proper
    # aligned columns instead of collapsing them onto a single line. Browsers
    # render the same table identically. The colored status dot uses a bullet
    # glyph (renders in Word, unlike a sized inline-block), and the accent
    # left-border conveys status even where border-radius is ignored.
    #
    # The status pill always gets its own fixed-width cell on the far right, even
    # for clean rows (renders '&nbsp;'), so the pill column lines up vertically
    # across every row and no metadata ever sits between/after the pills. The
    # 'rr-srow-status' class is only added when a pill is actually present (a test
    # asserts clean rows don't carry it), but the empty cell still reserves the
    # column width.
    #
    # Outlook (Word) and the browser need DIFFERENT pill cells, gated by
    # conditional comments so each client renders exactly one:
    #  * Outlook: an inline VML shape is BASELINE-aligned and Word does NOT
    #    vertically centre a nested table via valign (it top-anchors a block that
    #    is shorter than the row). Fix: a 3-row nested table with top/bottom spacer
    #    cells around the pill, sized so the WHOLE nested table is the TALLEST cell
    #    in the row (taller than the identifier cell). Being tallest, it drives the
    #    row height. The spacers are slightly asymmetric (8px top / 4px bottom) to
    #    nudge the pill down a couple of px, since the VML shape otherwise sits a
    #    touch high on its baseline. Table = 8+26+4 = 38px > identifier cell (29px
    #    text + 8px padding = 37px). The pill cell's line-height:26px fully contains
    #    the pill so it is not clipped.
    #  * Browser: the CSS span centres itself against the natural font baseline,
    #    so it keeps a normal font-size (line-height:16px lets the span's own box
    #    drive the line) and stays perfectly centred as before.
    if ($pillText) {
        $pillParts = New-PillHtmlHC -Text $pillText -Bg $pillBg -AsParts
        $msoSpacerTop = "<tr><td height='8' style='font-size:0; line-height:8px; mso-line-height-rule:exactly; padding:0;'>&#160;</td></tr>"
        $msoSpacerBottom = "<tr><td height='4' style='font-size:0; line-height:4px; mso-line-height-rule:exactly; padding:0;'>&#160;</td></tr>"
        $pillTd =
        "<!--[if mso]><td valign='middle' align='center' class='rr-srow-status' width='84' style='vertical-align:middle; padding:0 8px;'><table role='presentation' align='center' cellpadding='0' cellspacing='0' border='0' style='border-collapse:collapse;'>$msoSpacerTop<tr><td style='padding:0; font-size:0; line-height:26px; mso-line-height-rule:exactly;'>$($pillParts.Mso)</td></tr>$msoSpacerBottom</table></td><![endif]-->" +
        "<!--[if !mso]><!--><td valign='middle' align='right' class='rr-srow-status' width='84' style='vertical-align:middle; padding:4px 12px 4px 4px; white-space:nowrap; line-height:16px;'>$($pillParts.Browser)</td><!--<![endif]-->"
    }
    else {
        # Clean row: reserve the column width with a single empty cell.
        $pillTd = "<td valign='middle' align='right' width='84' style='vertical-align:middle; padding:4px 12px 4px 4px; white-space:nowrap; line-height:16px;'>&nbsp;</td>"
    }

    return @"
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" class="rr-srow" style="border-collapse:separate; width:100%; max-width:100%; margin:0 0 4px 0; table-layout:fixed; background-color:$($Script:Theme.BgWhite); border:1px solid $($Script:Theme.BorderMain); border-left:3px solid $accent; border-radius:6px;">
    <tr>
        <td valign='middle' width='20' style='vertical-align:middle; padding:4px 0 4px 12px; color:$accent; font-size:12px; line-height:15px; mso-line-height-rule:exactly;'>&#9679;</td>
        <td valign='middle' class='rr-srow-ident' style='vertical-align:middle; padding:4px 8px;'>
            <table role='presentation' cellpadding='0' cellspacing='0' border='0' width='100%' style='border-collapse:collapse; table-layout:fixed;'>
                <tr>
                    <td valign='top' style='padding:0; font-weight:700; color:$($Script:Theme.TextMain); font-size:13px; line-height:15px; mso-line-height-rule:exactly;'>
                        <a href='$link' target='_blank' rel='noopener noreferrer' style='text-decoration:none; color:$($Script:Theme.TextMain);'>$comp</a>
                    </td>
                </tr>
                <tr>
                    <td valign='top' class='rr-srow-path' style='padding:0; font-family:$($Script:Theme.MonoStack); font-size:11px; color:$($Script:Theme.TextMuted); line-height:14px; mso-line-height-rule:exactly; white-space:normal; overflow-wrap:anywhere; word-break:break-all;'$pathTitle>
                        <a href='$link' target='_blank' rel='noopener noreferrer' style='text-decoration:none; color:$($Script:Theme.TextMuted);'>$pathDisp</a>
                    </td>
                </tr>
            </table>
        </td>
        <td valign='middle' align='right' class='rr-srow-meta' width='104' style='vertical-align:middle; padding:4px 10px; color:$($Script:Theme.TextLight); font-size:11px; line-height:15px; mso-line-height-rule:exactly; white-space:nowrap;'>
            <span style='margin-right:14px;'>$action</span>
            <span style='font-family:$($Script:Theme.MonoStack);'>$dur</span>
        </td>
        $pillTd
    </tr>
</table>
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
     Two link locations live in this card:
        1. $matrixLink — opens the source .xlsx file directly. Used by the
        filename in the gradient header. When the file was archived
        (Matrix.Archive = true), it no longer exists at its original
        location, so we link to the archived copy instead.
        2. The footer link list — built further down, after the check
        tally. It links to the execution report and to the matrix Excel
        copy in the log folder.
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

    <#
     The footer offers direct links to the artifacts in the log folder:
        1. The standalone execution report ('00 - Execution Report.html').
        2. The matrix log copy: the processed matrix Excel file, including the
           'AccessList', 'GroupManagers' and 'AdObjects' sheets
           ($FileContext.LogMatrixFilePath, set by the END stage).
     Each link is only rendered when its file was actually created. When
     neither exists, fall back to the source/archived matrix file so the
     footer is never empty.
    #>
    $footerLinkStyle = "color:$($Script:Theme.LinkColor); text-decoration:none; font-weight:600;"
    $footerLinks = [System.Collections.Generic.List[string]]::new()

    if ($FileContext.ReportFilePath) {
        $reportLink = [System.Net.WebUtility]::HtmlEncode((ConvertTo-FileUrlHC $FileContext.ReportFilePath))
        $reportTitle = [System.Net.WebUtility]::HtmlEncode($FileContext.ReportFilePath)
        $footerLinks.Add("<a href='$reportLink' title=`"$reportTitle`" target='_blank' rel='noopener noreferrer' style='$footerLinkStyle'>Open execution report &rarr;</a>")
    }

    $logMatrixPath = if (
        $FileContext.PSObject.Properties.Match('LogMatrixFilePath').Count -and
        -not [string]::IsNullOrWhiteSpace($FileContext.LogMatrixFilePath)
    ) {
        $FileContext.LogMatrixFilePath
    }
    if ($logMatrixPath) {
        $excelLink = [System.Net.WebUtility]::HtmlEncode((ConvertTo-FileUrlHC $logMatrixPath))
        $excelTitle = [System.Net.WebUtility]::HtmlEncode($logMatrixPath)
        $footerLinks.Add("<a href='$excelLink' title=`"$excelTitle`" target='_blank' rel='noopener noreferrer' style='$footerLinkStyle'>Open matrix log copy &rarr;</a>")
    }

    if ($footerLinks.Count -eq 0 -and $matrixPath) {
        # Fall back to the source/archived matrix file if no log
        # artifacts were written
        $footerLinks.Add("<a href='$matrixLink' title=`"$matrixTitle`" target='_blank' rel='noopener noreferrer' style='$footerLinkStyle'>Open matrix file &rarr;</a>")
    }

    <#
     The footer links are joined by a middot separator. The browser variant
     spaces it with `padding:0 10px` on the <span>; Outlook's Word engine
     IGNORES padding on an inline <span>, which glued the two links together
     ("report &rarr;Open matrix..."). The MSO variant therefore spaces the
     middot with non-breaking spaces (`&nbsp;`), which Word does honour.
    #>
    $browserSep = "<span style='color:$($Script:Theme.TextLight); padding:0 10px;'>&middot;</span>"
    $msoSep = "<span style='color:$($Script:Theme.TextLight);'>&nbsp;&nbsp;&middot;&nbsp;&nbsp;</span>"
    $footerLinksHtmlBrowser = if ($footerLinks.Count -gt 0) { $footerLinks -join $browserSep } else { '&nbsp;' }
    $footerLinksHtmlMso = if ($footerLinks.Count -gt 0) { $footerLinks -join $msoSep } else { '&nbsp;' }

    <#
     Outlook's Word engine applies the stylesheet default
     `p { margin: 0 0 12px 0 }` to the implicit paragraph inside a <td>,
     which adds ~12px of space BELOW the footer links (more than above).
     Wrapping the links in a `<p style='margin:0'>` neutralises that so the
     td's symmetric top/bottom padding yields equal white space on both
     sides. The browser branch is untouched (still renders perfectly).
    #>
    $footerRowHtml = @"
<!--[if mso]>
    <tr>
        <td valign='middle' style='padding:6px 16px 6px 16px; text-align:center; font-size:12px; line-height:16px; mso-line-height-rule:exactly; color:$($Script:Theme.TextLight);'>
            <p style='margin:0; mso-line-height-rule:exactly; line-height:16px;'>$footerLinksHtmlMso</p>
        </td>
    </tr>
<![endif]-->
<!--[if !mso]><!-->
    <tr>
        <td valign='top' style='padding:4px 16px 12px 16px; text-align:center; font-size:12px; line-height:16px; color:$($Script:Theme.TextLight);'>
            $footerLinksHtmlBrowser
        </td>
    </tr>
<!--<![endif]-->
"@

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
    $headerLabelHtml = "<span style=`"font-size:12px; font-weight:700; color:#e5e7eb; text-transform:uppercase; letter-spacing:0.5px;`">$headerLabel</span>"

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

    # A file-level FatalError (e.g. 'Runspace processing failed') means the
    # settings never executed. Flag it so their rows render as "Skipped"
    # instead of green "success".
    $fileHasError = $false
    foreach ($g in $fileLevelGroups) {
        if ($g.Checks -and @($g.Checks | Where-Object Type -EQ 'FatalError').Count -gt 0) {
            $fileHasError = $true
            break
        }
    }

    if ($fileLevelCount -gt 0) {
        $contentRows += @"
<tr>
    <td style='padding:14px 16px 6px 16px; font-size:11px; font-weight:700; color:$($Script:Theme.TextLight); letter-spacing:1.5px; text-transform:uppercase;'><p style='margin:0; mso-line-height-rule:exactly; line-height:14px;'>File Issues ($fileLevelCount)</p></td>
</tr>
"@
        foreach ($g in $fileLevelGroups) {
            if ($g.Checks) {
                foreach ($c in $g.Checks) {
                    $contentRows += Build-FileLevelCheckRowHC -Check $c -SheetLabel $g.Label -ShowLabel $false
                }
            }
        }
    }

    # Settings rows — each row is a self-contained fluid flex card
    if ($FileContext.Matrices -and $FileContext.Matrices.Count -gt 0) {
        $sortedMatrices = $FileContext.Matrices |
        Sort-Object { $_.Setting.Formatted.ComputerName }, { $_.Setting.Formatted.Path }, { $_.ID }

        $settingsRowsHtml = ''
        $settingsIndex = 0
        foreach ($m in $sortedMatrices) {
            if ($settingsIndex -gt 0) {
                $settingsRowsHtml += '<!--[if mso]><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#ffffff" style="background-color:#ffffff;"><tr><td bgcolor="#ffffff" height="4" style="font-size:0; line-height:4px; mso-line-height-rule:exactly; background-color:#ffffff;">&#160;</td></tr></table><![endif]-->'
            }
            $settingsRowsHtml += Build-SettingsRowHC -MatrixItem $m -FileHasError $fileHasError
            $settingsIndex++
        }

        $matrixCount = @($sortedMatrices).Count
        $contentRows += @"
<tr>
    <td style='padding:14px 16px 6px 16px; font-size:11px; font-weight:700; color:$($Script:Theme.TextLight); letter-spacing:1.5px; text-transform:uppercase;'><p style='margin:0; mso-line-height-rule:exactly; line-height:14px;'>Settings ($matrixCount)</p></td>
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
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="$($Script:Theme.BgWhite)" style="border-collapse:separate; margin:0 0 16px 0; table-layout:fixed; width:100%; max-width:100%; background-color:$($Script:Theme.BgWhite); border:1px solid $($Script:Theme.BorderLight); border-radius:10px; overflow:hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.06);">
    <tr>
        <td style='padding:0; background-color:$gradTo; background-image: linear-gradient(135deg, $gradFrom 0%, $gradTo 100%); border-bottom:1px solid $($Script:Theme.BorderLight);'>
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;">
                <tr>
                    <td valign='middle' width='34' style='padding:14px 0 14px 18px; font-size:20px; font-weight:bold; color:#ffffff; line-height:1; text-align:left;'>$headerSymbol</td>
                    <td valign='middle' style='padding:14px 8px 14px 4px;'>
                        <div style='font-size:16px; font-weight:700; color:#ffffff; line-height:1.25;'>
                            <a href="$matrixLink" title="$matrixTitle" style="color:#ffffff; text-decoration:none;">$fileName</a>
                        </div>
                        <div style='font-size:12px; color:#f1f2f4; line-height:1.4; margin-top:2px;font-style:italic;'>
                            $lastChangeInfo
                        </div>
                    </td>
                    <td valign='middle' align='right' width='112' style='padding:14px 12px 14px 6px; white-space:nowrap; width:112px;'>$headerLabelHtml</td>
                </tr>
            </table>
        </td>
    </tr>
    $contentRows
    $footerRowHtml
</table>
<!--[if mso]>
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td height="16" style="font-size:0; line-height:16px; mso-line-height-rule:exactly;">&#160;</td></tr></table>
<![endif]-->
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
        [string]$BrowserViewFilePath,
        [datetime]$ScriptStartTime,
        [datetime]$ScriptEndTime = (Get-Date),
        $LogFolder
    )

    $scriptName = [System.Net.WebUtility]::HtmlEncode(
        (Get-StringOrDefaultHC $Settings.ScriptName 'Permission Matrix')
    )
    $userBody = Get-StringOrDefaultHC $Settings.SendMail.Body ''
    $bodyWidth = $Script:Theme.BodyWidth
    $bgPage = $Script:Theme.BgPage

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
    $topLinksBlock = Build-MailTopLinksBlockHC `
        -BrowserViewFilePath $BrowserViewFilePath `
        -ExportedFiles $ExportedFiles

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

        $startEnc = [System.Net.WebUtility]::HtmlEncode($startStr)
        $endEnc = [System.Net.WebUtility]::HtmlEncode($endStr)
        $durEnc = [System.Net.WebUtility]::HtmlEncode($durStr)

        # Word (Outlook) ignores margin/display:inline-block on inline spans, so
        # the three label/value pairs ran together as
        # "STARTED..date..ENDED..date..". Render them as a centered table
        # instead — cell padding provides the gaps and the label/value spacing
        # consistently in Outlook and browsers alike.
        $footLabelStyle = "font-size:10px; font-weight:700; color:$($Script:Theme.TextLight); text-transform:uppercase; letter-spacing:0.5px;"
        $footValueStyle = "font-size:11px; color:$($Script:Theme.TextLight); font-family:$($Script:Theme.MonoStack);"

        $footer = @"
<table role="presentation" align="center" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse; margin:16px auto 0 auto;">
    <tr>
        <td style="padding:0 5px 0 0; $footLabelStyle">Started</td>
        <td style="padding:0 20px 0 0; $footValueStyle">$startEnc</td>
        <td style="padding:0 5px 0 0; $footLabelStyle">Ended</td>
        <td style="padding:0 20px 0 0; $footValueStyle">$endEnc</td>
        <td style="padding:0 5px 0 0; $footLabelStyle">Duration</td>
        <td style="padding:0; $footValueStyle">$durEnc</td>
    </tr>
</table>
"@
    }

    @"
<!DOCTYPE html>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<!--[if mso]>
<style type="text/css">
    v\:* { behavior: url(#default#VML); display:inline-block; }
    table { mso-table-lspace:0pt; mso-table-rspace:0pt; }
    td { mso-line-height-rule:exactly; }
</style>
<![endif]-->
$($Html.Style)
</head>
<body style="margin:0; padding:0;">
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="$bgPage" style="border-collapse:collapse; background-color:$bgPage;">
    <tr>
        <td align="center" valign="top" bgcolor="$bgPage" style="padding:20px; background-color:$bgPage;">
            <!--[if mso]>
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="$bodyWidth" align="center"><tr><td>
            <![endif]-->
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; width:100%; margin:0 auto;">
                <tr><td style="padding:0 0 4px 0;"><h1>$scriptName</h1></td></tr>
                <tr><td style="padding:0 0 16px 0; color:$($Script:Theme.TextMuted); font-size:13px; line-height:1.6;">$userBody</td></tr>
                <tr><td style="padding:0;">$topLinksBlock</td></tr>
                <tr><td style="padding:0;">$($Html.ErrorWarningTable)</td></tr>
                <tr><td style="padding:0;">$systemErrorsBlock</td></tr>
                <tr><td style="padding:0;">$($Html.MatrixTables)</td></tr>
                <tr><td style="padding:0;">$footer</td></tr>
            </table>
            <!--[if mso]>
            </td></tr></table>
            <![endif]-->
        </td>
    </tr>
</table>
</body>
</html>
"@
}