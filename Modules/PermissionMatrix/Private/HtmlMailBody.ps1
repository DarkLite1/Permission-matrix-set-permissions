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

            Outlook (Word engine) spacing fixes:
             * The section header text is wrapped in <p style='margin:0'>.
               Word wraps bare td text in an implicit paragraph with a
               default 12px bottom margin, which added ~12px of phantom
               space between the header and the first card. The wrapper
               kills that; the td's own 8px bottom padding remains the
               single source of the gap (matching the browser).
             * Word ignores 'margin' on <table>, so the wrapping table's
               'margin:0 0 20px 0' — the space below the section — never
               rendered in Outlook. An MSO-only 20px spacer table after the
               section restores it; browsers skip the conditional and keep
               using the CSS margin.
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
            $catText = $category.ToUpper()
            $pillW = 18 + (7 * $catText.Length)

            $catHtml = @"
<!--[if mso]><v:roundrect xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="urn:schemas-microsoft-com:office:word" style="height:16px;width:${pillW}px;v-text-anchor:middle;display:inline-block;" arcsize="50%" fillcolor="$($Script:Theme.BgAlt)" strokecolor="$($Script:Theme.BorderLight)" strokeweight="1px"><w:anchorlock/><center style="color:$($Script:Theme.TextMuted); font-family:Arial,sans-serif; font-size:10px; font-weight:600; letter-spacing:0.5px;">$catText</center></v:roundrect>&nbsp;&nbsp;<![endif]--><!--[if !mso]><!--><span style='display:inline-block; margin-right:8px; padding:1px 8px; background-color:$($Script:Theme.BgAlt); border:1px solid $($Script:Theme.BorderLight); border-radius:10px; font-size:10px; font-weight:600; color:$($Script:Theme.TextMuted); letter-spacing:0.5px;'>$catText</span><!--<![endif]-->
"@
        }

        $pill = New-PillHtmlHC -Text $pillText -Bg $pillBg

        $rows += @"
<tr>
    <td style='padding:0 0 8px 0;'>
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" class="rr-syscard" bgcolor="$bgColor" style="border-collapse:separate; width:100%; max-width:100%; background-color:$bgColor; border-left:3px solid $accentColor; border-radius:6px;">
            <tr>
                <td valign="middle" width="26" style='padding:12px 0 12px 14px; color:$accentColor; font-size:16px; font-weight:bold; line-height:1;'>$glyph</td>
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
        <td style='padding:0 0 8px 0; font-size:11px; font-weight:700; color:$($Script:Theme.TextLight); letter-spacing:1.5px; text-transform:uppercase;'><p style='margin:0; mso-line-height-rule:exactly; line-height:14px;'>$headerLabel</p></td>
    </tr>
    $rows
</table>
<!--[if mso]>
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td height="12" style="font-size:0; line-height:12px; mso-line-height-rule:exactly;">&#160;</td></tr></table>
<![endif]-->
"@
}

function Build-MailTopLinksBlockHC {
    <#
        .DESCRIPTION
            Renders the "view in browser" line and the "Export files: A · B · C"
            link line at the top of the mail.

            Outlook (Word engine) fixes, mirroring Build-MatrixFileCardHC:
             * The export-link middot separator spaced itself with
               'padding:0 8px' on an inline <span>, which Word ignores —
               the links rendered glued together. The MSO variant spaces
               the middot with '&nbsp;' instead; the browser variant keeps
               the original padding, so browser rendering is unchanged.
               Each variant lives in its own conditional row so every
               client renders exactly one.
             * Both td texts are wrapped in <p style='margin:0'> to kill
               Word's implicit-paragraph 12px bottom margin.
             * Word ignores 'margin' on <table>, so the wrapping table's
               'margin:0 0 14px 0' (the gap above "Detected issues") never
               rendered in Outlook. An MSO-only 14px spacer table after the
               block restores it; browsers skip it and keep the CSS margin.
    #>
    param(
        [string]$BrowserViewFilePath,
        $ExportedFiles,
        # Run-level diagnostics page. Appended to the same link line as the
        # export files rather than given a block of its own: it is one more
        # run artifact, and the line is already Outlook-safe.
        [string]$DiagnosticsHtmlPath
    )

    $linkStyle = "color:$($Script:Theme.LinkColor); text-decoration:none; font-weight:600;"
    $mutedStyle = "color:$($Script:Theme.TextMuted); font-size:12px; line-height:1.45;"
    $rows = ''

    if (-not [string]::IsNullOrWhiteSpace($BrowserViewFilePath)) {
        $browserUrl = [System.Net.WebUtility]::HtmlEncode((ConvertTo-FileUrlHC $BrowserViewFilePath))
        $browserTitle = [System.Net.WebUtility]::HtmlEncode($BrowserViewFilePath)

        # MSO-only: the "view it in the browser" escape hatch only makes
        # sense inside a mail client with limited rendering (classic
        # Outlook). When the saved HTML file is opened in a browser, the
        # reader IS already in the browser, so the line is hidden there.
        # NOTE: this also hides the line in non-Word mail clients (Outlook
        # on the web, Gmail, mobile) — acceptable, since those render the
        # mail correctly anyway and share the same HTML as the saved file.
        $rows += @"
<!--[if mso]>
<tr>
    <td style='padding:0 0 8px 0; $mutedStyle'><p style='margin:0; mso-line-height-rule:exactly; line-height:17px;'>If this mail is not visible, please <a href='$browserUrl' title="$browserTitle" target='_blank' rel='noopener noreferrer' style='$linkStyle'>click here to view it in the browser</a>.</p></td>
</tr>
<![endif]-->
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

    if (-not [string]::IsNullOrWhiteSpace($DiagnosticsHtmlPath)) {
        $diagUrl = [System.Net.WebUtility]::HtmlEncode((ConvertTo-FileUrlHC $DiagnosticsHtmlPath))
        $diagTitle = [System.Net.WebUtility]::HtmlEncode($DiagnosticsHtmlPath)
        $exportLinks.Add("<a href='$diagUrl' title=`"$diagTitle`" target='_blank' rel='noopener noreferrer' style='$linkStyle'>Diagnostics</a>")
    }

    # 'Export files' stops being accurate once the diagnostics page is in the
    # list: that page is a run artifact, not one of the configured exports.
    # The label only changes when the link is actually present, so mails for a
    # run without diagnostics read exactly as before.
    $linkLineLabel = if (-not [string]::IsNullOrWhiteSpace($DiagnosticsHtmlPath)) {
        'Files'
    }
    else { 'Export files' }

    if ($exportLinks.Count -gt 0) {
        # Word ignores padding on inline spans, so the browser separator
        # (padding:0 8px) collapses in Outlook and the links run together.
        # MSO variant spaces the middot with non-breaking spaces instead.
        $browserSep = "<span style='color:$($Script:Theme.TextLight); padding:0 8px;'>&middot;</span>"
        $msoSep = "<span style='color:$($Script:Theme.TextLight);'>&nbsp;&nbsp;&middot;&nbsp;&nbsp;</span>"
        $linksHtmlBrowser = $exportLinks -join $browserSep
        $linksHtmlMso = $exportLinks -join $msoSep

        $rows += @"
<!--[if mso]>
<tr>
    <td style='padding:0; $mutedStyle'><p style='margin:0; mso-line-height-rule:exactly; line-height:17px;'>${linkLineLabel}: $linksHtmlMso</p></td>
</tr>
<![endif]-->
<!--[if !mso]><!-->
<tr>
    <td style='padding:0; $mutedStyle'>${linkLineLabel}: $linksHtmlBrowser</td>
</tr>
<!--<![endif]-->
"@
    }

    if (-not $rows) { return '' }

    return @"
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; margin:0 0 14px 0; table-layout:fixed; width:100%; max-width:100%;">
    $rows
</table>
<!--[if mso]>
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td height="14" style="font-size:0; line-height:14px; mso-line-height-rule:exactly;">&#160;</td></tr></table>
<![endif]-->
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

    # Notices that are neither errors nor warnings — Type 'Information', and any
    # unknown/future type, matching how Build-FileLevelCheckRowHC decides what
    # is "info". These live on the MATRIX (this row), not on the matrix FILE, so
    # they never reach the file-level tallies that colour the card header, and
    # they never earn a status pill. Before this, a matrix carrying only an
    # info notice (e.g. 'AD groups without members') looked completely clean in
    # the overview and the only trace of it was in the execution report.
    $infoCount = @($MatrixItem.Check | Where-Object {
            $_.Type -ne 'FatalError' -and $_.Type -ne 'Warning'
        }).Count

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

    # Small blue "i" shown next to the computer name whenever this matrix has
    # info-level notices — INDEPENDENT of the status pill, so it appears on
    # green, amber and red rows alike (a row can carry both a Warning and an
    # Information check, as in BNL-MTX-STAFF-HR).
    #
    # It deliberately does NOT go in the pill cell: that cell is a fixed 84px
    # and New-PillHtmlHC sizes its Outlook VML shape from the text length
    # ('Warning' already computes to ~85px), so a second pill there would
    # reopen the column-alignment problems. Sitting inline after the name — the
    # same trick HtmlReport.ps1 uses for its grey "Skipped" tag — costs no
    # layout at all.
    #
    # Metrics are matched to the surrounding name line on purpose (font-size
    # 13px, line-height 15px, mso-line-height-rule:exactly). Word grows or
    # clips a line box around an inline run with a LARGER font-size, and the
    # identifier cell is the cell that drives this row's height — changing it
    # would shift the vertical centring of the pill and the meta columns. A
    # same-size glyph leaves the line box untouched.
    #
    # &#8505; (U+2139) is the same glyph Build-FileLevelCheckRowHC already
    # renders for info cards, so it is proven to show up in Outlook Classic.
    # The colour is the theme's AccentInfo blue rather than the grey that
    # Get-CheckThemeHC gives info cards: at this size, in a dense list, grey on
    # a bold dark name is easy to miss.
    $infoTag = ''
    if ($infoCount -gt 0) {
        $infoTitle = "$infoCount information notice" +
        $(if ($infoCount -ne 1) { 's' }) +
        ' on this matrix - open the execution report for details'
        $infoTag = "&nbsp;<span title=`"$infoTitle`" style='color:$($Script:Theme.AccentInfo); font-size:13px; font-weight:400; line-height:15px; mso-line-height-rule:exactly;'>&#8505;</span>"
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
    # The pill cell is a plain valign='middle' td holding the standard combined
    # (MSO VML + browser span) pill from New-PillHtmlHC — the same pattern as
    # the file-level check rows, where it demonstrably centres correctly in
    # both Outlook and browsers. Two earlier constructions failed in Outlook:
    #  1. An MSO-only nested spacer table (8px/4px asymmetric) around the pill
    #     drove the whole row's height and skewed the vertical alignment of
    #     the pill AND the metadata cell.
    #  2. Even with that removed, the pill stayed off-centre because the
    #     IDENTIFIER cell stacked its two lines with a nested <table> — and
    #     a nested table in one cell breaks Word's valign='middle' for the
    #     SIBLING cells in the same row. The identifier now stacks its lines
    #     with <div>s (margin:0 + exact line-height), exactly like the
    #     file-level check rows where centring is proven to work.
    #  3. The cell must NOT carry a line-height: the global '<head>' MSO style
    #     applies 'mso-line-height-rule:exactly' to every td, and an explicit
    #     line-height:16px clamped Word's line box below the 26px VML pill,
    #     which then overflowed UPWARD from its baseline and sat high. The
    #     check-row pill cell has no line-height, which is why it centres.
    #  4. The cell DOES need 'font-size:0': Word reserves descender space
    #     below an inline VML shape proportional to the cell's font size.
    #     Without it the pill cell grew a few px taller than the identifier
    #     cell, became the row's tallest cell, and the extra height appeared
    #     as empty space UNDER the identifier text (which then looked
    #     top-aligned). Zero font size = zero descent = the pill's line box
    #     is exactly the 26px shape and the identifier cell drives the row
    #     height again. The browser pill span sets its own font-size, so
    #     browsers are unaffected.
    if ($pillText) {
        $pillHtml = New-PillHtmlHC -Text $pillText -Bg $pillBg
        $pillTd = "<td valign='middle' align='right' class='rr-srow-status' width='84' style='vertical-align:middle; padding:4px 12px 4px 4px; white-space:nowrap; font-size:0;'>$pillHtml</td>"
    }
    else {
        # Clean row: reserve the column width with a single empty cell.
        $pillTd = "<td valign='middle' align='right' width='84' style='vertical-align:middle; padding:4px 12px 4px 4px; white-space:nowrap;'>&nbsp;</td>"
    }

    # Action and Duration are two SEPARATE fixed-width cells, not two spans in
    # one right-aligned cell. Previously both lived in a single align='right'
    # td: the pair was flushed right as a block, so a short duration ('N/A',
    # 3 chars) pulled the Action label ~30px to the right compared with rows
    # showing a full '00:00:15' timestamp, and the Action column visibly
    # jittered from row to row.
    #
    # Now each has its own column, so neither can be moved by the other:
    #   - Action   : right-aligned in a 44px cell, so its right edge is fixed.
    #   - Duration : CENTRED in a 68px cell, so 'N/A' sits in the middle of the
    #                same column the timestamps occupy above and below it.
    # 44 + 68 = 112px replaces the old single 104px column; the 8px comes out
    # of the identifier cell, which wraps its path anyway.
    #
    # Both the presentational align attribute AND text-align are set: Outlook
    # Classic's Word engine honours the attribute, browsers honour either. The
    # duration cell also carries the nowrap ATTRIBUTE (Word ignores
    # white-space:nowrap in CSS) so Word can never break '00:00:15' across two
    # lines in the narrower cell — it shrinks the identifier column instead. In
    # browsers the <520px media query still wins over the attribute, because a
    # presentational hint loses to author CSS (and it carries !important).
    #
    # The mso-only '&nbsp;' spacer that used to separate the two spans is gone:
    # the gap is now real cell padding (0/10px inner edges), which Word renders
    # natively, so the spacing is identical in Outlook and the browser.
    #
    # Vertical metrics (padding:4px, font-size:11px, line-height:15px +
    # mso-line-height-rule:exactly) are unchanged and identical on both cells,
    # so the row height — and therefore the hard-won pill centring documented
    # above — is untouched.
    return @"
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" class="rr-srow" style="border-collapse:separate; width:100%; max-width:100%; margin:0 0 4px 0; table-layout:fixed; background-color:$($Script:Theme.BgWhite); border:1px solid $($Script:Theme.BorderMain); border-left:3px solid $accent; border-radius:6px;">
    <tr>
        <td valign='middle' width='20' style='vertical-align:middle; padding:4px 0 4px 12px; color:$accent; font-size:12px; line-height:15px; mso-line-height-rule:exactly;'>&#9679;</td>
        <td valign='middle' class='rr-srow-ident' style='vertical-align:middle; padding:4px 8px;'>
            <div style='margin:0; font-weight:700; color:$($Script:Theme.TextMain); font-size:13px; line-height:15px; mso-line-height-rule:exactly;'><a href='$link' target='_blank' rel='noopener noreferrer' style='text-decoration:none; color:$($Script:Theme.TextMain);'>$comp</a>$infoTag</div>
            <div style='margin:0; font-family:$($Script:Theme.MonoStack); font-size:11px; color:$($Script:Theme.TextMuted); line-height:14px; mso-line-height-rule:exactly; white-space:normal; overflow-wrap:anywhere; word-break:break-all;'$pathTitle><a href='$link' target='_blank' rel='noopener noreferrer' style='text-decoration:none; color:$($Script:Theme.TextMuted);'>$pathDisp</a></div>
        </td>
        <td valign='middle' align='right' class='rr-srow-meta' width='44' style='vertical-align:middle; padding:4px 0 4px 10px; color:$($Script:Theme.TextLight); font-size:11px; line-height:15px; mso-line-height-rule:exactly; white-space:nowrap; text-align:right;'>$action</td>
        <td valign='middle' align='center' nowrap='nowrap' class='rr-srow-meta rr-srow-dur' width='68' style='vertical-align:middle; padding:4px 10px 4px 8px; color:$($Script:Theme.TextLight); font-family:$($Script:Theme.MonoStack); font-size:11px; line-height:15px; mso-line-height-rule:exactly; white-space:nowrap; text-align:center;'>$dur</td>
        $pillTd
    </tr>
</table>
"@
}

function Build-MatrixFileCardHC {
    param([object]$FileContext)

    # File header info — resolved via the same helper the overview sorts on,
    # so the visible title and the card's position always agree.
    $fileName = [System.Net.WebUtility]::HtmlEncode(
        (Get-MatrixFileNameHC -FileResult $FileContext)
    )

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

    # Tally checks across all sources to decide header color and summary text.
    # Shared with the overview's sort so a card's colour and its position
    # always tell the same story.
    $tally = Get-FileCheckTallyHC -FileResult $FileContext
    $fileErrs = $tally.Errors
    $fileWarns = $tally.Warnings

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

    <#
     Outlook's Word engine cannot render CSS gradients, so it falls back to
     the header's flat 'background-color'. Using $gradTo (the brightest end)
     made the Outlook header noticeably lighter than the browser's gradient.
     Instead, compute the per-channel midpoint of the two gradient stops and
     use THAT as the fallback: Outlook gets the gradient's average tone,
     while browsers paint the gradient over the fallback, so they are
     unaffected. Falls back to $gradTo if the hex parse ever fails.
    #>
    $gradMid = $gradTo
    try {
        $f = $gradFrom.TrimStart('#')
        $t = $gradTo.TrimStart('#')
        if ($f.Length -eq 6 -and $t.Length -eq 6) {
            $gradMid = '#' + ((0, 2, 4 | ForEach-Object {
                        '{0:x2}' -f [int][Math]::Round(
                            ([Convert]::ToInt32($f.Substring($_, 2), 16) +
                            [Convert]::ToInt32($t.Substring($_, 2), 16)) / 2
                        )
                    }) -join '')
        }
    }
    catch { $gradMid = $gradTo }

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
        <td bgcolor="$gradMid" style='padding:0; background-color:$gradMid; background-image: linear-gradient(135deg, $gradFrom 0%, $gradTo 100%); border-bottom:1px solid $($Script:Theme.BorderLight);'>
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;">
                <tr>
                    <!--
                     Glyph cell: Word treats the title/subtitle blocks as
                     paragraphs and (without explicit margins) pads them with
                     its default 12px bottom margin, which threw off the row
                     height and made valign='middle' place the glyph off-centre
                     in Outlook. The <p margin:0> wrappers below fix the root
                     cause; the MSO glyph cell additionally gets an explicit
                     line-height (33px, tuned by eye — Word places the glyph on
                     the line box's baseline, so the line-height is the knob
                     that moves it up/down; smaller lifts it) so the glyph is
                     vertically centred regardless of Word's baseline handling. Horizontally, Word mis-renders the
                     browser's asymmetric 18px-left-padding layout, so the MSO
                     cell instead centres the glyph in a fixed 52px column —
                     the same footprint as the browser's 18px padding + 34px
                     cell — putting the glyph centre at ~26px in both clients.
                     Browsers keep the original cell, untouched.
                    -->
                    <!--[if mso]><td valign='middle' align='center' width='52' style='vertical-align:middle; text-align:center; padding:14px 0; font-size:20px; font-weight:bold; color:#ffffff; line-height:33px; mso-line-height-rule:exactly;'>$headerSymbol</td><![endif]-->
                    <!--[if !mso]><!--><td valign='middle' width='34' style='padding:14px 0 14px 18px; font-size:20px; font-weight:bold; color:#ffffff; line-height:1; text-align:left;'>$headerSymbol</td><!--<![endif]-->
                    <td valign='middle' style='padding:14px 8px 14px 4px;'>
                        <p style='margin:0; font-size:16px; font-weight:700; color:#ffffff; line-height:20px; mso-line-height-rule:exactly;'>
                            <a href="$matrixLink" title="$matrixTitle" style="color:#ffffff; text-decoration:none;">$fileName</a>
                        </p>
                        <p style='margin:2px 0 0 0; font-size:12px; color:#f1f2f4; line-height:17px; mso-line-height-rule:exactly; font-style:italic;'>
                            $lastChangeInfo
                        </p>
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

    <#
     Order the cards so the ones needing attention are reachable without
     scrolling: files with an ERROR first, then files with a WARNING, then
     everything else — each of those three groups sorted alphabetically by
     matrix file name (the same string the gradient header shows and links).

     Rank 0/1/2 is computed from Get-FileCheckTallyHC, the exact tally that
     colours the header, so the run of red cards, then amber, then green
     matches the sequence a reader sees. Informational notices deliberately
     do NOT promote a card: they are notices, not issues, so an info-only
     matrix stays in the alphabetical run with the other clean files.

     Sort-Object is stable and the name is the second key, so ties inside a
     rank resolve alphabetically and the whole output is deterministic.
     Without this, $Context.FileResults arrives in runspace COMPLETION order
     (the files are processed by Invoke-WithOptionalParallelismHC), so the
     cards appeared in a different, effectively random order on every run.

     Sorting here rather than at the source keeps the change scoped to
     presentation: the export sheets, the JSON detail files and the audit
     report keep whatever order they already had. Both the mailed body and
     the browser copy in the log folder are built from this one string, so a
     single sort fixes both.
    #>
    $sortedFileResults = $FileResults | Sort-Object -Property @{
        Expression = {
            $tally = Get-FileCheckTallyHC -FileResult $_

            if ($tally.Errors -gt 0) { 0 }
            elseif ($tally.Warnings -gt 0) { 1 }
            else { 2 }
        }
    }, @{
        Expression = { Get-MatrixFileNameHC -FileResult $_ }
    }

    $output = ''
    foreach ($fileContext in $sortedFileResults) {
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
    # Diagnostics page for this run. Derived from $LogFolder rather than
    # threaded through as a new parameter, and only linked when the file
    # actually exists, so a run that wrote no diagnostics shows no dead link.
    $diagnosticsHtmlPath = ''
    if ($LogFolder) {
        $logFolderPath = if ($LogFolder -is [string]) { $LogFolder } else { $LogFolder.FullName }

        if (-not [string]::IsNullOrWhiteSpace($logFolderPath)) {
            $candidate = Join-Path -Path $logFolderPath -ChildPath 'Diagnostics.html'
            if (Test-Path -LiteralPath $candidate) { $diagnosticsHtmlPath = $candidate }
        }
    }

    $topLinksBlock = Build-MailTopLinksBlockHC `
        -BrowserViewFilePath $BrowserViewFilePath `
        -ExportedFiles $ExportedFiles `
        -DiagnosticsHtmlPath $diagnosticsHtmlPath

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