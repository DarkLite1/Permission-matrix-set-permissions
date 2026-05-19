<#
    Html.ps1
    Consolidated HTML rendering logic for Toolbox.PermissionMatrixHC

    --- MODERN DASHBOARD REDESIGN ---
    The email body is a SUMMARY ONLY: per-file gradient header cards with the
    file's settings rows. Errors/warnings are linked but listed in the
    standalone "00 - Execution Report.html" produced by Write-MatrixExecutionReportHC.
    System errors (script-level exceptions) are surfaced separately at the top
    of the email since they don't belong to any single file.

    Layout uses table-based HTML with inline styles so modern Outlook (Windows,
    Mac, Web) and standalone browsers render the same picture. No flexbox, no
    CSS grid. Width is fixed at 620px so the email fits comfortably on small
    laptop screens beside the inbox sidebar.

    Color palette: amber for warnings, red for errors, green for success.
    Gradients in card headers use a 135deg dark-to-light variant of the accent
    color with a solid bgcolor fallback for any client that ignores gradients.
#>

# =====================================================================
# GLOBAL HTML THEME
# Centralized color palette and typography used by all HTML generation.
# Edit values here to retune the whole look.
# =====================================================================
$Script:Theme = @{
    # Status backgrounds (soft tints)
    StatusError    = '#fee2e2'
    StatusWarning  = '#fef3c7' # Amber tint
    StatusSuccess  = '#dcfce7'
    StatusSkipped  = '#f3f4f6'

    # Accent colors (used for icons, pills, left borders, status dots)
    AccentError    = '#dc2626'
    AccentWarning  = '#d97706' # Amber
    AccentSuccess  = '#16a34a'
    AccentSkipped  = '#6b7280'
    AccentInfo     = '#2563eb'
    AccentSystem   = '#7c2d12' # Maroon for system errors

    # Gradient stops for card headers (dark, mid)
    GradError      = @('#7f1d1d', '#dc2626')
    GradWarning    = @('#78350f', '#d97706')
    GradSuccess    = @('#14532d', '#16a34a')

    # Text colors
    TextMain       = '#111827'
    TextMuted      = '#374151'
    TextLight      = '#6b7280'

    # Page and surface colors
    BgPage         = '#e5e7eb' # Page background — slightly darker so cards pop
    BgWhite        = '#ffffff'
    BgAlt          = '#f9fafb' # Off-white for muted backgrounds

    # Borders
    BorderMain     = '#d1d5db'
    BorderLight    = '#e5e7eb'

    # Links
    LinkColor      = '#2563eb'
    LinkHoverColor = '#1d4ed8'

    # Typography stacks
    FontStack      = "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif"
    MonoStack      = "'SF Mono', Consolas, 'Liberation Mono', Menlo, monospace"

    # Body width — narrow enough to coexist with inbox sidebars on small laptops
    BodyWidth      = 620
}

function Initialize-HtmlStructureHC {

    $style = @"
<style type="text/css">
    body {
        font-family: $($Script:Theme.FontStack);
        font-size: 13px;
        color: $($Script:Theme.TextMain);
        background-color: $($Script:Theme.BgPage);
        margin: 0;
        padding: 20px;
        -webkit-font-smoothing: antialiased;
    }
    a { color: $($Script:Theme.LinkColor); text-decoration: none; }
    a:hover { color: $($Script:Theme.LinkHoverColor); text-decoration: underline; }
    h1 {
        font-size: 22px;
        font-weight: 700;
        color: $($Script:Theme.TextMain);
        margin: 0 0 4px 0;
        letter-spacing: -0.3px;
    }
    h2, h3 { margin-bottom: 0; }
    p { margin: 0 0 12px 0; }
    p.italic { font-style: italic; font-size: 12px; }
    hr {
        border: none;
        border-top: 1px solid $($Script:Theme.BorderLight);
        margin: 16px 0 20px 0;
    }
    table { border-collapse: collapse; }
    /* Legacy classes preserved for any external consumer; the new
       email layout uses inline styles exclusively. */
    .probTypeError { background-color: $($Script:Theme.StatusError); }
    .probTypeWarning { background-color: $($Script:Theme.StatusWarning); }
    .probTypeInfo { background-color: $($Script:Theme.BgAlt); }
</style>
"@

    $troubleshooting = @'
<style type="text/css">
    body { margin: 20px; }
</style>
'@

    @{
        Style                = $style
        TroubleshootingStyle = $troubleshooting
        Templates            = @{
            # Kept as no-op placeholders — the new layout doesn't use them,
            # but external callers might reference them.
            SettingsHeader = ''
            LegendTable    = ''
        }
    }
}

function Get-HtmlClassProbTypeHC {
    param([string]$Type)
    switch ($Type) {
        'FatalError' { 'probTypeError' }
        'Warning' { 'probTypeWarning' }
        default { 'probTypeInfo' }
    }
}

function Format-IssueCountLabelHC {
    param([int]$Errors, [int]$Warnings)
    $parts = @()
    if ($Errors -gt 0) {
        $parts += "$Errors Error" + $(if ($Errors -ne 1) { 's' })
    }
    if ($Warnings -gt 0) {
        $parts += "$Warnings Warning" + $(if ($Warnings -ne 1) { 's' })
    }
    if ($parts.Count -eq 0) { return 'Success' }
    return ($parts -join ', ')
}

function Format-LastChangeHC {
    <#
        .DESCRIPTION
            Build a "Last change: ..." line from an Excel file's last-modified
            metadata. Handles missing/unknown values gracefully:
                - Both user and date known  → "Last change: Brecht · 19/05/2026 13:30"
                - Only user known  → "Last change: Brecht"
                - Only date known  → "Last change: 19/05/2026 13:30"
                - Neither known    → "No modification metadata available"

            The user component is HTML-encoded. The separator is the HTML
            entity &middot;. Callers can drop the returned string directly
            into HTML; if empty, they should skip rendering the line.
        .PARAMETER LastModifiedBy
            Raw username string from ExcelInfo.LastModifiedBy. Treated as
            missing when null, empty, whitespace, or the literal 'Unknown'.
        .PARAMETER Modified
            Raw datetime from ExcelInfo.Modified. Treated as missing when
            not a [datetime] or equal to [datetime]::MinValue.
    #>
    param(
        [object]$LastModifiedBy,
        [object]$Modified
    )

    $rawBy = Get-StringOrDefaultHC $LastModifiedBy ''
    $hasBy = $rawBy -and $rawBy -ne 'Unknown'

    $hasDt = ($Modified -is [datetime]) -and ($Modified -gt [datetime]::MinValue)
    $dtStr = if ($hasDt) { $Modified.ToString('dd/MM/yyyy HH:mm') } else { '' }

    $modBy = [System.Net.WebUtility]::HtmlEncode($rawBy)

    if ($hasBy -and $hasDt) { return "Last change: $modBy &middot; $dtStr" }
    if ($hasBy) { return "Last change: $modBy" }
    if ($hasDt) { return "Last change: $dtStr" }
    return 'No modification metadata available'
}

function ConvertTo-FileUrlHC {
    <# 
    .DESCRIPTION
        Convert a Windows path (UNC or local) to a `file://` URL suitable for
        `href` attributes. Normalizes backslashes to forward slashes and
        percent-encodes spaces. Returns empty string for null/empty input.
    #>
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    return 'file://' + ($Path -replace '\\', '/' -replace ' ', '%20')
}

function Get-TruncatedPathHC {
    <# 
        .DESCRIPTION
            Truncate a file path in the middle if it exceeds a certain length,
            replacing the removed portion with an ellipsis. Attempts to break
            on backslash boundaries for cleaner output, but falls back to
            character-based truncation if necessary. Returns an array of the
            truncated display string and a boolean indicating whether truncation
            occurred.
    #>
    param(
        [string]$Path,
        [int]$MaxChars = 32
    )

    if ([string]::IsNullOrEmpty($Path) -or $Path.Length -le $MaxChars) {
        return @($Path, $false)
    }

    $ellipsis = '\...\'
    $keep = $MaxChars - $ellipsis.Length
    $left = [Math]::Floor($keep / 2)
    $right = $keep - $left

    # Try to break on backslash boundaries for cleaner output
    $parts = $Path -split '\\'
    if ($parts.Count -ge 3) {
        # Build right side: take segments from the end until we hit the budget
        $rightStr = $parts[-1]
        $idx = $parts.Count - 1
        while ($idx -gt 0 -and ($rightStr.Length + $parts[$idx - 1].Length + 1) -le $right) {
            $idx--
            $rightStr = $parts[$idx] + '\' + $rightStr
        }

        # Build left side
        $leftStr = $parts[0]
        $idx = 0
        while ($idx -lt ($parts.Count - 1) -and ($leftStr.Length + $parts[$idx + 1].Length + 1) -le $left) {
            $idx++
            $leftStr = $leftStr + '\' + $parts[$idx]
        }

        if ($leftStr -and $rightStr -and $leftStr -ne $rightStr) {
            return @("$leftStr$ellipsis$rightStr", $true)
        }
    }

    # Character-based fallback when segment-based slicing can't fit
    return @("$($Path.Substring(0, $left))$ellipsis$($Path.Substring($Path.Length - $right))", $true)
}

function New-PillHtmlHC {
    <# 
        .DESCRIPTION
            Render a colored pill — used for status labels in banners and rows.
    #>

    param(
        [string]$Text,
        [string]$Bg,
        [string]$Color = '#ffffff'
    )
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    return "<span style=`"display:inline-block; padding:3px 10px; background-color:$Bg; color:$Color; border-radius:12px; font-size:11px; font-weight:700; letter-spacing:0.3px; text-transform:uppercase; line-height:1.6;`">$Text</span>"
}

function Get-CheckThemeHC {
    param([string]$Type)
    switch ($Type) {
        'FatalError' {
            return @{
                Bg         = $Script:Theme.StatusError
                Accent     = $Script:Theme.AccentError
                Symbol     = '✖'
                Label      = 'ERROR'
                BorderLeft = $Script:Theme.AccentError
            }
        }
        'Warning' {
            return @{
                Bg         = $Script:Theme.StatusWarning
                Accent     = $Script:Theme.AccentWarning
                Symbol     = '⚠'
                Label      = 'WARNING'
                BorderLeft = $Script:Theme.AccentWarning
            }
        }
        default {
            return @{
                Bg         = $Script:Theme.BgAlt
                Accent     = $Script:Theme.AccentInfo
                Symbol     = 'ℹ'
                Label      = 'INFO'
                BorderLeft = $Script:Theme.AccentInfo
            }
        }
    }
}

function Build-ErrorWarningTableHC {
    <#
        .DESCRIPTION
            Build the global "Detected issues" banner shown at the top of the
            email. Renders one red pill for errors and one amber pill for
            warnings. Both counts include matrix-level checks AND script-level
            system errors, (filtered by Type) — the counter object passed in
            is the single source of truth (see Update-MatrixCounterHC).
    #>
    param($CounterData)

    $errs = [int]$CounterData.TotalErrors
    $warns = [int]$CounterData.TotalWarnings

    if ($errs -eq 0 -and $warns -eq 0) { return '' }

    $pills = @()
    if ($errs -gt 0) {
        $errLabel = "$errs Error" + $(if ($errs -ne 1) { 's' })
        $pills += "<td style='padding:0 6px 0 0;'>$(New-PillHtmlHC -Text $errLabel -Bg $Script:Theme.AccentError)</td>"
    }
    if ($warns -gt 0) {
        $warnLabel = "$warns Warning" + $(if ($warns -ne 1) { 's' })
        $pills += "<td style='padding:0 6px 0 0;'>$(New-PillHtmlHC -Text $warnLabel -Bg $Script:Theme.AccentWarning)</td>"
    }

    return @"
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; margin:0 0 16px 0;">
    <tr>
        <td style='padding:4px 0;'>
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">
                <tr>
                    <td style='padding:0 12px 0 0; font-size:13px; font-weight:600; color:$($Script:Theme.TextMain);'>Detected issues</td>
                    $($pills -join '')
                </tr>
            </table>
        </td>
    </tr>
</table>
"@
}

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
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:separate; background-color:$bgColor; border-left:3px solid $accentColor; border-radius:6px;">
            <tr>
                <td valign='middle' width='26' style='padding:10px 0 10px 12px; color:$accentColor; font-size:16px; font-weight:bold; line-height:1; text-align:left;'>$glyph</td>
                <td valign='middle' style='padding:10px 12px 10px 6px;'>
                    <div style='margin-bottom:4px;'>
                        $catHtml<span style='font-weight:700; color:$($Script:Theme.TextMain); font-size:13px;'>$name</span>
                    </div>
                    <div style='color:$($Script:Theme.TextMuted); font-size:12px; line-height:1.5; font-family:$($Script:Theme.MonoStack);'>$msg</div>
                </td>
                <td valign='middle' align='right' width='130' style='padding:10px 12px 10px 8px; width:130px; white-space:nowrap;'>$pill</td>
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
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; margin:0 0 20px 0;">
    <tr>
        <td style='padding:0 0 8px 0; font-size:11px; font-weight:700; color:$($Script:Theme.TextLight); letter-spacing:1.5px; text-transform:uppercase;'>$headerLabel</td>
    </tr>
    $rows
</table>
"@
}

function Build-FileLevelCheckRowHC {
    param(
        [object]$Check,
        [string]$SheetLabel,
        [bool]$IncludeWrapper = $true
    )

    $themeTokens = Get-CheckThemeHC $Check.Type
    $accent = $themeTokens.Accent

    $pillHtml = New-PillHtmlHC -Text $themeTokens.Label -Bg $accent

    $name = [System.Net.WebUtility]::HtmlEncode((Get-StringOrDefaultHC $Check.Name 'Unnamed check'))
    $desc = [System.Net.WebUtility]::HtmlEncode((Get-StringOrDefaultHC $Check.Description ''))
    $label = [System.Net.WebUtility]::HtmlEncode($SheetLabel)

    # The inner card markup is identical in both modes — only the outer
    # wrapper changes. Dot cell geometry (width=30, left padding=10, dot=10px)
    # gives 10px whitespace on each side of the dot — balanced visually,
    # and matches the Settings rows below for consistent vertical alignment.
    $cardHtml = @"
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:separate; background-color:$($Script:Theme.BgWhite); border:1px solid $($Script:Theme.BorderLight); border-left:3px solid $accent; border-radius:6px;">
    <tr>
        <td valign='middle' width='30' style='padding:14px 0 14px 10px;'>
            <span style='display:inline-block; width:10px; height:10px; background-color:$accent; border-radius:50%;'></span>
        </td>
        <td valign='middle' style='padding:10px 8px 10px 0;'>
            <div style='font-size:11px; font-weight:700; color:$($Script:Theme.TextLight); letter-spacing:0.5px; text-transform:uppercase; margin-bottom:2px;'>$label</div>
            <div style='font-size:13px; font-weight:700; color:$($Script:Theme.TextMain); margin-bottom:2px;'>$name</div>
            <div style='font-size:11px; color:$($Script:Theme.TextMuted); line-height:1.5;'>$desc</div>
        </td>
        <td valign='middle' align='right' width='110' style='padding:10px 12px 10px 8px; width:110px; white-space:nowrap;'>$pillHtml</td>
    </tr>
</table>
"@

    if ($IncludeWrapper) {
        # Email body / file card context: wrap in <tr><td> with 16px inset.
        return @"
<tr>
    <td style='padding:0 16px 8px 16px;'>$cardHtml</td>
</tr>
"@
    }
    else {
        # Standalone report context: wrap in <tr><td> with no inset and a
        # bottom margin between rows.
        return @"
<tr>
    <td style='padding:0 0 8px 0;'>$cardHtml</td>
</tr>
"@
    }
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
    $pathParts = Get-TruncatedPathHC -Path $pathRaw -MaxChars 48
    $pathDisp = [System.Net.WebUtility]::HtmlEncode($pathParts[0])
    $pathTitle = if ($pathParts[1]) {
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

    # Two rows: a transparent spacer for separation between cards, then the
    # data row whose cells carry the card styling (left accent border, top/bottom
    # borders, rounded outer corners). All cells share the parent table's column
    # widths so every row aligns column-for-column.
    return @"
<tr>
    <td colspan='6' style='padding:0 0 8px 0; height:0; line-height:0; font-size:0;'>&nbsp;</td>
</tr>
<tr>
    <td valign='middle' style='padding:9px 0 9px 12px; background-color:$($Script:Theme.BgWhite); border-top:1px solid $($Script:Theme.BorderLight); border-bottom:1px solid $($Script:Theme.BorderLight); border-left:3px solid $accent; border-top-left-radius:6px; border-bottom-left-radius:6px;'>
        <a href='$link' style='text-decoration:none; color:inherit; display:block;'>
            <span style='display:inline-block; width:8px; height:8px; background-color:$accent; border-radius:50%;'></span>
        </a>
    </td>
    <td valign='middle' style='padding:9px 8px; background-color:$($Script:Theme.BgWhite); border-top:1px solid $($Script:Theme.BorderLight); border-bottom:1px solid $($Script:Theme.BorderLight); font-weight:700; color:$($Script:Theme.TextMain); font-size:13px; white-space:nowrap;'>
        <a href='$link' style='text-decoration:none; color:inherit;'>$comp</a>
    </td>
    <td valign='middle' style='padding:9px 8px; background-color:$($Script:Theme.BgWhite); border-top:1px solid $($Script:Theme.BorderLight); border-bottom:1px solid $($Script:Theme.BorderLight); font-family:$($Script:Theme.MonoStack); font-size:11px; color:$($Script:Theme.TextMuted); white-space:nowrap; overflow:hidden;'$pathTitle>
        <a href='$link' style='text-decoration:none; color:inherit;'>$pathDisp</a>
    </td>
    <td valign='middle' style='padding:9px 8px; background-color:$($Script:Theme.BgWhite); border-top:1px solid $($Script:Theme.BorderLight); border-bottom:1px solid $($Script:Theme.BorderLight); font-size:11px; color:$($Script:Theme.TextLight); white-space:nowrap;'>
        <a href='$link' style='text-decoration:none; color:inherit;'>$action</a>
    </td>
    <td valign='middle' align='right' style='padding:9px 8px; background-color:$($Script:Theme.BgWhite); border-top:1px solid $($Script:Theme.BorderLight); border-bottom:1px solid $($Script:Theme.BorderLight); font-family:$($Script:Theme.MonoStack); font-size:11px; color:$($Script:Theme.TextLight); white-space:nowrap;'>
        <a href='$link' style='text-decoration:none; color:inherit;'>$dur</a>
    </td>
    <td valign='middle' align='right' style='padding:9px 12px 9px 8px; background-color:$($Script:Theme.BgWhite); border-top:1px solid $($Script:Theme.BorderLight); border-bottom:1px solid $($Script:Theme.BorderLight); border-right:1px solid $($Script:Theme.BorderLight); border-top-right-radius:6px; border-bottom-right-radius:6px; white-space:nowrap;'>
        <a href='$link' target='_blank' rel='noopener noreferrer'  style='text-decoration:none; color:inherit;'>$pillHtml</a>
    </td>
</tr>
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
        filename in the gradient header.
        2. $reportLink — opens the standalone execution report HTML. Used by
        the "Open full report &rarr;" footer link. 
    #>
    $matrixPath = Get-StringOrDefaultHC $FileContext.Item.FullName ''
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

    # Settings rows table — shared <colgroup> guarantees alignment across all rows
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
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;">
            <colgroup>
                <col style="width:22px;">
                <col style="width:108px;">
                <col>
                <col style="width:60px;">
                <col style="width:65px;">
                <col style="width:100px;">
            </colgroup>
            $settingsRowsHtml
        </table>
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
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:separate; margin:0 0 24px 0; background-color:$($Script:Theme.BgWhite); border:1px solid $($Script:Theme.BorderLight); border-radius:10px; overflow:hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.06);">
    <tr>
        <td bgcolor="$gradTo" style='padding:0; background-color:$gradTo; background-image: linear-gradient(135deg, $gradFrom 0%, $gradTo 100%); border-bottom:1px solid $($Script:Theme.BorderLight);'>
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;">
                <tr>
                    <td valign='middle' width='44' style='padding:14px 0 14px 18px; font-size:20px; font-weight:bold; color:#ffffff; line-height:1; text-align:left;'>$headerSymbol</td>
                    <td valign='middle' style='padding:14px 10px;'>
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

function Generate-MailBodyHtmlHC {
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
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="$bodyWidth" style="border-collapse:collapse; width:${bodyWidth}px;">
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

function Build-ExecutionDetailsBlockHC {
    param(
        [object]$FileResult,
        [string]$DefaultsFilePath
    )

    # Helper: turn a Windows path into a clickable <a href="file://..."> link
    function Convert-PathToFileLink {
        param([string]$Path)
        if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
        $displayHtml = [System.Net.WebUtility]::HtmlEncode($Path)
        $urlHtml = [System.Net.WebUtility]::HtmlEncode(
            (ConvertTo-FileUrlHC $Path)
        )
        return "<a href=`"$urlHtml`" style=`"color:$($Script:Theme.LinkColor); text-decoration:none;`">$displayHtml</a>"
    }

    # Gather values (any missing/empty values are simply skipped)
    $matrixPath = Get-StringOrDefaultHC $FileResult.Item.FullName ''
    $defaultsPath = Get-StringOrDefaultHC $DefaultsFilePath ''

    $lastChange = Format-LastChangeHC `
        -LastModifiedBy $FileResult.ExcelInfo.LastModifiedBy `
        -Modified $FileResult.ExcelInfo.Modified
    $lastChangeValue = $lastChange -replace '^Last change:\s*', ''

    $startTime = if ($FileResult.JobTime.StartTime -is [datetime]) {
        $FileResult.JobTime.StartTime.ToString('dd/MM/yyyy HH:mm:ss')
    }
    else { '' }
    $endTime = if ($FileResult.JobTime.EndTime -is [datetime]) {
        $FileResult.JobTime.EndTime.ToString('dd/MM/yyyy HH:mm:ss')
    }
    else { '' }

    # Each row: (label, value-html, use-mono-font?)
    $items = @(
        @{ Label = 'Matrix file'; Value = (Convert-PathToFileLink $matrixPath); Mono = $true }
        @{ Label = 'Defaults file'; Value = (Convert-PathToFileLink $defaultsPath); Mono = $true }
        @{ Label = 'Last change'; Value = $lastChangeValue; Mono = $false }
        @{ Label = 'Start time'; Value = [System.Net.WebUtility]::HtmlEncode($startTime); Mono = $true }
        @{ Label = 'End time'; Value = [System.Net.WebUtility]::HtmlEncode($endTime); Mono = $true }
    )

    $rowsHtml = ''
    foreach ($item in $items) {
        if ([string]::IsNullOrWhiteSpace($item.Value)) { continue }
        $valueStyle = if ($item.Mono) {
            "font-family:$($Script:Theme.MonoStack); font-size:12px;"
        }
        else { 'font-size:13px;' }

        $rowsHtml += @"
<tr>
    <td valign='top' style='padding:8px 16px 8px 0; font-size:11px; font-weight:700; color:$($Script:Theme.TextLight); text-transform:uppercase; letter-spacing:0.5px; white-space:nowrap; width:120px;'>$($item.Label)</td>
    <td valign='top' style='padding:8px 0; color:$($Script:Theme.TextMuted); $valueStyle word-break:break-all;'>$($item.Value)</td>
</tr>
"@
    }

    return @"
<details style='margin:32px 0 0 0;' open>
    <summary style='cursor:pointer; padding:12px 16px; background-color:$($Script:Theme.BgWhite); border:1px solid $($Script:Theme.BorderLight); border-radius:8px; font-size:11px; font-weight:700; color:$($Script:Theme.TextLight); letter-spacing:1.5px; text-transform:uppercase; list-style:none; user-select:none;'>
        Execution details
    </summary>
    <div style='padding:12px 16px 4px 16px; background-color:$($Script:Theme.BgWhite); border:1px solid $($Script:Theme.BorderLight); border-top:none; border-radius:0 0 8px 8px; margin-top:-1px;'>
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;">
            $rowsHtml
        </table>
    </div>
</details>
"@
}

function Write-MatrixExecutionReportHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$FileResult,
        [Parameter(Mandatory)][hashtable]$Html,
        [Parameter(Mandatory)][string]$LogFolder,
        [Parameter(Mandatory = $false)][string]$DefaultsFilePath
    )

    if (-not (Test-Path -LiteralPath $LogFolder -PathType Container)) {
        return $null
    }

    $fileName = [System.Net.WebUtility]::HtmlEncode($FileResult.Item.Name)

    $lastChangeInfo = Format-LastChangeHC `
        -LastModifiedBy $FileResult.ExcelInfo.LastModifiedBy `
        -Modified $FileResult.ExcelInfo.Modified

    # Tally for header status pill
    $allChecks = @()
    if ($FileResult.Check) { $allChecks += $FileResult.Check }
    if ($FileResult.Sheets.FormData.Check) { $allChecks += $FileResult.Sheets.FormData.Check }
    if ($FileResult.Sheets.Permissions.Check) { $allChecks += $FileResult.Sheets.Permissions.Check }
    if ($FileResult.Matrices) {
        foreach ($m in $FileResult.Matrices) {
            if ($m.Check) { $allChecks += $m.Check }
        }
    }
    $fileErrs = @($allChecks | Where-Object Type -EQ 'FatalError').Count
    $fileWarns = @($allChecks | Where-Object Type -EQ 'Warning').Count

    if ($fileErrs -gt 0) {
        $hdrSymbol = '✖'
        $gradFrom, $gradTo = $Script:Theme.GradError
    }
    elseif ($fileWarns -gt 0) {
        $hdrSymbol = '⚠'
        $gradFrom, $gradTo = $Script:Theme.GradWarning
    }
    else {
        $hdrSymbol = '✓'
        $gradFrom, $gradTo = $Script:Theme.GradSuccess
    }

    $hdrLabel = Format-IssueCountLabelHC -Errors $fileErrs -Warnings $fileWarns

    # ---- File Issues block: render each file-level check as a detailed card ----
    $fileIssuesHtml = ''
    $fileLevelGroups = @(
        @{ Label = 'Excel File'; Checks = $FileResult.Check }
        @{ Label = 'FormData Sheet'; Checks = $FileResult.Sheets.FormData.Check }
        @{ Label = 'Permissions Sheet'; Checks = $FileResult.Sheets.Permissions.Check }
    )
    $fileLevelCount = 0
    foreach ($g in $fileLevelGroups) {
        if ($g.Checks) { $fileLevelCount += @($g.Checks).Count }
    }

    if ($fileLevelCount -gt 0) {
        $issueRows = ''
        foreach ($g in $fileLevelGroups) {
            if ($g.Checks) {
                foreach ($c in $g.Checks) {
                    # Standalone report: include the 16px inset wrapper so File Issues rows
                    # have the same indented look as the Settings rows below them.
                    $issueRows += Build-FileLevelCheckRowHC -Check $c -SheetLabel $g.Label
                }
            }
        }
        $fileIssuesHtml = @"
<h2 style="font-size:11px; font-weight:700; color:$($Script:Theme.TextLight); letter-spacing:1.5px; text-transform:uppercase; margin:24px 0 12px 0;">File Issues ($fileLevelCount)</h2>
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;">
$issueRows
</table>
"@
    }

    # ---- Per-matrix detail sections: each matrix row gets a full card showing every check ----
    # Cards now return <tr> markup wrapped in a 16px-inset padding cell, so we wrap them in
    # a <table> to make the inset apply correctly (matching the File Issues table).
    $matrixDetailsHtml = ''
    if ($FileResult.Matrices) {
        $sortedMatrices = $FileResult.Matrices |
        Sort-Object { $_.Setting.Formatted.ComputerName }, { $_.Setting.Formatted.Path }, { $_.ID }

        $matrixRowsHtml = ''
        foreach ($m in $sortedMatrices) {
            $matrixRowsHtml += Build-MatrixDetailCardHC -MatrixItem $m
        }
        $matrixDetailsHtml = @"
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;">
$matrixRowsHtml
</table>
"@
    }
    elseif ($fileLevelCount -eq 0) {
        $matrixDetailsHtml = @"
<p style='padding:12px 16px; color:$($Script:Theme.TextLight); font-style:italic;'>
    No settings rows were processed for this file.
</p>
"@
    }

    # ---- Execution details block (collapsible, at the bottom) ----
    $executionDetailsHtml = Build-ExecutionDetailsBlockHC `
        -FileResult $FileResult `
        -DefaultsFilePath $DefaultsFilePath

    # Settings section header — only show if there are matrices
    $settingsHeaderHtml = if ($FileResult.Matrices -and @($FileResult.Matrices).Count -gt 0) {
        "<h2 style=`"font-size:11px; font-weight:700; color:$($Script:Theme.TextLight); letter-spacing:1.5px; text-transform:uppercase; margin:24px 0 12px 0;`">Settings ($(@($FileResult.Matrices).Count))</h2>"
    }
    else { '' }

    # CSS for the <details>/<summary> element — hides the default marker triangle
    # so our custom styling reads clean.
    $detailsCss = @'
<style type="text/css">
    details summary::-webkit-details-marker { display: none; }
    details summary::marker { display: none; }
</style>
'@

    # ---- Final HTML ----
    $reportHtml = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Execution Report - $fileName</title>
$($Html.Style)
$($Html.TroubleshootingStyle)
$detailsCss
</head>
<body>
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; background-color:$($Script:Theme.BgPage);">
    <tr>
        <td align="left" valign="top" style="padding:0;">
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="900" style="border-collapse:collapse; width:900px; max-width:100%;">
                <!-- File header -->
                <tr>
                    <td style="padding:0 0 24px 0;">
                        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:separate; background-color:$($Script:Theme.BgWhite); border:1px solid $($Script:Theme.BorderLight); border-radius:10px; overflow:hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.06);">
                            <tr>
                                <td bgcolor="$gradTo" style='padding:0; background-color:$gradTo; background-image: linear-gradient(135deg, $gradFrom 0%, $gradTo 100%); border-bottom:1px solid $($Script:Theme.BorderLight);'>
                                    <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;">
                                        <tr>
                                            <td valign='middle' width='52' style='padding:18px 0 18px 22px; font-size:24px; font-weight:bold; color:#ffffff; line-height:1; text-align:left;'>$hdrSymbol</td>
                                            <td valign='middle' style='padding:18px 10px;'>
                                                <div style='font-size:11px; font-weight:700; color:rgba(255,255,255,0.8); text-transform:uppercase; letter-spacing:1.5px; margin-bottom:4px;'>Execution Report</div>
                                                <div style='font-size:20px; font-weight:700; color:#ffffff; line-height:1.25;'>$fileName</div>
                                                <div style='font-size:12px; color:rgba(255,255,255,0.85); line-height:1.4; margin-top:4px;font-style:italic;'>
                                                    $lastChangeInfo
                                                </div>
                                            </td>
                                            <td valign='middle' align='right' style='padding:18px 22px 18px 10px; white-space:nowrap;'>
                                                <span style="font-size:13px; font-weight:700; color:#ffffff; text-transform:uppercase; letter-spacing:0.5px;">$hdrLabel</span>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                $(if ($fileIssuesHtml) { "<tr><td style='padding:0;'>$fileIssuesHtml</td></tr>" })
                <tr>
                    <td style='padding:0;'>
                        $settingsHeaderHtml
                        $matrixDetailsHtml
                    </td>
                </tr>
                <tr>
                    <td style='padding:0;'>$executionDetailsHtml</td>
                </tr>
            </table>
        </td>
    </tr>
</table>
</body>
</html>
"@

    $logFilePath = Join-Path $LogFolder '00 - Execution Report.html'
    $reportHtml | Out-File -FilePath $logFilePath -Encoding UTF8 -Force
}

function Build-MatrixDetailCardHC {
    param([object]$MatrixItem)

    # Determine card status
    $err = @($MatrixItem.Check | Where-Object Type -EQ 'FatalError').Count
    $warn = @($MatrixItem.Check | Where-Object Type -EQ 'Warning').Count
    $hasChecks = ($err + $warn) -gt 0

    if ($err -gt 0) {
        $accent = $Script:Theme.AccentError
    }
    elseif ($warn -gt 0) {
        $accent = $Script:Theme.AccentWarning
    }
    else {
        $accent = $Script:Theme.AccentSuccess
    }
    $statusLabel = Format-IssueCountLabelHC -Errors $err -Warnings $warn

    # Extract & encode row values
    $idFull = Get-StringOrDefaultHC $MatrixItem.ID 'N/A'
    $idShort = if ($idFull.Length -gt 9) {
        "$($idFull.Substring(0, 3))...$($idFull.Substring($idFull.Length - 3))"
    }
    else { $idFull }
    $idShortHtml = [System.Net.WebUtility]::HtmlEncode($idShort)
    $idFullHtml = [System.Net.WebUtility]::HtmlEncode($idFull)

    $comp = [System.Net.WebUtility]::HtmlEncode((Get-StringOrDefaultHC $MatrixItem.Setting.Formatted.ComputerName ''))
    $path = [System.Net.WebUtility]::HtmlEncode((Get-StringOrDefaultHC $MatrixItem.Setting.Formatted.Path ''))
    $action = [System.Net.WebUtility]::HtmlEncode((Get-StringOrDefaultHC $MatrixItem.Setting.Formatted.Action ''))

    $dur = if ($MatrixItem.JobTime.Duration) {
        '{0:00}:{1:00}:{2:00}' -f $MatrixItem.JobTime.Duration.Hours, $MatrixItem.JobTime.Duration.Minutes, $MatrixItem.JobTime.Duration.Seconds
    }
    else { 'N/A' }

    # Optional metadata — only shown if present on the matrix item
    $groupName = [System.Net.WebUtility]::HtmlEncode((Get-StringOrDefaultHC $MatrixItem.Setting.Formatted.GroupName ''))
    $siteCode = [System.Net.WebUtility]::HtmlEncode((Get-StringOrDefaultHC $MatrixItem.Setting.Formatted.SiteCode ''))
    $applyDefaultVal = $MatrixItem.Setting.Formatted.ApplyDefaultPermissions
    $applyDefaultStr = if ($null -ne $applyDefaultVal -and $applyDefaultVal) { 'Yes' } else { 'No' }

    $dotHtml = "<span style='display:inline-block; width:10px; height:10px; background-color:$accent; border-radius:50%;'></span>"

    # Helper: render a single metadata cell with label above value.
    # Used to build a 2-row × 3-pair grid rather than a 6-row × 2-column stack
    # — much shorter vertically.
    function New-MetaCellHtml {
        param(
            [string]$Label,
            [string]$Value,
            [bool]$Mono = $false,
            [string]$TitleAttr = ''
        )
        $valueStyle = if ($Mono) { "font-family:$($Script:Theme.MonoStack); font-size:11px;" } else { 'font-size:12px;' }
        $titleHtml = if ($TitleAttr) { " title=`"$TitleAttr`"" } else { '' }
        return "<td valign='middle'$titleHtml style='padding:3px 28px 3px 0; white-space:nowrap;'><div style='font-size:10px; font-weight:700; color:$($Script:Theme.TextLight); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:1px;'>$Label</div><div style='color:$($Script:Theme.TextMuted); $valueStyle'>$Value</div></td>"
    }

    # Build a 2-row × 3-cell grid. Anchors stay fixed for consistent scanning:
    # Row 1: Action | Duration | ID  (the operational essentials)
    # Row 2: Group  | Site     | Apply Defaults  (configuration context)
    # If Group or Site is missing, the cell falls back to a non-breaking
    # space so column positions stay stable across rows.
    $row1Cells = @(
        (New-MetaCellHtml -Label 'Action' -Value $action)
        (New-MetaCellHtml -Label 'Duration' -Value $dur -Mono $true)
        (New-MetaCellHtml -Label 'ID' -Value $idShortHtml -Mono $true -TitleAttr $idFullHtml)
    )
    $row2Cells = @(
        $(if ($groupName) { New-MetaCellHtml -Label 'Group' -Value $groupName } else { '<td>&nbsp;</td>' })
        $(if ($siteCode) { New-MetaCellHtml -Label 'Site' -Value $siteCode } else { '<td>&nbsp;</td>' })
        (New-MetaCellHtml -Label 'Apply Defaults' -Value $applyDefaultStr)
    )

    $metadataTable = "<table role='presentation' cellpadding='0' cellspacing='0' border='0' style='border-collapse:collapse;'><tr>$($row1Cells -join '')</tr><tr>$($row2Cells -join '')</tr></table>"

    # Three-column horizontal header — no visible dividers, just consistent
    # padding. Dot cell has 10px whitespace on each side (width=30, left
    # padding=10, dot=10) so the text after the dot starts close to it.
    # Metadata column width=460 comfortably holds 3 nowrap cells in 2 rows.
    $headerBlock = @"
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;">
    <tr>
        <td valign='middle' width='30' style='padding:14px 0 14px 10px;'>$dotHtml</td>
        <td valign='middle' style='padding:14px 16px 14px 0;'>
            <div style='font-size:14px; font-weight:700; color:$($Script:Theme.TextMain); line-height:1.25;'>$comp</div>
            <div style='font-size:12px; color:$($Script:Theme.TextMuted); font-family:$($Script:Theme.MonoStack); line-height:1.4; margin-top:2px; word-break:break-all;'>$path</div>
        </td>
        <td valign='middle' width='460' style='padding:12px 16px;'>
            $metadataTable
        </td>
        <td valign='middle' align='right' width='110' style='padding:14px 16px 14px 10px; white-space:nowrap;'>
            <span style="font-size:11px; font-weight:700; color:$accent; text-transform:uppercase; letter-spacing:0.5px;">$statusLabel</span>
        </td>
    </tr>
</table>
"@

    $borderStyle = "border:1px solid $($Script:Theme.BorderLight); border-left:3px solid $accent;"

    # ---------- COMPACT MODE: success rows ----------
    if (-not $hasChecks) {
        $cardHtml = @"
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:separate; background-color:$($Script:Theme.BgWhite); $borderStyle border-radius:8px; overflow:hidden;">
    <tr><td style='padding:0;'>$headerBlock</td></tr>
</table>
"@
        # Wrap with 16px horizontal inset to align with File Issues rows
        return @"
<tr>
    <td style='padding:0 16px 12px 16px;'>$cardHtml</td>
</tr>
"@
    }

    # ---------- FULL MODE: rows with errors/warnings ----------
    $checkRows = ''
    foreach ($c in $MatrixItem.Check) {
        $tt = Get-CheckThemeHC $c.Type
        $name = [System.Net.WebUtility]::HtmlEncode((Get-StringOrDefaultHC $c.Name 'Unnamed check'))
        $desc = [System.Net.WebUtility]::HtmlEncode((Get-StringOrDefaultHC $c.Description ''))

        if (-not [string]::IsNullOrWhiteSpace($c.JsonFileName)) {
            $nameHtml = "<a href='$([System.Net.WebUtility]::HtmlEncode($c.JsonFileName))' style='color:$($Script:Theme.TextMain); text-decoration:underline;'>$name</a>"
        }
        else {
            $nameHtml = $name
        }

        $pillHtml = New-PillHtmlHC -Text $tt.Label -Bg $tt.Accent

        $checkRows += @"
<tr>
    <td style='padding:0 0 8px 0;'>
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:separate; background-color:$($tt.Bg); border-left:3px solid $($tt.BorderLeft); border-radius:6px;">
            <tr>
                <td valign='middle' width='36' style='padding:12px 0 12px 12px; text-align:left; color:$($tt.Accent); font-size:18px; font-weight:bold; line-height:1;'>$($tt.Symbol)</td>
                <td valign='middle' style='padding:12px 12px 12px 0;'>
                    <div style='font-size:14px; font-weight:700; color:$($Script:Theme.TextMain); margin-bottom:4px;'>$nameHtml</div>
                    <div style='font-size:13px; color:$($Script:Theme.TextMuted); line-height:1.55;'>$desc</div>
                </td>
                <td valign='middle' align='right' width='110' style='padding:12px 14px 12px 8px; white-space:nowrap;'>$pillHtml</td>
            </tr>
        </table>
    </td>
</tr>
"@
    }

    $cardHtml = @"
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:separate; background-color:$($Script:Theme.BgWhite); $borderStyle border-radius:8px; overflow:hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
    <tr><td style='padding:0; border-bottom:1px solid $($Script:Theme.BorderLight);'>$headerBlock</td></tr>
    <tr>
        <td style='padding:14px 18px 8px 18px;'>
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;">
                $checkRows
            </table>
        </td>
    </tr>
</table>
"@

    # Wrap with 16px horizontal inset to align with File Issues rows
    return @"
<tr>
    <td style='padding:0 16px 12px 16px;'>$cardHtml</td>
</tr>
"@
}

function New-HtmlCheckRowHC {
    param([object]$CheckItem)
    # Simple two-cell row, kept minimal — used (if at all) by ad-hoc consumers.
    $cls = Get-HtmlClassProbTypeHC $CheckItem.Type
    $name = [System.Net.WebUtility]::HtmlEncode($CheckItem.Name)
    $desc = [System.Net.WebUtility]::HtmlEncode($CheckItem.Description)
    return "<tr class='$cls'><td style='padding:8px 6px; font-weight:600;'>$name</td><td style='padding:8px 6px;'>$desc</td></tr>"
}

function New-HtmlSectionHC {
    param([string]$Title, [array]$Checks)
    # Build a flat section using the new file-level check row style.
    $out = ''
    if (-not [string]::IsNullOrWhiteSpace($Title)) {
        $out += "<tr><td style='padding:14px 16px 6px 16px; font-size:11px; font-weight:700; color:$($Script:Theme.TextLight); letter-spacing:1.5px; text-transform:uppercase;'>$([System.Net.WebUtility]::HtmlEncode($Title))</td></tr>"
    }
    foreach ($c in $Checks) {
        $out += Build-FileLevelCheckRowHC -Check $c -SheetLabel $Title
    }
    return $out
}

function New-SettingsCardHtmlHC {
    param(
        [Parameter(Mandatory)][object]$MatrixItem,
        [Parameter()][bool]$FileHasFatalError = $false
    )
    return Build-MatrixDetailCardHC -MatrixItem $MatrixItem
}

function New-SettingsOverviewHtmlHC {
    param([array]$MatrixRows, [hashtable]$Html)
    # No-op in the new layout — overview is now embedded in each file card.
    return ''
}

function Write-MatrixSettingLogHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Matrix,
        [Parameter(Mandatory)][hashtable]$Html,
        [Parameter(Mandatory)][string]$LogFolder
    )
    if (-not (Test-Path -LiteralPath $LogFolder -PathType Container)) { return $null }

    $detail = Build-MatrixDetailCardHC -MatrixItem $Matrix

    $htmlOut = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
$($Html.Style)
</head>
<body>
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; background-color:$($Script:Theme.BgPage);">
    <tr>
        <td align="left" valign="top" style="padding:0;">
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="900" style="border-collapse:collapse; width:900px; max-width:100%;">
                <tr><td style="padding:0;"><h1>Settings Log &mdash; ID $([System.Net.WebUtility]::HtmlEncode($Matrix.ID))</h1></td></tr>
                <tr><td style="padding:16px 0 0 0;">
                    <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;">
                        $detail
                    </table>
                </td></tr>
            </table>
        </td>
    </tr>
</table>
</body>
</html>
"@

    $logFilePath = Join-Path -Path $LogFolder -ChildPath "ID $($Matrix.ID) - Settings.html"
    $htmlOut | Out-File -FilePath $logFilePath -Encoding UTF8 -Force
}

function New-OverviewHtmlHC {
    <#
    .SYNOPSIS
        Builds the standalone overview HTML page from FormData rows.
    .DESCRIPTION
        Returns an HTML string suitable for writing to a .html file that a
        user can open in a browser. The page lists each matrix file by
        category and links to the matrix file plus the responsible parties.
    .PARAMETER FormData
        Array of objects, each representing one matrix file.
    .OUTPUTS
        [string] Complete HTML page content.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$FormData
    )

    $style = @'
<style type="text/css">
body {
    background-color: #f0f0f0;
    color: #004e2b;
    font-family: Arial, sans-serif;
    padding: 20px;
}
a { color: #004e2b; text-decoration: none; }
a:hover { color: #00dd39; text-decoration: underline; }
h1 {
    border-bottom: 2px solid #004e2b;
    padding-bottom: 10px;
    margin-bottom: 25px;
    color: #004e2b;
    text-transform: uppercase;
    font-size: 1.8em;
}
table {
    width: 100%;
    max-width: 1200px;
    margin: 20px 0;
    border-collapse: separate;
    border-spacing: 0;
    box-shadow: 0 6px 15px rgba(0, 0, 0, 0.2);
    background-color: #ffffff;
    border-radius: 8px;
    overflow: hidden;
    table-layout: auto;
    border: none;
}
table th {
    background-color: #004e2b;
    color: #ffffff;
    text-align: left;
    padding: 15px 20px;
    font-weight: bold;
    text-transform: uppercase;
    border: none;
    font-size: 0.9em;
}
table thead tr:first-child th:first-child { border-top-left-radius: 8px; }
table thead tr:first-child th:last-child  { border-top-right-radius: 8px; }
table th:nth-child(3) { text-align: left; word-break: normal; }
table td {
    text-align: center;
    padding: 10px 15px;
    border: none;
    border-bottom: 1px solid #e0e0e0;
    vertical-align: middle;
    color: #004e2b;
}
table tbody tr:last-child td { border-bottom: none; }
table td:nth-child(3),
table td:nth-child(4),
table td:nth-child(5) {
    text-align: left;
    white-space: nowrap;
    word-break: normal;
    overflow: hidden;
    text-overflow: ellipsis;
}
table tbody tr:nth-child(even) { background-color: #f8f8f8b7; }
table tbody tr:nth-child(odd)  { background-color: #ffffff; }
table tbody tr:hover { background-color: #c2ebcf; color: #004e2b; }
table tbody tr td a { display: block; width: 100%; height: 100%; color: #004e2b; }
table td:last-child a { display: inline; color: #004e2b; }
table tbody tr:hover td a { color: #004e2b; }
</style>
'@

    $rows = $FormData |
    Sort-Object -Property 'MatrixCategoryName', 'MatrixSubCategoryName', 'MatrixFolderDisplayName' |
    ForEach-Object {
        $emailLinks = foreach ($email in ($_.MatrixResponsible -split ',')) {
            $trimmed = $email.Trim()
            "<a href=`"mailto:$trimmed`">$trimmed</a>"
        }

        @"
<tr>
    <td>$([System.Net.WebUtility]::HtmlEncode($_.MatrixCategoryName))</td>
    <td>$([System.Net.WebUtility]::HtmlEncode($_.MatrixSubCategoryName))</td>
    <td><a href="$($_.MatrixFolderDisplayName)">$([System.Net.WebUtility]::HtmlEncode($_.MatrixFolderDisplayName))</a></td>
    <td><a href="$($_.MatrixFilePath)">$([System.Net.WebUtility]::HtmlEncode($_.MatrixFileName))</a></td>
    <td>$($emailLinks -join ' ')</td>
</tr>
"@
    }

    @"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Matrix files overview</title>
$style
</head>
<body>
<h1>Matrix files overview</h1>
<table>
    <thead>
        <tr>
            <th>Category</th>
            <th>Subcategory</th>
            <th>Folder</th>
            <th>Link to the matrix</th>
            <th>Responsible</th>
        </tr>
    </thead>
    <tbody>
        $($rows -join "`n        ")
    </tbody>
</table>
</body>
</html>
"@
}
