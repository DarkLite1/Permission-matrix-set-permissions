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
        overflow-x: hidden;
    }
    /* Force the page-root table to honor the viewport instead of its
       declared width. The report is browser-only, so we override the
       email-compatibility 900px width when the viewport is narrower. */
    body > table { max-width: 100% !important; }
    body > table table { max-width: 100% !important; }
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
    /* Settings rows are fluid flex cards. On narrow reading panes the
       identifier and metadata wrap below the status line instead of the
       status pill clipping off the right edge. New Outlook / Outlook web
       use a Chromium engine, so this media query applies as in a browser. */
    @media (max-width: 520px) {
        .rr-srow { flex-wrap: wrap; }
        .rr-srow .rr-srow-ident { flex-basis: 100%; order: 3; }
        .rr-srow .rr-srow-meta { flex-basis: 100%; order: 4; }
        .rr-syscard { flex-wrap: wrap; }
        .rr-syscard .rr-syscard-body { flex-basis: 100%; order: 3; }
        .rr-check-row { flex-wrap: wrap; }
    }
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

    # Fluid flex card mirroring the settings rows: accent dot, the text block
    # which flexes and wraps on narrow panes, then the status pill on the right
    # (flex:0 0 auto so it never clips — the text block absorbs overflow).
    $cardHtml = @"
<div class="rr-check-row" style="display:flex; align-items:center; gap:16px; background-color:$($Script:Theme.BgWhite); border:1px solid $($Script:Theme.BorderLight); border-left:3px solid $accent; border-radius:6px; padding:12px 14px;">
    <span style='flex:0 0 auto; width:10px; height:10px; background-color:$accent; border-radius:50%;'></span>
    <span style='flex:1 1 auto; min-width:0;'>
        <span style='display:block; font-size:11px; font-weight:700; color:$($Script:Theme.TextLight); letter-spacing:0.5px; text-transform:uppercase; margin-bottom:2px;'>$label</span>
        <span style='display:block; font-size:13px; font-weight:700; color:$($Script:Theme.TextMain); margin-bottom:2px;'>$name</span>
        <span style='display:block; font-size:11px; color:$($Script:Theme.TextMuted); line-height:1.5;'>$desc</span>
    </span>
    <span style='flex:0 0 auto;'>$pillHtml</span>
</div>
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

function Build-ExecutionDetailsBlockHC {
    param(
        [object]$FileResult,
        [string]$DefaultsFilePath,
        [datetime]$ScriptStartTime,
        [datetime]$ScriptEndTime
    )

    # Helper: turn a Windows path into a clickable <a href="file://..."> link
    function Convert-PathToFileLink {
        param([string]$Path)
        if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
        $displayHtml = [System.Net.WebUtility]::HtmlEncode($Path)
        $urlHtml = [System.Net.WebUtility]::HtmlEncode(
            (ConvertTo-FileUrlHC $Path)
        )
        return "<a href=`"$urlHtml`" target='_blank' rel='noopener noreferrer'  style=`"color:$($Script:Theme.LinkColor); text-decoration:none;`">$displayHtml</a>"
    }

    # Gather values (any missing/empty values are simply skipped)
    $matrixPath = Get-StringOrDefaultHC $FileResult.Item.FullName ''
    $defaultsPath = Get-StringOrDefaultHC $DefaultsFilePath ''

    $lastChange = Format-LastChangeHC `
        -LastModifiedBy $FileResult.ExcelInfo.LastModifiedBy `
        -Modified $FileResult.ExcelInfo.Modified
    $lastChangeValue = $lastChange -replace '^Last change:\s*', ''

    $startTime = if ($ScriptStartTime -is [datetime]) {
        $ScriptStartTime.ToString('dd/MM/yyyy HH:mm:ss')
    }
    else { '' }
    $endTime = if ($ScriptEndTime -is [datetime]) {
        $ScriptEndTime.ToString('dd/MM/yyyy HH:mm:ss')
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
    <td class="rr-mono-wrap" valign='top' style='padding:8px 0; color:$($Script:Theme.TextMuted); $valueStyle word-break:break-all;'>$($item.Value)</td>
</tr>
"@
    }

    # Quiet metadata footer. No heading — the content (file paths, timestamps)
    # is self-evident, and a thin horizontal separator above the panel is
    # enough to mark it as a distinct section. The panel spans the full
    # outer-table width (matching the Execution Report header bar at the
    # top of the page)
    return @"
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; margin-top:32px; table-layout:fixed;">
    <tr>
        <td style='padding:0;'>
            <div style='padding:14px 18px 8px 18px; background-color:$($Script:Theme.BgAlt); border-radius:8px;'>
                <table role="presentation" class="rr-footer-grid" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; table-layout:fixed;">
                    <colgroup>
                        <col style="width:120px;">
                        <col>
                    </colgroup>
                    $rowsHtml
                </table>
            </div>
        </td>
    </tr>
</table>
"@
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

    # Two-row compact metadata layout. Column 1 anchors short values (Action,
    # Duration), column 2 holds short labeled values (Apply Defaults, ID),
    # column 3 holds the potentially-long values (Group, Site).
    #
    #   Col 1                Col 2                Col 3
    #   ----------------     ------------------   ---------------
    #   ACTION: x            APPLY DEFAULTS: x    GROUP: x
    #   [clock] Duration     ID: x                SITE: x
    #
    # Duration keeps an inline SVG clock icon (universally readable as "time")
    # in place of a text label. Everything else uses inline "LABEL: value"
    # styling. Column positions are reserved (with &nbsp; fallbacks for
    # missing optional fields) so cells align vertically across the two rows.

    # Inline SVG clock icon — Tabler Icons (MIT). Inline rather than webfont
    # so it renders in both browser file-views and email clients that strip
    # @font-face rules.
    $iconStyle = "width:13px; height:13px; vertical-align:-2px; margin-right:6px; stroke:$($Script:Theme.TextLight); fill:none; stroke-width:2; stroke-linecap:round; stroke-linejoin:round;"
    $iconDuration = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' style='$iconStyle' aria-hidden='true'><circle cx='12' cy='12' r='9'/><polyline points='12 7 12 12 15 15'/></svg>"

    # Helper for the Duration cell — icon in place of a text label.
    function New-IconMetaCellHtml {
        param(
            [string]$IconHtml,
            [string]$Value,
            [bool]$Mono = $false,
            [string]$TitleAttr = ''
        )
        $valueStyle = if ($Mono) { "font-family:$($Script:Theme.MonoStack); font-size:11px;" } else { 'font-size:12px;' }
        $titleHtml = if ($TitleAttr) { " title=`"$TitleAttr`"" } else { '' }
        $valueHtml = "<span$titleHtml style='color:$($Script:Theme.TextMuted); $valueStyle'>$Value</span>"
        return "<td valign='middle' style='padding:3px 28px 3px 0; white-space:nowrap;'>$IconHtml$valueHtml</td>"
    }

    # Helper for inline "LABEL: value" cells — used by every other cell.
    function New-InlineMetaCellHtml {
        param(
            [string]$Label,
            [string]$Value,
            [bool]$Mono = $false,
            [string]$TitleAttr = ''
        )
        $valueStyle = if ($Mono) { "font-family:$($Script:Theme.MonoStack); font-size:11px;" } else { 'font-size:12px;' }
        $titleHtml = if ($TitleAttr) { " title=`"$TitleAttr`"" } else { '' }
        $labelHtml = "<span style='font-size:10px; font-weight:700; color:$($Script:Theme.TextLight); text-transform:uppercase; letter-spacing:0.5px; margin-right:6px;'>$Label`:</span>"
        $valueHtml = "<span$titleHtml style='color:$($Script:Theme.TextMuted); $valueStyle'>$Value</span>"
        return "<td valign='middle' style='padding:3px 28px 3px 0; white-space:nowrap;'>$labelHtml$valueHtml</td>"
    }

    # Row 1: Action          | Apply Defaults | Group
    # Row 2: Duration (icon) | ID             | Site
    $row1Cells = @(
        (New-InlineMetaCellHtml -Label 'Action' -Value $action)
        (New-InlineMetaCellHtml -Label 'Apply Defaults' -Value $applyDefaultStr)
        $(if ($groupName) { New-InlineMetaCellHtml -Label 'Group' -Value $groupName } else { '<td>&nbsp;</td>' })
    )
    $row2Cells = @(
        (New-IconMetaCellHtml -IconHtml $iconDuration -Value $dur -Mono $true)
        (New-InlineMetaCellHtml -Label 'ID' -Value $idShortHtml -Mono $true -TitleAttr $idFullHtml)
        $(if ($siteCode) { New-InlineMetaCellHtml -Label 'Site' -Value $siteCode } else { '<td>&nbsp;</td>' })
    )

    $metadataTable = "<table role='presentation' cellpadding='0' cellspacing='0' border='0' style='border-collapse:collapse;'>" +
    "<tr>$($row1Cells -join '')</tr>" +
    "<tr>$($row2Cells -join '')</tr>" +
    '</table>'

    # Three-column horizontal header — no visible dividers, just consistent
    # padding. table-layout:fixed plus an explicit 55% width on the metadata
    # column gives the three pills (Action, Apply Defaults, Group) enough
    # room to fit on one line at the report's 900px design width, and pushes
    # the long monospace path to wrap onto its own line sooner — leaving
    # more breathing room overall instead of forcing the whole card past
    # the viewport edge.
    $headerBlock = @"
<table role="presentation" class="rr-settings-head" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; table-layout:fixed;">
    <tr>
        <td class="rr-icon-cell" valign='middle' width='40' style='padding:14px 8px 14px 10px;'>$dotHtml</td>
        <td class="rr-content-cell" valign='middle' style='padding:14px 16px 14px 0;'>
            <div style='font-size:14px; font-weight:700; color:$($Script:Theme.TextMain); line-height:1.25;'>$comp</div>
            <div class="rr-path" style='font-size:12px; color:$($Script:Theme.TextMuted); font-family:$($Script:Theme.MonoStack); line-height:1.4; margin-top:2px; word-break:break-all;'>$path</div>
        </td>
        <td class="rr-meta-cell" valign='middle' width='55%' style='padding:12px 16px;'>
            $metadataTable
        </td>
    </tr>
</table>
"@

    $borderStyle = "border:1px solid $($Script:Theme.BorderLight); border-left:3px solid $accent;"

    # ---------- COMPACT MODE: success rows ----------
    if (-not $hasChecks) {
        $cardHtml = @"
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:separate; background-color:$($Script:Theme.BgWhite); $borderStyle border-radius:8px; overflow:hidden; table-layout:fixed;">
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
            $nameHtml = "<a href='$([System.Net.WebUtility]::HtmlEncode($c.JsonFileName))' target='_blank' rel='noopener noreferrer' style='color:$($Script:Theme.TextMain); text-decoration:underline;'>$name</a>"
        }
        else {
            $nameHtml = $name
        }

        $pillHtml = New-PillHtmlHC -Text $tt.Label -Bg $tt.Accent

        $checkRows += @"
<tr>
    <td style='padding:0 0 8px 0;'>
        <table role="presentation" class="rr-check-row" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:separate; background-color:$($tt.Bg); border-left:3px solid $($tt.BorderLeft); border-radius:6px;">
            <tr>
                <td class="rr-icon-cell" valign='middle' width='36' style='padding:12px 0 12px 12px; text-align:left; color:$($tt.Accent); font-size:18px; font-weight:bold; line-height:1;'>$($tt.Symbol)</td>
                <td class="rr-content-cell" valign='middle' style='padding:12px 12px 12px 0;'>
                    <div style='font-size:14px; font-weight:700; color:$($Script:Theme.TextMain); margin-bottom:4px;'>$nameHtml</div>
                    <div style='font-size:13px; color:$($Script:Theme.TextMuted); line-height:1.55;'>$desc</div>
                </td>
                <td class="rr-check-pill" valign='middle' align='right' width='110' style='padding:12px 14px 12px 8px; white-space:nowrap;'>$pillHtml</td>
            </tr>
        </table>
    </td>
</tr>
"@
    }

    $cardHtml = @"
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:separate; background-color:$($Script:Theme.BgWhite); $borderStyle border-radius:8px; overflow:hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.05); table-layout:fixed;">
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

function Write-MatrixExecutionReportHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$FileResult,
        [Parameter(Mandatory)][hashtable]$Html,
        [Parameter(Mandatory)][datetime]$ScriptStartTime,
        [Parameter(Mandatory)][datetime]$ScriptEndTime,
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
        -DefaultsFilePath $DefaultsFilePath `
        -ScriptStartTime $ScriptStartTime `
        -ScriptEndTime $ScriptEndTime

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

    # ---- Responsive CSS (browser-only) ----
    # The page is built with email-compatible table markup (fixed widths,
    # nowrap pill cells, multi-column rows). The rules below collapse those
    # tables into stacked blocks below 900px so the report wraps cleanly
    # on tablet/laptop window resizes without horizontal scroll.
    #
    # All rules are scoped to `.report-root` so they don't affect any other
    # consumer of $Html.Style (the email body in particular).
    #
    # Strategy: at viewports below 900px (the page's design width), collapse
    # the email-compatible multi-column rows into stacked single-column
    # blocks. The pill cells drop below their content, the metadata sub-
    # table (Action, Apply Defaults, Group, etc.) wraps as inline chips,
    # the footer label/value grid stacks label above value, and long
    # monospace paths break at any character so they never force a
    # horizontal scrollbar. Looks good down to ~360px even though the
    # explicit design target is ~768px (tablet).
    $responsiveCss = @'
<style type="text/css">
    .report-root { width: 100%; max-width: 900px; margin: 0; box-sizing: border-box; }
    .report-root * { box-sizing: border-box; }

    /* Anywhere: long monospace paths must be allowed to wrap. */
    .report-root .rr-path,
    .report-root .rr-mono-wrap { word-break: break-all; overflow-wrap: anywhere; }

    /* The metadata sub-table (Action / Apply Defaults / Group / etc.) is
       built with white-space:nowrap on each cell so the cells stay on one
       line in email clients. In the browser we allow them to wrap when
       horizontal room runs out — single-line whenever they fit, multi-line
       when they don't, with no clipping at any width. */
    .report-root .rr-meta-cell table > tbody > tr > td { white-space: normal !important; }

    @media (max-width: 900px) {
        /* Strategy: turn the affected <tr> into a flex container with
           flex-wrap. Cells stay as flex items, which natively gives us
           vertical centering (align-items: center) and the ability to
           force a cell onto its own row via `flex: 1 1 100%`.
           
           Why not the more conventional `display: block` on every cell?
           Because then the icon stacks ABOVE the content instead of beside
           it. And why not `float: left` on the icon? Because then a wrapped
           third line in the content drops UNDER the float and shifts left,
           breaking horizontal alignment with the lines above.
           
           Status pills (rr-status-cell / rr-check-pill) are taken out of
           the flex flow with position:absolute + top:50% + translateY(-50%)
           so they sit middle-right anchored to the relatively-positioned
           parent table. The content cell reserves padding-right to keep
           text from running under the pill. */

        /* Page header */
        .report-root .rr-header-row { position: relative; }
        .report-root .rr-header-row > tbody > tr {
            display: flex; align-items: center;
        }
        .report-root .rr-header-row > tbody > tr > td.rr-icon-cell {
            flex: 0 0 auto; width: 52px !important; text-align: left;
            padding: 18px 0 18px 22px !important;
        }
        .report-root .rr-header-row > tbody > tr > td.rr-content-cell {
            flex: 1 1 auto; min-width: 0;
            padding: 18px 130px 18px 10px !important;
        }
        .report-root .rr-header-row > tbody > tr > td.rr-status-cell {
            position: absolute; top: 50%; right: 22px;
            transform: translateY(-50%);
            padding: 0 !important;
            text-align: right !important; white-space: nowrap !important;
            width: auto !important;
        }

        /* Settings card header: icon + content stay side-by-side (vertically
           centered), meta drops to its own row. flex-wrap:wrap enables the
           wrap; meta's flex-basis of 100% forces it onto a new line. */
        .report-root .rr-settings-head > tbody > tr {
            display: flex; flex-wrap: wrap; align-items: center;
        }
        .report-root .rr-settings-head > tbody > tr > td.rr-icon-cell {
            flex: 0 0 auto; width: 30px !important; text-align: left;
            padding: 14px 0 14px 14px !important;
        }
        .report-root .rr-settings-head > tbody > tr > td.rr-content-cell {
            flex: 1 1 0; min-width: 0;
            padding: 14px 16px 14px 8px !important;
            white-space: normal !important;
        }
        .report-root .rr-settings-head > tbody > tr > td.rr-meta-cell {
            flex: 1 1 100%; width: 100% !important;
            padding: 0 16px 14px 22px !important;
        }

        /* Flow the metadata pill rows as inline-block chips. */
        .report-root .rr-meta-cell table { width: 100% !important; }
        .report-root .rr-meta-cell table,
        .report-root .rr-meta-cell table > tbody { display: block; }
        .report-root .rr-meta-cell table > tbody > tr { display: block; margin-bottom: 2px; }
        .report-root .rr-meta-cell table > tbody > tr > td {
            display: inline-block !important;
            padding: 3px 16px 3px 0 !important; vertical-align: top;
        }

        /* Check rows */
        .report-root .rr-check-row { position: relative; }
        .report-root .rr-check-row > tbody > tr {
            display: flex; align-items: center;
        }
        .report-root .rr-check-row > tbody > tr > td.rr-icon-cell {
            flex: 0 0 auto; width: 36px !important; text-align: left;
            padding: 12px 0 12px 12px !important;
        }
        .report-root .rr-check-row > tbody > tr > td.rr-content-cell {
            flex: 1 1 0; min-width: 0;
            padding: 12px 110px 12px 8px !important;
            white-space: normal !important;
        }
        .report-root .rr-check-row > tbody > tr > td.rr-check-pill {
            position: absolute; top: 50%; right: 14px;
            transform: translateY(-50%);
            padding: 0 !important;
            text-align: right !important; white-space: nowrap !important;
            width: auto !important;
        }

        /* Footer "label : value" rows: stack label above value. */
        .report-root .rr-footer-grid,
        .report-root .rr-footer-grid > tbody { display: block; width: 100%; }
        .report-root .rr-footer-grid > colgroup { display: none; }
        .report-root .rr-footer-grid > tbody > tr { display: block; margin-bottom: 10px; }
        .report-root .rr-footer-grid > tbody > tr > td { display: block; width: auto !important; white-space: normal !important; padding: 2px 0 !important; }
    }
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
$responsiveCss
</head>
<body>
<div class="report-root">
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; background-color:$($Script:Theme.BgPage);">
    <tr>
        <td align="left" valign="top" style="padding:0;">
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; width:100%; max-width:900px;">
                <!-- File header -->
                <tr>
                    <td style="padding:0 0 24px 0;">
                        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:separate; background-color:$($Script:Theme.BgWhite); border:1px solid $($Script:Theme.BorderLight); border-radius:10px; overflow:hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.06);">
                            <tr>
                                <td bgcolor="$gradTo" style='padding:0; background-color:$gradTo; background-image: linear-gradient(135deg, $gradFrom 0%, $gradTo 100%); border-bottom:1px solid $($Script:Theme.BorderLight);'>
                                    <table role="presentation" class="rr-header-row" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;">
                                        <tr>
                                            <td class="rr-icon rr-icon-cell" valign='middle' width='52' style='padding:18px 0 18px 22px; font-size:24px; font-weight:bold; color:#ffffff; line-height:1; text-align:left;'>$hdrSymbol</td>
                                            <td class="rr-content-cell" valign='middle' style='padding:18px 10px;'>
                                                <div style='font-size:11px; font-weight:700; color:rgba(255,255,255,0.8); text-transform:uppercase; letter-spacing:1.5px; margin-bottom:4px;'>Execution Report</div>
                                                <div style='font-size:20px; font-weight:700; color:#ffffff; line-height:1.25;'>$fileName</div>
                                                <div style='font-size:12px; color:rgba(255,255,255,0.85); line-height:1.4; margin-top:4px;font-style:italic;'>
                                                    $lastChangeInfo
                                                </div>
                                            </td>
                                            <td class="rr-status-cell" valign='middle' align='right' style='padding:18px 22px 18px 10px; white-space:nowrap;'>
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
</div>
</body>
</html>
"@

    $logFilePath = Join-Path $LogFolder '00 - Execution Report.html'
    $reportHtml | Out-File -FilePath $logFilePath -Encoding UTF8 -Force
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
    <td><a href="$($_.MatrixFolderDisplayName)" target='_blank' rel='noopener noreferrer' >$([System.Net.WebUtility]::HtmlEncode($_.MatrixFolderDisplayName))</a></td>
    <td><a href="$($_.MatrixFilePath)" target='_blank' rel='noopener noreferrer' >$([System.Net.WebUtility]::HtmlEncode($_.MatrixFileName))</a></td>
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