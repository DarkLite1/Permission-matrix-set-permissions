# HtmlCommon.ps1
# Shared HTML theme + low-level primitives used by BOTH the email body
# (HtmlMailBody.ps1) and the on-disk report (HtmlReport.ps1).
# Must be loaded before those two: it defines $Script:Theme.

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

