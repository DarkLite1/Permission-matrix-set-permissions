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
    MonoStack      = "Consolas, Menlo, monospace"

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
    /* Settings rows, system-error cards and file-check rows are table-based
       (not flexbox) so Outlook Classic's Word engine renders them as aligned
       columns. On very narrow reading panes the meta/status columns are
       allowed to wrap so nothing clips off the right edge; the path and
       message cells already wrap via word-break. */
    @media (max-width: 520px) {
        .rr-srow .rr-srow-meta,
        .rr-srow .rr-srow-status,
        .rr-syscard .rr-syscard-status { white-space: normal !important; }
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

function Get-FileCheckTallyHC {
    <#
        .DESCRIPTION
            Tally the FatalError and Warning checks a matrix file carries,
            gathered from every place a check can live: the file itself, its
            'FormData' and 'Permissions' sheets, and each of its matrices.

            This is the tally that decides a card's header colour, glyph and
            status pill, and it is also the sort key used to float problem
            cards to the top of the overview — sharing one implementation
            keeps the two in step.

            Informational checks are deliberately NOT counted. They are
            notices, not issues: they must not colour a header and they must
            not lift a card out of the alphabetical run.

        .PARAMETER FileResult
            One element of $Context.FileResults.
    #>
    param([object]$FileResult)

    $allChecks = @()

    if ($FileResult) {
        if ($FileResult.Check) { $allChecks += $FileResult.Check }
        if ($FileResult.Sheets.FormData.Check) { $allChecks += $FileResult.Sheets.FormData.Check }
        if ($FileResult.Sheets.Permissions.Check) { $allChecks += $FileResult.Sheets.Permissions.Check }

        if ($FileResult.Matrices) {
            foreach ($matrix in $FileResult.Matrices) {
                if ($matrix.Check) { $allChecks += $matrix.Check }
            }
        }
    }

    return @{
        Errors   = @($allChecks | Where-Object Type -EQ 'FatalError').Count
        Warnings = @($allChecks | Where-Object Type -EQ 'Warning').Count
    }
}

function Get-MatrixFileNameHC {
    <#
        .DESCRIPTION
            Resolve the display name of a matrix file from a file-result
            object. Used for the card header title AND as the sort key for
            the overview, so the visible title and the ordering can never
            disagree.

            A file result normally carries an 'Item' (the FileInfo of the
            .xlsx). When the runspace threw before that was set, the fallback
            object built by the catch block carries 'File' instead — the same
            two-step lookup Invoke-PermissionMatrixAuditReport already does.

        .PARAMETER FileResult
            One element of $Context.FileResults.

        .PARAMETER Default
            Returned when neither 'Item' nor 'File' yields a name.
    #>
    param(
        [object]$FileResult,
        [string]$Default = ''
    )

    if (-not $FileResult) { return $Default }

    foreach ($propertyName in @('Item', 'File')) {
        $container = $FileResult.PSObject.Properties[$propertyName]

        if ($container -and $container.Value -and $container.Value.Name) {
            return [string]$container.Value.Name
        }
    }

    return $Default
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
            # Info/informational checks use the same neutral grey as a
            # "Skipped" row (background + accent) so they read as low-key
            # notices instead of an attention-grabbing blue.
            return @{
                Bg         = $Script:Theme.StatusSkipped
                Accent     = $Script:Theme.AccentSkipped
                Symbol     = 'ℹ'
                Label      = 'INFO'
                BorderLeft = $Script:Theme.AccentSkipped
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

            Modern clients and browsers get a CSS `border-radius` span. Outlook
            on Windows (Word rendering engine) ignores border-radius, so an
            MSO-only VML <v:roundrect> with the same fill/text is emitted for
            those clients. Because VML needs an explicit width, it is estimated
            from the (upper-cased) text length. The two variants are gated by
            downlevel-hidden conditional comments so each client renders exactly
            one pill.
    #>

    param(
        [string]$Text,
        [string]$Bg,
        [string]$Color = '#ffffff',
        # When set, returns an object with the raw MSO (VML) and Browser (span)
        # markup separately (WITHOUT the conditional-comment wrappers) so a caller
        # can place each variant in its own client-gated <td>. Default returns the
        # single combined string used everywhere else.
        [switch]$AsParts
    )
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }

    $span = "<span style=`"display:inline-block; padding:3px 10px; background-color:$Bg; color:$Color; border-radius:12px; font-size:11px; font-weight:700; letter-spacing:0.3px; text-transform:uppercase; line-height:1.6;`">$Text</span>"

    # The CSS pill uppercases via text-transform; mirror that in the VML text
    # so both variants match. Width is a generous estimate (uppercase + letter
    # spacing) so the label never clips inside the fixed-width VML shape.
    $upper = $Text.ToUpper()
    $width = [int][Math]::Ceiling(($upper.Length * 8.5) + 26)
    # Word gives the VML shape a fixed height; <w:anchorlock/> + v-text-anchor
    # keeps the single line of text vertically centered so the pill isn't
    # squashed. Height 26px mirrors the browser span (padding 3px + 11px text
    # at line-height 1.6 ≈ 24-26px) so both clients look the same.
    $vmlInner = "<v:roundrect xmlns:v=`"urn:schemas-microsoft-com:vml`" xmlns:w=`"urn:schemas-microsoft-com:office:word`" arcsize=`"50%`" fillcolor=`"$Bg`" stroked=`"f`" style=`"height:26px; width:${width}px; v-text-anchor:middle; mso-padding-alt:0;`">" +
    "<w:anchorlock/>" +
    "<center style=`"color:$Color; font-family:sans-serif; font-size:11px; font-weight:700; letter-spacing:0.3px;`">$upper</center>" +
    "</v:roundrect>"

    if ($AsParts) {
        return [pscustomobject]@{ Mso = $vmlInner; Browser = $span }
    }

    $vml = "<!--[if mso]>$vmlInner<![endif]-->"

    return "$vml<!--[if !mso]><!-->$span<!--<![endif]-->"
}

function Build-ErrorWarningTableHC {
    <#
        .DESCRIPTION
            Build the global "Detected issues" banner shown at the top of the
            email. Renders one red pill for errors and one amber pill for
            warnings. Both counts include matrix-level checks AND script-level
            system errors, (filtered by Type) — the counter object passed in
            is the single source of truth (see Update-MatrixCounterHC).

            Outlook (Word engine) ignores 'margin' on <table>, so the wrapping
            table's 'margin:0 0 16px 0' — the gap below the banner — never
            rendered there. An MSO-only 16px spacer table after the banner
            restores it; browsers skip the conditional and keep the CSS margin.
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
<!--[if mso]>
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td height="16" style="font-size:0; line-height:16px; mso-line-height-rule:exactly;">&#160;</td></tr></table>
<![endif]-->
"@
}

function Build-FileLevelCheckRowHC {
    param(
        [object]$Check,
        [string]$SheetLabel,
        [bool]$IncludeWrapper = $true,
        # When $true and the check has a JSON detail file (only created when
        # the check's 'Value' is not null), the check name becomes a link to
        # that file. The href is relative (just the file name) because the
        # detail JSON is written to the same folder as the execution report.
        # Keep this $false in the email context, where a relative link is
        # meaningless.
        [bool]$LinkJsonDetail = $false,
        # When $false the small uppercase sheet label (e.g. "EXCEL FILE") is
        # not rendered. The email drops it because the "File Issues" section
        # header already conveys the context; the standalone report keeps it
        # (there the label is the matrix file title).
        [bool]$ShowLabel = $true
    )

    $themeTokens = Get-CheckThemeHC $Check.Type
    $accent = $themeTokens.Accent

    # Info notices use the neutral grey card background + an "i" glyph; errors
    # and warnings keep a white card with a bullet dot so they still stand out.
    $isInfo = ($Check.Type -ne 'FatalError') -and ($Check.Type -ne 'Warning')
    $cardBg = if ($isInfo) { $themeTokens.Bg } else { $Script:Theme.BgWhite }
    $icon = if ($isInfo) { '&#8505;' } else { '&#9679;' }

    $pillHtml = New-PillHtmlHC -Text $themeTokens.Label -Bg $accent

    $name = [System.Net.WebUtility]::HtmlEncode((Get-StringOrDefaultHC $Check.Name 'Unnamed check'))
    $desc = [System.Net.WebUtility]::HtmlEncode((Get-StringOrDefaultHC $Check.Description ''))
    $label = [System.Net.WebUtility]::HtmlEncode($SheetLabel)

    $nameHtml = $name
    if (
        $LinkJsonDetail -and
        $Check.PSObject.Properties.Match('JsonFileName').Count -and
        -not [string]::IsNullOrWhiteSpace($Check.JsonFileName)
    ) {
        $jsonHref = [System.Net.WebUtility]::HtmlEncode($Check.JsonFileName)
        $nameHtml = "<a href='$jsonHref' target='_blank' rel='noopener noreferrer' style='color:$($Script:Theme.TextMain); text-decoration:underline;'>$name</a>"
    }

    # Small uppercase sheet label above the check name. Rendered as a block
    # <div> (not a <span display:block>, which Outlook's Word engine ignores)
    # and only when requested and non-empty.
    $labelHtml = ''
    if ($ShowLabel -and -not [string]::IsNullOrWhiteSpace($SheetLabel)) {
        $labelHtml = "<div style='font-size:11px; font-weight:700; color:$($Script:Theme.TextLight); letter-spacing:0.5px; text-transform:uppercase; line-height:14px; margin:0 0 2px 0; mso-line-height-rule:exactly;'>$label</div>"
    }

    # Table-based card mirroring the settings rows so Outlook Classic (Word
    # engine, no flexbox) renders the accent icon, the text block and the status
    # pill as aligned columns. Browsers render the same table identically. The
    # name/description use <div> blocks (with margin:0 + exact line-height) so
    # the name stacks ABOVE the description in Outlook too — Word collapses
    # <span style='display:block'> onto one inline line. The icon and pill cells
    # are vertically centered (valign='middle' attribute — honored by Word,
    # unlike the CSS property alone) while the text cell stays top-aligned; the
    # icon cell also centers horizontally via align='center' + text-align.
    $cardHtml = @"
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" class="rr-check-row" style="border-collapse:separate; width:100%; max-width:100%; background-color:$cardBg; border:1px solid $($Script:Theme.BorderLight); border-left:3px solid $accent; border-radius:6px;">
    <tr>
        <td valign="middle" align="center" width="24" style='vertical-align:middle; text-align:center; padding:12px 0 12px 14px; color:$accent; font-size:15px; line-height:16px; mso-line-height-rule:exactly;'>$icon</td>
        <td valign="top" style='padding:12px 10px;'>
            $labelHtml<div style='font-size:13px; font-weight:700; color:$($Script:Theme.TextMain); line-height:16px; margin:0 0 2px 0; mso-line-height-rule:exactly;'>$nameHtml</div>
            <div style='font-size:11px; color:$($Script:Theme.TextMuted); line-height:15px; margin:0; mso-line-height-rule:exactly;'>$desc</div>
        </td>
        <td valign="middle" align="right" style='vertical-align:middle; padding:12px 14px 12px 6px; white-space:nowrap;'>$pillHtml</td>
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