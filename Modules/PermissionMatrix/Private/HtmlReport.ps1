# HtmlReport.ps1
# Builds the standalone on-disk execution/overview HTML report. Depends on
# HtmlCommon.ps1 for $Script:Theme and shared primitives. Load HtmlCommon.ps1 first.

function Build-ExecutionDetailsBlockHC {
    param(
        [object]$FileResult,
        [string]$DefaultsFilePath,
        [datetime]$ScriptStartTime,
        [datetime]$ScriptEndTime
    )

    # Helper: turn a Windows path into a clickable <a href="file://..."> link.
    # An optional -Title renders as a hover tooltip on the link.
    function Convert-PathToFileLink {
        param(
            [string]$Path,
            [string]$Title
        )
        if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
        $displayHtml = [System.Net.WebUtility]::HtmlEncode($Path)
        $urlHtml = [System.Net.WebUtility]::HtmlEncode(
            (ConvertTo-FileUrlHC $Path)
        )
        $titleAttr = if ($Title) {
            " title=`"$([System.Net.WebUtility]::HtmlEncode($Title))`""
        }
        else { '' }
        return "<a href=`"$urlHtml`"$titleAttr target='_blank' rel='noopener noreferrer'  style=`"color:$($Script:Theme.LinkColor); text-decoration:none;`">$displayHtml</a>"
    }

    # Gather values (any missing/empty values are simply skipped)
    # When the matrix file was archived (Matrix.Archive = true), the original
    # path no longer exists — link to the archived copy instead.
    $matrixPath = if (
        $FileResult.PSObject.Properties.Match('ArchivedPath').Count -and
        -not [string]::IsNullOrWhiteSpace($FileResult.ArchivedPath)
    ) {
        $FileResult.ArchivedPath
    }
    else {
        Get-StringOrDefaultHC $FileResult.Item.FullName ''
    }
    $defaultsPath = Get-StringOrDefaultHC $DefaultsFilePath ''

    # Copy of the processed matrix file in the log folder, written by the
    # END stage and extended with the 'AccessList', 'GroupManagers' and
    # 'AdObjects' sheets. Older runs and partially built file results
    # don't have this property; the row is then skipped.
    $logMatrixPath = if (
        $FileResult.PSObject.Properties.Match('LogMatrixFilePath').Count -and
        -not [string]::IsNullOrWhiteSpace($FileResult.LogMatrixFilePath)
    ) {
        $FileResult.LogMatrixFilePath
    }
    else { '' }

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

    # Run-level diagnostics roll-up, written by the END stage at the root of
    # the dated run folder (one level above this matrix's own folder). Linked
    # from the footer rather than the cards because it covers the whole run,
    # not this one matrix — it is the file to diff between nights.
    $runDiagnosticsPath = if (
        -not [string]::IsNullOrWhiteSpace($FileResult.LogFolder)
    ) {
        $candidate = Join-Path `
            -Path (Split-Path -Path $FileResult.LogFolder -Parent) `
            -ChildPath 'Diagnostics.json'

        if (Test-Path -LiteralPath $candidate) { $candidate } else { '' }
    }
    else { '' }

    # Sortable diagnostics page for the whole run, same run folder.
    $diagnosticsHtmlPath = if (
        -not [string]::IsNullOrWhiteSpace($FileResult.LogFolder)
    ) {
        $candidateHtml = Join-Path `
            -Path (Split-Path -Path $FileResult.LogFolder -Parent) `
            -ChildPath 'Diagnostics.html'

        if (Test-Path -LiteralPath $candidateHtml) { $candidateHtml } else { '' }
    }
    else { '' }

    # Companion field reference, written to the same run folder. Skipped when
    # absent so an older log folder still renders.
    $diagnosticsFieldsPath = if (
        -not [string]::IsNullOrWhiteSpace($FileResult.LogFolder)
    ) {
        $candidateFields = Join-Path `
            -Path (Split-Path -Path $FileResult.LogFolder -Parent) `
            -ChildPath 'Diagnostics.Fields.json'

        if (Test-Path -LiteralPath $candidateFields) { $candidateFields } else { '' }
    }
    else { '' }

    # Each row: (label, value-html, use-mono-font?)
    $items = @(
        @{ Label = 'Matrix log copy'; Value = (Convert-PathToFileLink -Path $logMatrixPath -Title 'Copy of the processed matrix file, including the AccessList, GroupManagers and AdObjects sheets'); Mono = $true }
        @{ Label = 'Diagnostics page'; Value = (Convert-PathToFileLink -Path $diagnosticsHtmlPath -Title 'Sortable, self-contained page holding every diagnostics row for this run — open two runs side by side to compare'); Mono = $true }
        @{ Label = 'Run diagnostics'; Value = (Convert-PathToFileLink -Path $runDiagnosticsPath -Title 'Volume and cost counters for every Settings row in this run — compare the same path across runs to separate data growth from storage slowdown'); Mono = $true }
        @{ Label = 'Diagnostics fields'; Value = (Convert-PathToFileLink -Path $diagnosticsFieldsPath -Title 'What every diagnostics counter means, how to read them together, and the caveats'); Mono = $true }
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
    #
    # margin-top is 12px, not 32px: the last settings card already ends with
    # 12px of its own bottom padding (Build-MatrixDetailCardHC wraps every card
    # in 'padding:0 16px 12px 16px'), so 32px here added up to a 44px gap. 12+12
    # gives 24px, the same rhythm as the gap under the header banner.
    return @"
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; margin-top:12px; table-layout:fixed;">
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
    param(
        [object]$MatrixItem,
        # When the parent matrix file hit a file-level FatalError its settings
        # never executed, so a clean row must not read as green "success" — it
        # is shown as grey "Skipped" instead.
        [bool]$FileHasFatalError = $false
    )

    # Determine card status
    $err = @($MatrixItem.Check | Where-Object Type -EQ 'FatalError').Count
    $warn = @($MatrixItem.Check | Where-Object Type -EQ 'Warning').Count
    $fixed = @($MatrixItem.Check | Where-Object Type -EQ 'Fixed').Count

    # Notices that are neither errors, warnings nor fixes — Type 'Information',
    # and any unknown/future type. Counted with the same rule
    # Build-SettingsRowHC uses for the overview's small blue 'i', so the two
    # views agree: if the overview shows an 'i' for a matrix, opening its
    # execution report must show the matching notice cards.
    $info = @($MatrixItem.Check | Where-Object {
            $_.Type -notin @('FatalError', 'Warning', 'Fixed')
        }).Count

    # Render the full card whenever the row carries ANY check. Info notices
    # used to be excluded here, so a row whose only checks were informational
    # fell through to the compact header-only card and its notices were
    # silently dropped — visible only in the log folder's JSON detail files.
    $hasChecks = ($err + $warn + $fixed + $info) -gt 0

    $isSkipped = $false
    if ($err -gt 0) {
        $accent = $Script:Theme.AccentError
    }
    elseif ($warn -gt 0) {
        $accent = $Script:Theme.AccentWarning
    }
    elseif ($FileHasFatalError) {
        $accent = $Script:Theme.AccentSkipped
        $isSkipped = $true
    }
    else {
        # A row whose only finding is 'Fixed' is green: the run resolved it.
        $accent = $Script:Theme.AccentSuccess
    }
    $statusLabel = Format-IssueCountLabelHC -Errors $err -Warnings $warn -Fixed $fixed

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

    # When the file errored out, a clean row is "Skipped" (grey) rather than
    # green; surface that with an inline tag next to the ComputerName.
    $skippedTag = if ($isSkipped) {
        " &nbsp;<span style='display:inline-block; padding:2px 8px; background-color:$($Script:Theme.AccentSkipped); color:#ffffff; border-radius:10px; font-size:10px; font-weight:700; letter-spacing:0.3px; text-transform:uppercase; vertical-align:middle;'>Skipped</span>"
    }
    else { '' }

    # Two-column, three-row metadata layout. Three columns squeezed into the
    # 55% meta cell left each value roughly 150px, which is less than several
    # of them need: labels split across lines mid-phrase ("APPLY / DEFAULTS")
    # and values broke mid-word at hyphens ("BEL ROL- / AGG-SAGREX"). Trading
    # a column for a row roughly doubles the width available to every cell
    # without widening the card or shrinking the ComputerName/Path block.
    #
    # Ordering runs most-identifying first: what the permissions apply to
    # (Group, Site), then what was done to it (Action, Apply Defaults), then
    # the run's bookkeeping (ID, Duration).
    #
    #   Col 1                Col 2
    #   ------------------   ------------------
    #   GROUP: x             SITE: x
    #   ACTION: x            APPLY DEFAULTS: x
    #   ID: x                [clock] Duration
    #
    # Duration keeps an inline SVG clock icon (universally readable as "time")
    # in place of a text label. Everything else uses inline "LABEL: value"
    # styling. Column positions are reserved (with &nbsp; fallbacks for
    # missing optional fields) so cells align vertically down the grid.

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
            [string]$TitleAttr = '',
            # Raw HTML appended after the value span, inside the same cell.
            # Used for the Diagnostics chip so it can sit beside the duration
            # without claiming a grid column of its own.
            [string]$TrailingHtml = ''
        )
        $valueStyle = if ($Mono) { "font-family:$($Script:Theme.MonoStack); font-size:11px;" } else { 'font-size:12px;' }
        $titleHtml = if ($TitleAttr) { " title=`"$TitleAttr`"" } else { '' }
        $valueHtml = "<span$titleHtml style='color:$($Script:Theme.TextMuted); $valueStyle'>$Value</span>"
        $trailing = if ($TrailingHtml) { "&nbsp;&nbsp;$TrailingHtml" } else { '' }
        return "<td valign='middle' width='50%' style='padding:5px 20px 5px 0; white-space:nowrap;'>$IconHtml$valueHtml$trailing</td>"
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
        # The label is kept non-breaking on its own: the browser stylesheet
        # relaxes white-space on these cells so long values can wrap instead
        # of clipping, and without this a two-word label like "Apply
        # Defaults" would be a candidate for that wrap too. Values may wrap,
        # labels never do.
        $labelHtml = "<span style='font-size:10px; font-weight:700; color:$($Script:Theme.TextLight); text-transform:uppercase; letter-spacing:0.5px; margin-right:6px; white-space:nowrap;'>$Label`:</span>"
        $valueHtml = "<span$titleHtml style='color:$($Script:Theme.TextMuted); $valueStyle'>$Value</span>"
        return "<td valign='middle' width='50%' style='padding:5px 20px 5px 0; white-space:nowrap;'>$labelHtml$valueHtml</td>"
    }

    # Reserves a column position when an optional field is absent, so the two
    # columns stay aligned down the grid instead of collapsing.
    $emptyMetaCell = "<td valign='middle' width='50%' style='padding:5px 20px 5px 0;'>&nbsp;</td>"

    # Diagnostics link — a quiet outlined chip rather than a labeled metadata
    # cell. The counters behind it are only interesting when someone is
    # actively investigating a slow path, so putting the numbers themselves in
    # the card would tax every reader to serve the rare one. A chip next to
    # the duration is discoverable exactly when you are already looking at the
    # time, and costs one line of pixels.
    #
    # Rendered only when the END stage actually wrote the file (rows that
    # never executed have no telemetry), so this never links into the void.
    $diagFile = Get-StringOrDefaultHC $MatrixItem.DiagnosticsFileName ''
    $diagHtml = if ($diagFile) {
        $diagHref = [System.Net.WebUtility]::HtmlEncode($diagFile)

        # A '?' beside the chip, linking to the field reference in the run
        # folder one level up. The counters are useless to anyone who does not
        # know what they mean, and the answer should be one click away from the
        # numbers rather than filed somewhere the reader has to go looking for.
        # Rendered as a separate small link so the chip itself still goes
        # straight to the data for readers who already know the fields.
        "<a href='$diagHref' target='_blank' rel='noopener noreferrer' title='Volume and cost counters for this path (JSON)' style='display:inline-block; padding:1px 8px; border:1px solid $($Script:Theme.BorderMain); border-radius:10px; font-size:10px; font-weight:700; letter-spacing:0.3px; text-transform:uppercase; color:$($Script:Theme.TextLight); text-decoration:none; vertical-align:middle;'>Diagnostics</a>" +
        "<a href='../Diagnostics.Fields.json' target='_blank' rel='noopener noreferrer' title='What do these counters mean?' style='display:inline-block; margin-left:4px; width:15px; height:15px; line-height:15px; text-align:center; border:1px solid $($Script:Theme.BorderMain); border-radius:50%; font-size:10px; font-weight:700; color:$($Script:Theme.TextLight); text-decoration:none; vertical-align:middle;'>?</a>"
    }
    else { '' }

    $metaRows = @()

    # Row 1: Group | Site — both optional. When neither is present the row is
    # dropped entirely rather than rendered as two blank cells, so matrices
    # without them do not carry an empty band of whitespace.
    if ($groupName -or $siteCode) {
        $row1Cells = @(
            $(if ($groupName) { New-InlineMetaCellHtml -Label 'Group' -Value $groupName } else { $emptyMetaCell })
            $(if ($siteCode) { New-InlineMetaCellHtml -Label 'Site' -Value $siteCode } else { $emptyMetaCell })
        )
        $metaRows += "<tr>$($row1Cells -join '')</tr>"
    }

    # Row 2: Action | Apply Defaults — always present.
    $row2Cells = @(
        (New-InlineMetaCellHtml -Label 'Action' -Value $action)
        (New-InlineMetaCellHtml -Label 'Apply Defaults' -Value $applyDefaultStr)
    )
    $metaRows += "<tr>$($row2Cells -join '')</tr>"

    # Row 3: ID | Duration — the Diagnostics chip shares the Duration cell
    # rather than claiming a column of its own, so the grid is identical
    # whether or not diagnostics were written.
    $row3Cells = @(
        (New-InlineMetaCellHtml -Label 'ID' -Value $idShortHtml -Mono $true -TitleAttr $idFullHtml)
        (New-IconMetaCellHtml -IconHtml $iconDuration -Value $dur -Mono $true -TrailingHtml $diagHtml)
    )
    $metaRows += "<tr>$($row3Cells -join '')</tr>"

    # width:100% lets the two columns split the meta cell evenly instead of
    # shrink-wrapping to their content, which is what actually hands the space
    # back to the values.
    $metadataTable = "<table role='presentation' cellpadding='0' cellspacing='0' border='0' width='100%' style='border-collapse:collapse; width:100%;'>" +
    ($metaRows -join '') +
    '</table>'

    # Three-column horizontal header — no visible dividers, just consistent
    # padding. table-layout:fixed plus an explicit 55% width on the metadata
    # column gives the metadata pairs (Group/Site, Action/Apply Defaults,
    # ID/Duration) enough room to fit on one line each at the report's 900px
    # design width, and pushes the long monospace path to wrap onto its own
    # line sooner — leaving more breathing room overall instead of forcing
    # the whole card past the viewport edge.
    $headerBlock = @"
<table role="presentation" class="rr-settings-head" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; table-layout:fixed;">
    <tr>
        <td class="rr-icon-cell" valign='middle' width='40' style='padding:14px 8px 14px 10px;'>$dotHtml</td>
        <td class="rr-content-cell" valign='middle' style='padding:14px 16px 14px 0;'>
            <div style='font-size:14px; font-weight:700; color:$($Script:Theme.TextMain); line-height:1.25;'>$comp$skippedTag</div>
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

    # Tally for header status pill — shared with the email card so the report
    # header and the overview card agree on a file's status.
    $tally = Get-FileCheckTallyHC -FileResult $FileResult
    $fileErrs = $tally.Errors
    $fileWarns = $tally.Warnings
    $fileFixed = $tally.Fixed

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

    $hdrLabel = Format-IssueCountLabelHC -Errors $fileErrs -Warnings $fileWarns -Fixed $fileFixed

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

    # A file-level FatalError (e.g. 'Runspace processing failed') means the
    # settings never executed; their rows render as "Skipped" instead of green.
    $fileHasFatalError = $false
    foreach ($g in $fileLevelGroups) {
        if ($g.Checks -and @($g.Checks | Where-Object Type -EQ 'FatalError').Count -gt 0) {
            $fileHasFatalError = $true
            break
        }
    }

    if ($fileLevelCount -gt 0) {
        $issueRows = ''
        foreach ($g in $fileLevelGroups) {
            if ($g.Checks) {
                foreach ($c in $g.Checks) {
                    # Standalone report: include the 16px inset wrapper so File Issues rows
                    # have the same indented look as the Settings rows below them.
                    # LinkJsonDetail: the detail JSON (written only when the
                    # check has a 'Value') sits next to this report, so the
                    # check name links to it — same as matrix-level checks.
                    $issueRows += Build-FileLevelCheckRowHC `
                        -Check $c `
                        -SheetLabel $g.Label `
                        -LinkJsonDetail $true
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
            $matrixRowsHtml += Build-MatrixDetailCardHC -MatrixItem $m -FileHasFatalError $fileHasFatalError
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
            /* The desktop grid pins each cell to 50% so the two columns split
               the meta cell evenly. Here the cells flow as chips instead, so
               they size to their content and two short ones can share a line
               that a 50% floor would have wasted. */
            width: auto !important;
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
                <!-- File header. No bottom padding: whatever follows (the
                     "File Issues" or "Settings" h2, or the fallback "no
                     settings" paragraph) brings its own top margin, and
                     stacking the two produced a 48px band of empty page
                     colour under the banner. -->
                <tr>
                    <td style="padding:0;">
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
function Write-RunDiagnosticsHtmlHC {
    <#
    .SYNOPSIS
        Writes 'Diagnostics.html': one self-contained, sortable page holding
        every diagnostics row for the whole run.

    .DESCRIPTION
        Built for SIDE-BY-SIDE COMPARISON of two nights. Open yesterday's file in
        one window and today's in another, apply the same sort and filter to
        both, and read the differences off the screen.

        SELF-CONTAINED IS A HARD REQUIREMENT, NOT A PREFERENCE
        The data is embedded in the page as JSON rather than fetched from the
        sibling .json files. A page opened over file:// cannot fetch a local
        file — the browser blocks it as a cross-origin request — so a page that
        loaded its data at runtime would simply render empty tables from a log
        share. Embedding also means a single file can be mailed or copied out of
        the run folder and still work.

        NO EXTERNAL ASSETS
        No CDN, no web fonts, no frameworks. Log shares are often reached from
        machines with no internet route, and a diagnostics page that needs the
        network to render its own sort arrows is not a diagnostics page.

        DEFAULT SORT IS STABLE, NOT INTERESTING
        Rows default to computer then path, NOT to cost descending. Cost order
        differs between nights, so a cost-sorted pair of windows would show the
        same folder on different screen rows and defeat the whole purpose. The
        stable default makes two files line up; one click gets cost order when
        that is what is wanted.

    .NOTES
        Failures are swallowed, as with the other diagnostics writers.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]]$Matrices,
        [Parameter(Mandatory)] [string]$LogFolder,
        [Parameter()] [datetime]$RunStartTime
    )

    try {
        #region Build the two row sets
        $settingRows = [System.Collections.Generic.List[object]]::new()
        $pathRows = [System.Collections.Generic.List[object]]::new()

        foreach ($m in $Matrices) {
            if (-not $m.Telemetry) { continue }

            $duration = if ($m.JobTime.Duration) {
                [math]::Round($m.JobTime.Duration.TotalSeconds, 1)
            }
            else { $null }

            $settingRow = [ordered]@{
                MatrixFile = $m.FileContext.Item.Name
                Duration_s = $duration
            }
            foreach ($t in $m.Telemetry.GetEnumerator()) {
                if ($t.Key -eq 'Paths') { continue }
                $settingRow[$t.Key] = $t.Value
            }
            $settingRows.Add([PSCustomObject]$settingRow)

            foreach ($p in $m.Telemetry.Paths) {
                $pathRow = [ordered]@{
                    MatrixFile   = $m.FileContext.Item.Name
                    ComputerName = $m.Telemetry.ComputerName
                    SettingPath  = $m.Telemetry.Path
                }
                foreach ($field in $p.GetEnumerator()) {
                    $pathRow[$field.Key] = $field.Value
                }
                $pathRows.Add([PSCustomObject]$pathRow)
            }
        }

        if ($settingRows.Count -eq 0) { return }
        #endregion

        #region Embed as JSON
        # '<' is escaped so a path containing '</script>' cannot break out of the
        # script block. Depth 5 is ample for flat rows and keeps the file small.
        $encode = {
            param($Rows)
            $json = if ($Rows.Count -eq 0) { '[]' }
            else { @($Rows) | ConvertTo-Json -Depth 5 -Compress -AsArray }
            return ($json -replace '<', '\u003c' -replace '>', '\u003e' -replace '&', '\u0026')
        }

        $settingsJson = & $encode $settingRows
        $pathsJson = & $encode $pathRows
        #endregion

        $runLabel = if ($RunStartTime) {
            $RunStartTime.ToString('yyyy-MM-dd HH:mm:ss')
        }
        else { Split-Path -Path $LogFolder -Leaf }

        $folderLabel = [System.Net.WebUtility]::HtmlEncode((Split-Path -Path $LogFolder -Leaf))
        $generated = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')

        $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Diagnostics $folderLabel</title>
<style>
  :root { color-scheme: light; }
  body { margin:0; padding:16px; background:$($Script:Theme.BgPage);
         font-family:$($Script:Theme.FontStack); color:$($Script:Theme.TextMain); font-size:13px; }
  header { background:$($Script:Theme.BgWhite); border:1px solid $($Script:Theme.BorderMain);
           border-radius:8px; padding:12px 16px; margin-bottom:12px; }
  h1 { margin:0 0 4px 0; font-size:16px; }
  .run { font-family:$($Script:Theme.MonoStack); font-size:14px; font-weight:700;
         color:$($Script:Theme.AccentInfo); }
  .meta { color:$($Script:Theme.TextLight); font-size:12px; }
  .controls { display:flex; flex-wrap:wrap; gap:8px; align-items:center; margin:10px 0 0 0; }
  .controls input[type=text] { padding:5px 8px; border:1px solid $($Script:Theme.BorderMain);
       border-radius:4px; font-family:$($Script:Theme.MonoStack); font-size:12px; min-width:280px; }
  button { padding:5px 10px; border:1px solid $($Script:Theme.BorderMain);
           background:$($Script:Theme.BgAlt); border-radius:4px; cursor:pointer; font-size:12px; }
  button.active { background:$($Script:Theme.AccentInfo); color:#fff;
                  border-color:$($Script:Theme.AccentInfo); }
  .wrap { background:$($Script:Theme.BgWhite); border:1px solid $($Script:Theme.BorderMain);
          border-radius:8px; overflow:auto; max-height:78vh; }
  table { border-collapse:collapse; width:100%; }
  th, td { padding:4px 8px; border-bottom:1px solid $($Script:Theme.BorderLight);
           white-space:nowrap; text-align:left; }
  th { position:sticky; top:0; background:$($Script:Theme.BgAlt); cursor:pointer;
       font-size:11px; text-transform:uppercase; letter-spacing:0.3px;
       color:$($Script:Theme.TextMuted); border-bottom:2px solid $($Script:Theme.BorderMain);
       user-select:none; }
  th:hover { background:$($Script:Theme.BorderLight); }
  th .arrow { color:$($Script:Theme.AccentInfo); font-weight:700; }
  /* Right-aligned + monospace + tabular figures: with a fixed decimal count
     this puts every decimal point on the same vertical line down a column. */
  td.num { text-align:right; font-family:$($Script:Theme.MonoStack);
           font-variant-numeric:tabular-nums; }
  td.path { font-family:$($Script:Theme.MonoStack); font-size:11px; }
  tbody tr:nth-child(even) { background:$($Script:Theme.BgAlt); }
  tbody tr:hover { background:$($Script:Theme.StatusSkipped); }
  .warnbasis { color:$($Script:Theme.AccentWarning); font-weight:700; }
  .zero { color:$($Script:Theme.TextLight); }
  .count { color:$($Script:Theme.TextLight); font-size:12px; }
  .hint { margin-top:10px; color:$($Script:Theme.TextLight); font-size:11px; line-height:1.5; }
</style>
</head>
<body>
<header>
  <h1>Permission matrix diagnostics</h1>
  <div class="run">$([System.Net.WebUtility]::HtmlEncode($runLabel))</div>
  <div class="meta">$folderLabel &nbsp;&middot;&nbsp; page generated $generated</div>
  <div class="controls">
    <button id="btnSettings" class="active" onclick="showGrain('settings')">Per Settings row</button>
    <button id="btnPaths" onclick="showGrain('paths')">Per matrix folder</button>
    <input type="text" id="filter" placeholder="filter any column&hellip;" oninput="render()">
    <button onclick="resetView()">Reset sort &amp; filter</button>
    <span class="count" id="count"></span>
  </div>
  <div class="hint">
    Click a column to sort. Rows start in a <b>stable order</b> (computer, then path) so that two
    runs opened side by side line up row for row &mdash; sort by cost only once you know which
    row you are chasing. Everything is embedded in this file: no other log file is needed.
  </div>
</header>

<div class="wrap"><table id="grid"><thead></thead><tbody></tbody></table></div>

<script>
// JSON is valid JavaScript, so the payload is embedded as an object literal
// rather than as a quoted string handed to JSON.parse. Wrapping it in a JS
// string would mean re-escaping every backslash in every Windows path, which is
// exactly the kind of double-escaping that silently corrupts one row in ten
// thousand. Direct embedding needs no escaping at all beyond the < > & already
// neutralised server-side so no path can close this script block.
const DATA = {
  settings: $settingsJson,
  paths: $pathsJson
};

// Stable default: computer, then the path being described. Cost order differs
// between nights, so it must never be the default - see the function notes.
const STABLE_KEYS = ['ComputerName', 'SettingPath', 'Path', 'MatrixFile'];

let grain = 'settings';
let sortCol = null;
let sortAsc = true;

function columns() {
  const rows = DATA[grain];
  return rows.length ? Object.keys(rows[0]) : [];
}

// Columns whose values are decimals get a FIXED number of decimal places, so
// the digits line up in the column and two runs opened side by side can be
// compared by eye. A raw JSON number prints 0.6 next to 0.61 next to 1, which
// puts the decimal point in a different place on every row and makes a column
// of timings unreadable.
//
// The decision is made per column from the DATA itself rather than from a
// hard-coded field list: any column where at least one value is a
// non-integer number is treated as decimal, so a new counter formats
// correctly without anyone remembering to register it here.
const DECIMALS = 2;
const decimalCols = {};

function computeDecimalCols() {
  for (const g of ['settings', 'paths']) {
    decimalCols[g] = {};
    for (const row of DATA[g]) {
      for (const k in row) {
        const v = row[k];
        if (typeof v === 'number' && !Number.isInteger(v)) decimalCols[g][k] = true;
      }
    }
  }
}
computeDecimalCols();

function formatCell(v, col) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') {
    if (decimalCols[grain][col]) return v.toFixed(DECIMALS);
    // Thousands separators on whole numbers, so 5512094 is readable at a
    // glance. Uses a fixed locale so two machines produce identical output.
    return v.toLocaleString('en-US');
  }
  return String(v);
}

function stableCompare(a, b) {
  for (const k of STABLE_KEYS) {
    if (a[k] === undefined || b[k] === undefined) continue;
    const r = String(a[k]).localeCompare(String(b[k]));
    if (r !== 0) return r;
  }
  return 0;
}

function compare(a, b, col) {
  const x = a[col], y = b[col];
  const xn = (x === null || x === '' || x === undefined), yn = (y === null || y === '' || y === undefined);
  if (xn && yn) return 0;
  if (xn) return 1;   // blanks always last, whichever direction
  if (yn) return -1;
  if (typeof x === 'number' && typeof y === 'number') return x - y;
  if (typeof x === 'boolean' && typeof y === 'boolean') return (x ? 1 : 0) - (y ? 1 : 0);
  const nx = Number(x), ny = Number(y);
  if (!isNaN(nx) && !isNaN(ny) && x !== '' && y !== '') return nx - ny;
  return String(x).localeCompare(String(y));
}

function render() {
  const cols = columns();
  const term = document.getElementById('filter').value.toLowerCase();

  let rows = DATA[grain].slice();
  if (term) {
    rows = rows.filter(r => cols.some(c => String(r[c] === null ? '' : r[c]).toLowerCase().includes(term)));
  }

  if (sortCol === null) {
    rows.sort(stableCompare);
  } else {
    // Ties fall back to the stable order, so a sort never scrambles equal rows
    // differently in two files.
    rows.sort((a, b) => {
      const c = compare(a, b, sortCol);
      return (c !== 0) ? (sortAsc ? c : -c) : stableCompare(a, b);
    });
  }

  const thead = document.querySelector('#grid thead');
  thead.innerHTML = '<tr>' + cols.map(c => {
    const arrow = (c === sortCol) ? ' <span class="arrow">' + (sortAsc ? '\u25B2' : '\u25BC') + '</span>' : '';
    return '<th onclick="sortBy(\'' + c + '\')" title="' + c + '">' + c + arrow + '</th>';
  }).join('') + '</tr>';

  const tbody = document.querySelector('#grid tbody');
  tbody.innerHTML = rows.map(r => '<tr>' + cols.map(c => {
    const raw = r[c];
    let cls = '';
    if (typeof raw === 'number') { cls = 'num'; if (raw === 0) cls += ' zero'; }
    else if (c === 'Path' || c === 'SettingPath') cls = 'path';
    else if (c === 'AclReadBasis' && (raw === 'warmup+stride' || raw === 'none')) cls = 'warnbasis';
    const text = formatCell(raw, c).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    return '<td class="' + cls + '">' + text + '</td>';
  }).join('') + '</tr>').join('');

  document.getElementById('count').textContent =
    rows.length + ' of ' + DATA[grain].length + ' rows' + (term ? ' (filtered)' : '');
}

function sortBy(col) {
  if (sortCol === col) { sortAsc = !sortAsc; } else { sortCol = col; sortAsc = false; }
  render();
}

function showGrain(g) {
  grain = g;
  sortCol = null; sortAsc = true;
  document.getElementById('btnSettings').className = (g === 'settings') ? 'active' : '';
  document.getElementById('btnPaths').className = (g === 'paths') ? 'active' : '';
  render();
}

function resetView() {
  sortCol = null; sortAsc = true;
  document.getElementById('filter').value = '';
  render();
}

render();
</script>
</body>
</html>
"@

        $html | Out-File `
            -FilePath (Join-Path -Path $LogFolder -ChildPath 'Diagnostics.html') `
            -Encoding UTF8 -Force
    }
    catch {
        Write-Verbose "Failed writing the diagnostics HTML page: $_"
    }
}