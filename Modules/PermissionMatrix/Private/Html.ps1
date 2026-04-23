<#
    Html.ps1
    Consolidated HTML rendering logic for Toolbox.PermissionMatrixHC
#>

# =====================================================================
# GLOBAL HTML THEME
# Centralized color palette used by all HTML generation functions
# =====================================================================
$Script:Theme = @{
    # Status Colors (Cards & Table Rows)
    StatusError    = '#fee2e2' # Light Red
    StatusWarning  = '#ffedd5' # Light Orange
    StatusSuccess  = '#dcfce7' # Light Green
    
    # Text Colors
    TextMain       = '#111827' # Dark Gray/Black
    TextMuted      = '#374151' # Medium Gray
    TextLight      = '#6b7280' # Light Gray
    
    # Backgrounds & Borders
    BgWhite        = '#ffffff'
    BgAlt          = '#f9fafb' # Off-white for "About" sections
    BorderMain     = '#d1d5db'
    BorderLight    = '#e5e7eb'
    
    # Links
    LinkColor      = '#2563eb' # Standard Blue
    LinkHoverColor = '#1d4ed8' # Darker Blue for hover effect
}

function Initialize-HtmlStructureHC {

    $style = @"
<style type="text/css">
    a { color: $($Script:Theme.LinkColor); text-decoration: underline; }
    a:hover { color: $($Script:Theme.LinkHoverColor); }
    body { font-family:verdana; font-size:14px; background-color:white; }
    h1, h2, h3 { margin-bottom: 0; }
    p.italic { font-style: italic; font-size: 12px; }
    table { border-collapse:collapse; padding:3px; }
    td, th { border:1px none; padding:3px; }
    .matrixTable { border: 1px solid $($Script:Theme.BorderMain); border-spacing: 0; width: 600px; }
    .matrixTitle { background-color:lightgrey; text-align:center; padding:6px; }
    .matrixHeader { letter-spacing:5pt; font-style:italic; }
    .matrixFileInfo { font-size:12px; font-style:italic; text-align:center; }
    .legendTable { border:1px solid black; table-layout:fixed; }
    .legendTable td { text-align:center; }
    .probTypeError { background-color: $($Script:Theme.StatusError); }
    .probTypeWarning { background-color: $($Script:Theme.StatusWarning); }
    .probTypeInfo { background-color: $($Script:Theme.BgAlt); }
    .probTextError { color:#fee2e2; font-weight:bold; }
    .probTextWarning { color:#ffedd5; font-weight:bold; }
    .aboutTable th, .aboutTable td { color:#8f8c8c; }
    base { target="_blank" }
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
            SettingsHeader =
            '<th class="matrixHeader" colspan="8">Settings</th>' +
            '<tr><td></td><td>ID</td><td>ComputerName</td><td>Path</td><td>Action</td><td>Duration</td></tr>'

            LegendTable    =
            '<table class="legendTable"><tr>' +
            '<td class="probTypeError">Error</td>' +
            '<td class="probTypeWarning">Warning</td>' +
            '<td class="probTypeInfo">Information</td>' +
            '</tr></table>'
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

function New-HtmlCheckRowHC {
    param(
        [object]$CheckItem
    )

    $cls = Get-HtmlClassProbTypeHC $CheckItem.Type
    $name = [System.Net.WebUtility]::HtmlEncode($CheckItem.Name)
    $desc = [System.Net.WebUtility]::HtmlEncode($CheckItem.Description)
    
    # Grab the JSON filename we just injected!
    if ($CheckItem.JsonFileName) {
        $nameHtml = "<a href='$($CheckItem.JsonFileName)' style='color: $($Script:Theme.TextMain); text-decoration: underline;' title='View JSON Details'>$name</a>"
    }
    else {
        $nameHtml = $name
    }

    return @"
    <tr class='$cls' style='border-bottom: 1px solid $($Script:Theme.BorderLight);'>
        <td style='width: 30%; padding: 8px 6px; font-weight: 600;'>$nameHtml</td>
        <td style='padding: 8px 6px;'>$desc</td>
    </tr>
"@
}

function New-HtmlSectionHC {
    param(
        [string]$Title,
        [object]$Checks
    )

    if (-not $Checks) { return '' }

    $rows = $Checks |
    ConvertTo-StructuredObjectHC |
    ForEach-Object { New-HtmlCheckRowHC $_ }

    "<tr><th class='matrixHeader' colspan='8'>$Title</th></tr>$($rows -join '')"
}

function New-SettingsCardHtmlHC {
    param(
        [object]$MatrixItem
    )

    # =====================================================================
    # CSS STYLE DICTIONARY 
    # Edit these strings to quickly change the look of the card elements!
    # =====================================================================
    $css = @{
        CardOuter    = "border: 1px solid $($Script:Theme.BorderMain); border-radius: 8px; margin-bottom: 25px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); page-break-inside: avoid; overflow: hidden; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;"
        CardHeader   = 'padding: 12px 16px; border-bottom: 1px solid #d1d5db; display: flex; justify-content: space-between; align-items: center; gap: 20px;'
        HeaderLeft   = 'font-size: 15px;'
        HeaderRight  = 'font-size: 13px; font-weight: 700; color: #111827;'
        PillComp     = 'background-color: rgba(255,255,255,0.6); border: 1px solid rgba(0,0,0,0.1); padding: 3px 12px; border-radius: 12px; font-size: 13px; font-weight: 700; margin-right: 10px; color: #1f2937;'
        PathText     = 'font-family: Consolas, monospace; font-size: 13.5px; color: #374151; background-color: rgba(255,255,255,0.4); padding: 2px 6px; border-radius: 4px;'
        
        AboutOuter   = "padding: 12px 16px; background-color: $($Script:Theme.BgAlt); border-bottom: 1px solid $($Script:Theme.BorderLight);"
        AboutTitle   = 'margin-top:0; margin-bottom:0px; font-size: 14px; color: #374151;'
        AboutTable   = 'border:none; font-size:13px; border-collapse: separate; border-spacing: 0 6px;'
        AboutLabel   = "width:100px; font-weight:600; color:$($Script:Theme.TextLight);"
        AboutVal     = 'color: #111827;'
        AboutIdVal   = 'color: #111827; font-family: Consolas, monospace; font-size: 12px;'
        
        CheckOuter   = 'padding: 16px; background-color: #ffffff;'
        CheckTitle   = 'margin:0 0 8px 0; color: #374151;'
        CheckTable   = 'width:100%; border-collapse: collapse; font-size: 13px; background-color: white; border: 1px solid #d1d5db; border-radius: 4px; overflow: hidden;'
        CheckRow     = "border-bottom: 1px solid $($Script:Theme.BorderLight);"
        CheckLinkTd = 'font-weight: 600; white-space: nowrap; width: 1%; padding: 8px 15px 8px 6px; text-align: left;'
        CheckDesc   = 'padding: 8px 6px; color: #374151; text-align: left; width: auto;'
        CheckBadgeTd = 'width: 80px; text-align: right; padding: 8px 16px 8px 6px; font-weight: 700; font-size: 12px; letter-spacing: 0.5px;'
        SuccessText  = 'padding-top:5px; font-style:italic; color: #6b7280; margin: 0;'
    }

    #region Get ID, ComputerName, Path, GroupName, SiteCode
    $fullId = if (-not [string]::IsNullOrWhiteSpace($MatrixItem.Setting.Raw.ID)) { 
        $MatrixItem.Setting.Raw.ID 
    }
    elseif ($MatrixItem.ID) { 
        $MatrixItem.ID 
    }
    else { 
        'N/A' 
    }
        
    $comp = [System.Net.WebUtility]::HtmlEncode($MatrixItem.Setting.Raw.ComputerName)
    $path = [System.Net.WebUtility]::HtmlEncode($MatrixItem.Setting.Raw.Path)
    $group = [System.Net.WebUtility]::HtmlEncode($MatrixItem.Setting.Raw.GroupName)
    $site = [System.Net.WebUtility]::HtmlEncode($MatrixItem.Setting.Raw.SiteCode)
    #endregion

    #region Calculate Status & Colors
    $errCount = @($MatrixItem.Check | Where-Object Type -EQ 'FatalError').Count
    $warnCount = @($MatrixItem.Check | Where-Object Type -EQ 'Warning').Count
        
    $statusSymbol = $null
    $statusText = $null

    if ($errCount -gt 0) {
        $headerColor = $Script:Theme.StatusError
        
        $warnText = if ($warnCount -gt 0) {
            ", $warnCount $(Plural $warnCount 'Warning')"
        }
        
        $statusText = "Failed ($errCount $(Plural $errCount 'Error')$warnText)"
        $statusSymbol = '✖'
    }
    elseif ($warnCount -gt 0) {
        $headerColor = $Script:Theme.StatusWarning
        
        $statusText = "Completed with $warnCount $(Plural $warnCount 'Warning')"
        $statusSymbol = '⚠'
    }
    else {
        $headerColor = $Script:Theme.StatusSuccess
        
        $statusText = 'Success'
        $statusSymbol = '✓'
    }
    #endregion

    #region Calculate Duration String
    if ($MatrixItem.JobTime.Start -and $MatrixItem.JobTime.End) {
        $ts = New-TimeSpan -Start $MatrixItem.JobTime.Start -End $MatrixItem.JobTime.End
        $durVal = if ($ts.TotalMinutes -ge 1) { "$([math]::Round($ts.TotalMinutes)) min" } else { "$([math]::Round($ts.TotalSeconds)) sec" }
        $startStr = $MatrixItem.JobTime.Start.ToString('HH:mm')
        $endStr = $MatrixItem.JobTime.End.ToString('HH:mm')
        
        $timeStr = "$durVal ($startStr - $endStr)"
    }
    else {
        $timeStr = 'N/A'
    }
    #endregion

    #region Create Check Table or Success Message
    $checkTable = ''
    if ($MatrixItem.Check -and $MatrixItem.Check.Count -gt 0) {
        $checkRows = ''
        foreach ($c in $MatrixItem.Check) {
            $cls = Get-HtmlClassProbTypeHC $c.Type
            $name = [System.Net.WebUtility]::HtmlEncode($c.Name)
            $desc = [System.Net.WebUtility]::HtmlEncode($c.Description)
            
            # Determine Icon and Badge Text dynamically based on error type
            if ($c.Type -eq 'FatalError') {
                $rowIcon = '✖'
                $rowBadge = 'ERROR'
                $rowColor = '#b91c1c' # Deeper red for readability
            }
            elseif ($c.Type -eq 'Warning') {
                $rowIcon = '⚠'
                $rowBadge = 'WARNING'
                $rowColor = '#c2410c' # Deeper orange for readability
            }
            else {
                $rowIcon = 'ℹ'
                $rowBadge = 'INFO'
                $rowColor = '#1d4ed8' # Standard blue
            }
            
            $linkTarget = if ($c.JsonFileName) { $c.JsonFileName } else { '#' }
            
            $checkRows += @"
            <tr class='$cls' style='$($css.CheckRow)'>
                <td style='width: 30px; text-align: center; font-weight: bold; color: $rowColor; font-size: 14px;'>$rowIcon</td>
                <td style='$($css.CheckLinkTd)'>
                    <a href="$linkTarget" style="$($css.CheckLink)" title="Click to view full JSON details">
                        $name
                    </a>
                </td>
                <td style='$($css.CheckDesc)'>$desc</td>
                <td style='$($css.CheckBadgeTd) color: $rowColor;'>$rowBadge</td>
            </tr>
"@
        }
        
        $checkTable = @"
        <h4 style="$($css.CheckTitle)">Detailed Results</h4>
        <table style='$($css.CheckTable)'>
            $checkRows
        </table>
"@
    }
    else {
        $checkTable = "<p style='$($css.SuccessText)'>No issues detected. Execution successful.</p>"
    }
    #endregion

    return @"
<div style="$($css.CardOuter)">
    <div style="background-color: $headerColor; $($css.CardHeader)">
        <div style="$($css.HeaderLeft)">
            $statusSymbol
            <span style="$($css.PillComp)">$comp</span>
            <span style="$($css.PathText)">$path</span>
        </div>
        <div style="$($css.HeaderRight)">
            $statusText
        </div>
    </div>
    <div style="$($css.AboutOuter)">
        <h3 style="$($css.AboutTitle)">About</h3>
        <table style="$($css.AboutTable)">
            <tr><td style="$($css.AboutLabel)">ID:</td><td style="$($css.AboutIdVal)">$fullId</td></tr>
            <tr><td style="$($css.AboutLabel)">GroupName:</td><td style="$($css.AboutVal)">$group</td></tr>
            <tr><td style="$($css.AboutLabel)">SiteCode:</td><td style="$($css.AboutVal)">$site</td></tr>
            <tr><td style="$($css.AboutLabel)">Duration:</td><td style="$($css.AboutVal)">$timeStr</td></tr>
        </table>
    </div>
    <div style="$($css.CheckOuter)">
        $checkTable
    </div>
</div>
"@
}

function Write-MatrixExecutionReportHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$FileMatrices, 
        [Parameter(Mandatory)][hashtable]$Html,
        [Parameter(Mandatory)][string]$LogFolder
    )

    if (-not (Test-Path -LiteralPath $LogFolder -PathType Container)) {
        return $null 
    }

    $firstMatrix = $FileMatrices[0]

    $modBy = [System.Net.WebUtility]::HtmlEncode($firstMatrix.FileContext.ExcelInfo.LastModifiedBy ?? 'Unknown')

    $modDt = if ($firstMatrix.FileContext.ExcelInfo.Modified -is [datetime]) {
        $firstMatrix.FileContext.ExcelInfo.Modified.ToString('dd/MM/yyyy HH:mm:ss')
    }
    else { 'Unknown' }

    # 1. Gather all the file-level sections
    $fileSections = @(
        New-HtmlSectionHC 'File Checks' $firstMatrix.FileContext.Check
        if ($firstMatrix.FileContext.Sheets.FormData.Check) {
            New-HtmlSectionHC 'FormData Checks' $firstMatrix.FileContext.Sheets.FormData.Check
        }
        if ($firstMatrix.FileContext.Sheets.Permissions.Check) {
            New-HtmlSectionHC 'Permissions Checks' $firstMatrix.FileContext.Sheets.Permissions.Check
        }
    ) -join ''

    if ([string]::IsNullOrWhiteSpace($fileSections)) {
        $globalFileTableHtml = @"
<table class="matrixTable" style="width:100%; margin-bottom: 25px; background-color: $($Script:Theme.StatusSuccess); border: 1px solid $($Script:Theme.BorderLight); border-radius: 6px;">
    <tr>
        <td style="padding: 12px; font-weight: 600; color: $($Script:Theme.TextMain); width: 30%;">✓ Validation Passed</td>
        <td style="padding: 12px; color: $($Script:Theme.TextMuted);">No global file issues detected. All required sheets and data formats are valid.</td>
    </tr>
</table>
"@
    } 
    else {
        # Errors found! Wrap them in your standard table
        $globalFileTableHtml = @"
<table class="matrixTable" style="width:100%; margin-bottom: 25px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
$fileSections
</table>
"@
    }

    $settingsSections = ''
    foreach ($matrix in ($FileMatrices | Sort-Object { $_.Setting.Raw.ComputerName }, { $_.Setting.Raw.Path }, { $_.ID })) {
        $settingsSections += New-SettingsCardHtmlHC -MatrixItem $matrix
    }

    # 3. Inject $globalFileTableHtml instead of wrapping $fileSections directly
    $htmlOut = @"
<!DOCTYPE html>
<html><head>
$($Html.Style)
$($Html.TroubleshootingStyle)
</head><body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #111827;">
<h1 style="margin-bottom: 5px;">Execution & Troubleshooting Report</h1>
<h2 style="margin-top: 5px; color: #374151;">File: $($firstMatrix.FileContext.Item.Name)</h2>
<p class="matrixFileInfo" style="text-align:left; margin-top:5px; margin-bottom:25px; color: #6b7280; font-style: italic;">
    Last change: $modBy @ $modDt
</p>

<h3 style="border-bottom: 2px solid #e5e7eb; padding-bottom: 5px; margin-bottom: 15px;">Global File Status</h3>
$globalFileTableHtml

<h3 style="border-bottom: 2px solid #e5e7eb; padding-bottom: 5px; margin-bottom: 25px;">Settings Execution Status</h3>
$settingsSections

</body></html>
"@

    $logFilePath = Join-Path $LogFolder '00 - Execution Report.html'
    $htmlOut | Out-File -FilePath $logFilePath -Encoding UTF8 -Force   
}

function New-SettingsOverviewHtmlHC {
    param(
        [array]$MatrixRows, 
        [hashtable]$Html
    )

    $rowsHtml = foreach ($S in $MatrixRows | Sort-Object ID) {
        $types = @($S.Check.Type)
        $cls = Get-HtmlClassProbTypeHC ($types | Select-Object -First 1)

        $dur = if ($S.JobTime.Duration) {
            '{0:00}:{1:00}:{2:00}' -f $S.JobTime.Duration.Hours, $S.JobTime.Duration.Minutes, $S.JobTime.Duration.Seconds
        }
        else { 'NA' }

        $link = if ($S.FileContext.ReportFilePath) { 
            $S.FileContext.ReportFilePath
        }
        else { '#' }

        "<tr>
            <td class='$cls'></td>
            <td><a href='$link'>$($S.ID)</a></td>
            <td><a href='$link'>$([System.Net.WebUtility]::HtmlEncode($S.Setting.Formatted.ComputerName))</a></td>
            <td><a href='$link'>$([System.Net.WebUtility]::HtmlEncode($S.Setting.Formatted.Path))</a></td>
            <td><a href='$link'>$([System.Net.WebUtility]::HtmlEncode($S.Setting.Formatted.Action))</a></td>
            <td><a href='$link'>$dur</a></td>
        </tr>"
    }

    if (-not $rowsHtml) { return '' }
    
    return $Html.Templates.SettingsHeader + ($rowsHtml -join '')
}

function Build-MatrixEmailHtmlHC {
    param(
        [array]$AllMatrices, # Now accepts the flattened Context.AllMatrices array
        [hashtable]$Html
    )

    $output = ''

    # Group the flat jobs back by their parent Excel file
    $matricesByFile = $AllMatrices | Group-Object -Property { 
        $_.FileContext.Item.FullName } | Sort-Object Name

    foreach ($fileGroup in $matricesByFile) {
        $firstMatrix = $fileGroup.Group[0]

        # 1. Metadata & Excel Info Header
        $file = [System.Net.WebUtility]::HtmlEncode($firstMatrix.Item.Name)
        $modBy = [System.Net.WebUtility]::HtmlEncode($firstMatrix.ExcelInfo.LastModifiedBy ?? 'Unknown')
        $modDt = if ($firstMatrix.ExcelInfo.Modified -is [datetime]) {
            $firstMatrix.ExcelInfo.Modified.ToString('dd/MM/yyyy HH:mm:ss')
        }
        else { 'Unknown' }

        # 2. Global File/Sheet Checks
        $globalSections = @(
            New-HtmlSectionHC 'File Checks' $firstMatrix.Check
            if ($firstMatrix.FileContext.Sheets.FormData.Check) {
                New-HtmlSectionHC 'FormData Checks' $firstMatrix.FileContext.Sheets.FormData.Check
            }
            if ($firstMatrix.FileContext.Sheets.Permissions.Check) {
                New-HtmlSectionHC 'Permissions Checks' $firstMatrix.FileContext.Sheets.Permissions.Check
            }
        ) -join ''

        # 3. Settings Overview Table (Calling our updated function)
        $settingsOverview = New-SettingsOverviewHtmlHC `
            -MatrixRows $fileGroup.Group `
            -Html $Html

        # 4. Settings Detailed Results (This adds the checks below the overview!)
        $settingsDetails = ''
        foreach ($matrixRow in ($fileGroup.Group | Sort-Object ID)) {
            if ($matrixRow.Check -and $matrixRow.Check.Count -gt 0) {
                $header = "Settings Details (ID: $($matrixRow.ID)) - $($matrixRow.Setting.Formatted.ComputerName)"
                
                $settingsDetails += New-HtmlSectionHC $header $matrixRow.Check
            }
        }

        # 5. Assemble the HTML block for this specific Excel file
        $saveLink = $firstMatrix.SaveFullName ?? '#' 
        
        $output += @"
<table class="matrixTable">
<tr><th class="matrixTitle" colspan="8"><a href="$saveLink">$file</a></th></tr>
<tr><th class="matrixFileInfo" colspan="8">Last change: $modBy @ $modDt</th></tr>
$globalSections
$settingsOverview
$settingsDetails
</table><br><br>
"@
    }

    return $output
}

function Write-MatrixSettingLogHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Matrix,
        [Parameter(Mandatory)][hashtable]$Html,
        [Parameter(Mandatory)][string]$LogFolder
    )

    if (-not (Test-Path -LiteralPath $LogFolder -PathType Container)) { 
        return $null 
    }

    $sections = @(
        New-HtmlSectionHC 'Setting Validation' $Matrix.Check
    ) -join ''


    $htmlOut = @"
<!DOCTYPE html>
<html><head>
$($Html.Style)
$($Html.TroubleshootingStyle)
</head><body>
<h1>Settings Log - ID $($Matrix.ID)</h1>
<table class="matrixTable" style="width:100%;">
$sections
</table>
$($Html.Templates.LegendTable)
</body></html>
"@

    $logFilePath = Join-Path `
        -Path $LogFolder `
        -ChildPath "ID $($Matrix.ID) - Settings.html"

    $htmlOut | Out-File -FilePath $logFilePath -Encoding UTF8 -Force   
}

function Build-ErrorWarningTableHC {
    param($CounterData, $SystemErrors)

    if ($CounterData.TotalErrors -eq 0 -and $CounterData.TotalWarnings -eq 0) {
        return ''
    }

    '<p><b>Detected issues:</b></p>'
}

function Generate-MailBodyHtmlHC {
    param(
        $Settings,
        $Html,
        $ExportedFiles,
        $AttNote,
        $DurStr,
        [datetime]$ScriptStartTime,
        $LogFolder
    )

    @"
<!DOCTYPE html>
<html>
<head>$($Html.Style)</head>
<body>
<h1>$($Settings.ScriptName)</h1>
<hr>
$($Settings.SendMail.Body)
$($Html.ErrorWarningTable)
$($Html.MatrixTables)
<hr>
</body>
</html>
"@
}