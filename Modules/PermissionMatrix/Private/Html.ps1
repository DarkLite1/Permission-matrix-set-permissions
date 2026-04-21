<#
    Html.ps1
    Consolidated HTML rendering logic for Toolbox.PermissionMatrixHC
#>

function Initialize-HtmlStructureHC {

    $style = @'
<style type="text/css">
    a { color: black; text-decoration: underline; }
    a:hover { color: blue; }
    body { font-family:verdana; font-size:14px; background-color:white; }
    h1, h2, h3 { margin-bottom: 0; }
    p.italic { font-style: italic; font-size: 12px; }
    table { border-collapse:collapse; padding:3px; }
    td, th { border:1px none; padding:3px; }
    .matrixTable { border:1px solid black; border-spacing:0 0.6em; width:600px; }
    .matrixTitle { background-color:lightgrey; text-align:center; padding:6px; }
    .matrixHeader { letter-spacing:5pt; font-style:italic; }
    .matrixFileInfo { font-size:12px; font-style:italic; text-align:center; }
    .legendTable { border:1px solid black; table-layout:fixed; }
    .legendTable td { text-align:center; }
    .probTypeError { background-color:red; }
    .probTypeWarning { background-color:orange; }
    .probTypeInfo { background-color:lightgrey; }
    .probTextError { color:red; font-weight:bold; }
    .probTextWarning { color:orange; font-weight:bold; }
    .aboutTable th, .aboutTable td { color:#8f8c8c; }
    base { target="_blank" }
</style>
'@

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
    param([object]$CheckObject)

    $cls = Get-HtmlClassProbTypeHC $CheckObject.Type
    $msg = [System.Net.WebUtility]::HtmlEncode($CheckObject.Message)
    $desc = [System.Net.WebUtility]::HtmlEncode($CheckObject.Description)

    "<tr class='$cls'><td></td><td>$($CheckObject.Name)</td><td>$msg</td><td>$desc</td></tr>"
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

    $comp = [System.Net.WebUtility]::HtmlEncode($MatrixItem.Setting.Raw.ComputerName)
    $path = [System.Net.WebUtility]::HtmlEncode($MatrixItem.Setting.Raw.Path)
    $group = [System.Net.WebUtility]::HtmlEncode($MatrixItem.Setting.Raw.GroupName)
    $site = [System.Net.WebUtility]::HtmlEncode($MatrixItem.Setting.Raw.SiteCode)

    #region Get Status & Colors
    $errCount = @($MatrixItem.Check | Where-Object Type -EQ 'FatalError').Count
    $warnCount = @($MatrixItem.Check | Where-Object Type -EQ 'Warning').Count
        
    if ($errCount -gt 0) {
        $headerColor = '#ffcccc' # Light Red
        $statusText = "Failed ($errCount Errors, $warnCount Warnings)"
    }
    elseif ($warnCount -gt 0) {
        $headerColor = '#ffe6cc' # Light Orange
        $statusText = "Completed with Warnings ($warnCount)"
    }
    else {
        $headerColor = '#d9f2d9' # Light Green
        $statusText = 'Success'
    }
    #endregion

    #region Get job start & end time
    $start = if ($MatrixItem.JobTime.Start) {
        $MatrixItem.JobTime.Start.ToString('dd/MM/yyyy HH:mm:ss (dddd)') 
    }
    else { 'N/A' }
    $end = if ($MatrixItem.JobTime.End) {
        $MatrixItem.JobTime.End.ToString('dd/MM/yyyy HH:mm:ss (dddd)') 
    }
    else { 'N/A' }
    #endregion

    #region Create HTML check table
    $checkTable = if ($MatrixItem.Check -and $MatrixItem.Check.Count -gt 0) {
        "<table class='matrixTable' style='width:100%; border:none; margin-top:10px;'>$(New-HtmlSectionHC 'Detailed Results' $MatrixItem.Check)</table>"
    }
    else {
        "<p style='padding-top:10px; font-style:italic;'>No issues detected. Execution successful.</p>"
    }
    #endregion

    #region Create JSON file link ONLY if there are errors
    $jsonLink = ''
    if ($MatrixItem.Check -and $MatrixItem.Check.Count -gt 0) {
        $jsonFileName = "ID $($MatrixItem.ExcelID) - Details.json"
        $jsonLink = @"
        <div style="margin-top: 15px; text-align: right; font-size: 12px;">
            <a href="$jsonFileName" style="color: blue;">View Raw JSON Details ($($MatrixItem.Check.Count) records)</a>
        </div>
"@
    }
    #endregion

    return @"
<div style="border: 1px solid black; margin-bottom: 25px;">
    <div style="background-color: $headerColor; padding: 10px; border-bottom: 1px solid black; font-weight: bold; font-size: 14px;">
        ID $safeId | $comp | $path
        <span style="float: right;">Status: $statusText</span>
    </div>
    <div style="padding: 10px; background-color: #f2f2f2; border-bottom: 1px solid #ccc;">
        <h3 style="margin-top:0; margin-bottom:5px;">About</h3>
        <table style="border:none; font-size:13px;">
            <tr><td style="width:120px; font-weight:bold; color:#8f8c8c;">GroupName:</td><td>$group</td></tr>
            <tr><td style="font-weight:bold; color:#8f8c8c;">SiteCode:</td><td>$site</td></tr>
            <tr><td style="font-weight:bold; color:#8f8c8c;">Start time:</td><td>$start</td></tr>
            <tr><td style="font-weight:bold; color:#8f8c8c;">End time:</td><td>$end</td></tr>
        </table>
    </div>
    <div style="padding: 10px;">
        $checkTable
        $jsonLink
    </div>
</div>
"@
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

        $link = if ($S.FileContext.LogFolder) { 
            "$($S.FileContext.LogFolder)\00 - Execution Report.html"
        }
        else { '#' }

        "<tr>
            <td class='$cls'></td>
            <td><a href='$link'>$($S.ID)</a></td>
            <td><a href='$link'>$([System.Net.WebUtility]::HtmlEncode($S.Setting.Raw.ComputerName))</a></td>
            <td><a href='$link'>$([System.Net.WebUtility]::HtmlEncode($S.Setting.Raw.Path))</a></td>
            <td><a href='$link'>$([System.Net.WebUtility]::HtmlEncode($S.Setting.Raw.Action))</a></td>
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
                $safeId = if ($matrixRow.ID) { $matrixRow.ID } else { 'Unknown' }
                $header = "Settings Details (ID: $safeId) - $($matrixRow.Setting.Raw.ComputerName)"
                
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

function Write-MatrixExecutionReportHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$FileMatrices, 
        [Parameter(Mandatory)][hashtable]$Html,
        [Parameter(Mandatory)][string]$LogFolder
    )

    if (-not (Test-Path `
                -LiteralPath $LogFolder `
                -PathType Container)
    ) { return $null }

    $firstMatrix = $FileMatrices[0]

    $modBy = [System.Net.WebUtility]::HtmlEncode(
        $firstMatrix.FileContext.ExcelInfo.LastModifiedBy ?? 'Unknown'
    )
    $modDt = if ($firstMatrix.FileContext.ExcelInfo.Modified -is [datetime]) {
        $firstMatrix.FileContext.ExcelInfo.Modified.ToString('dd/MM/yyyy HH:mm:ss')
    }
    else { 'Unknown' }

    $fileSections = @(
        New-HtmlSectionHC 'File Checks' $firstMatrix.FileContext.Check

        if ($firstMatrix.FileContext.Sheets.FormData.Check) {
            New-HtmlSectionHC 'FormData Checks' `
                $firstMatrix.FileContext.Sheets.FormData.Check
        }
        if ($firstMatrix.FileContext.Sheets.Permissions.Check) {
            New-HtmlSectionHC 'Permissions Checks' `
                $firstMatrix.FileContext.Sheets.Permissions.Check
        }
    ) -join ''

    $settingsSections = ''
    foreach (
        $matrix in 
        ($FileMatrices | Sort-Object `
        { $_.Setting.Raw.ComputerName }, `
        { $_.Setting.Raw.Path }, `
        { $_.ID }
        )
    ) {
        $settingsSections += New-SettingsCardHtmlHC -MatrixItem $matrix
    }

    $htmlOut = @"
<!DOCTYPE html>
<html><head>
$($Html.Style)
$($Html.TroubleshootingStyle)
</head><body>
<h1>Execution & Troubleshooting Report</h1>
<h2>File: $($firstMatrix.FileContext.Item.Name)</h2>
<p class="matrixFileInfo" style="text-align:left; margin-top:5px; margin-bottom:20px;">
    Last change: $modBy @ $modDt
</p>

<h3>Global File Status</h3>
<table class="matrixTable" style="width:100%;">
$fileSections
</table>

<br>
<h3>Settings Execution Status</h3>
$settingsSections

$($Html.Templates.LegendTable)
</body></html>
"@

    $logFilePath = Join-Path $LogFolder '00 - Execution Report.html'
    $htmlOut | Out-File -FilePath $logFilePath -Encoding UTF8 -Force   
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

    $safeId = if ($Matrix.ID) { $Matrix.ID } else { 'Unknown' }

    $htmlOut = @"
<!DOCTYPE html>
<html><head>
$($Html.Style)
$($Html.TroubleshootingStyle)
</head><body>
<h1>Settings Log - ID $safeId</h1>
<table class="matrixTable" style="width:100%;">
$sections
</table>
$($Html.Templates.LegendTable)
</body></html>
"@

    $logFilePath = Join-Path `
        -Path $LogFolder `
        -ChildPath "ID $safeId - Settings.html"

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