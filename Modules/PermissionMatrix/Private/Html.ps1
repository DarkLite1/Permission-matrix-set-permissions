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

    # Grab the FULL ID string for the About table
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

    # Calculate Status & Colors
    $errCount = @($MatrixItem.Check | Where-Object Type -EQ 'FatalError').Count
    $warnCount = @($MatrixItem.Check | Where-Object Type -EQ 'Warning').Count
        
    if ($errCount -gt 0) {
        $headerColor = '#fee2e2' # Softer Red
        $statusText = "Failed ($errCount Errors, $warnCount Warnings)"
    }
    elseif ($warnCount -gt 0) {
        $headerColor = '#ffedd5' # Softer Orange
        $statusText = "Completed with Warnings ($warnCount)"
    }
    else {
        $headerColor = '#dcfce7' # Softer Green
        $statusText = 'Success'
    }

    # Calculate Duration String
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

    # Build custom lightweight check rows
    $checkTable = ''
    if ($MatrixItem.Check -and $MatrixItem.Check.Count -gt 0) {
        $checkRows = ''
        foreach ($c in $MatrixItem.Check) {
            $cls = Get-HtmlClassProbTypeHC $c.Type
            $name = [System.Net.WebUtility]::HtmlEncode($c.Name)
            $desc = [System.Net.WebUtility]::HtmlEncode($c.Description)
            
            $linkTarget = if ($c.JsonFileName) { $c.JsonFileName } else { '#' }
            
            $checkRows += @"
            <tr class='$cls' style='border-bottom: 1px solid #e5e7eb;'>
                <td style='width: 8px;'></td>
                <td style='font-weight: 600; width: 35%; padding: 8px 6px;'>
                    <a href="$linkTarget" style="color: #111827; text-decoration: underline;" title="Click to view full JSON details">
                        $name
                    </a>
                </td>
                <td style='padding: 8px 6px; color: #374151;'>$desc</td>
            </tr>
"@
        }
        
        $checkTable = @"
        <h4 style="margin:0 0 8px 0; color: #374151;">Detailed Results</h4>
        <table style='width:100%; border-collapse: collapse; font-size: 13px; background-color: white; border: 1px solid #d1d5db; border-radius: 4px; overflow: hidden;'>
            $checkRows
        </table>
"@
    }
    else {
        $checkTable = "<p style='padding-top:5px; font-style:italic; color: #6b7280; margin: 0;'>No issues detected. Execution successful.</p>"
    }

    return @"
<div style="border: 1px solid #d1d5db; border-radius: 8px; margin-bottom: 25px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); page-break-inside: avoid; overflow: hidden; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;">
    <div style="background-color: $headerColor; padding: 12px 16px; border-bottom: 1px solid #d1d5db; display: flex; justify-content: space-between; align-items: center;">
        <div style="font-size: 15px;">
            <span style="background-color: rgba(255,255,255,0.6); border: 1px solid rgba(0,0,0,0.1); padding: 3px 12px; border-radius: 12px; font-size: 13px; font-weight: 700; margin-right: 10px; color: #1f2937;">$comp</span>
            <span style="font-family: Consolas, monospace; font-size: 13.5px; color: #374151; background-color: rgba(255,255,255,0.4); padding: 2px 6px; border-radius: 4px;">$path</span>
        </div>
        <div style="font-size: 13px; font-weight: 700; color: #111827;">
            $statusText
        </div>
    </div>
    <div style="padding: 12px 16px; background-color: #f9fafb; border-bottom: 1px solid #e5e7eb;">
        <h3 style="margin-top:0; margin-bottom:0px; font-size: 14px; color: #374151;">About</h3>
        <table style="border:none; font-size:13px; border-collapse: separate; border-spacing: 0 6px;">
            <tr><td style="width:100px; font-weight:600; color:#6b7280;">ID:</td><td style="color: #111827; font-family: Consolas, monospace; font-size: 12px;">$fullId</td></tr>
            <tr><td style="font-weight:600; color:#6b7280;">GroupName:</td><td style="color: #111827;">$group</td></tr>
            <tr><td style="font-weight:600; color:#6b7280;">SiteCode:</td><td style="color: #111827;">$site</td></tr>
            <tr><td style="font-weight:600; color:#6b7280;">Duration:</td><td style="color: #111827;">$timeStr</td></tr>
        </table>
    </div>
    <div style="padding: 16px; background-color: #ffffff;">
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

    if (-not (Test-Path -LiteralPath $LogFolder -PathType Container)) { return $null }

    $firstMatrix = $FileMatrices[0]

    $modBy = [System.Net.WebUtility]::HtmlEncode($firstMatrix.FileContext.ExcelInfo.LastModifiedBy ?? 'Unknown')
    $modDt = if ($firstMatrix.FileContext.ExcelInfo.Modified -is [datetime]) {
        $firstMatrix.FileContext.ExcelInfo.Modified.ToString('dd/MM/yyyy HH:mm:ss')
    }
    else { 'Unknown' }

    $fileSections = @(
        New-HtmlSectionHC 'File Checks' $firstMatrix.FileContext.Check
        if ($firstMatrix.FileContext.Sheets.FormData.Check) {
            New-HtmlSectionHC 'FormData Checks' $firstMatrix.FileContext.Sheets.FormData.Check
        }
        if ($firstMatrix.FileContext.Sheets.Permissions.Check) {
            New-HtmlSectionHC 'Permissions Checks' $firstMatrix.FileContext.Sheets.Permissions.Check
        }
    ) -join ''

    $settingsSections = ''
    foreach ($matrix in ($FileMatrices | Sort-Object { $_.Setting.Raw.ComputerName }, { $_.Setting.Raw.Path }, { $_.ID })) {
        $settingsSections += New-SettingsCardHtmlHC -MatrixItem $matrix
    }

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

<h3 style="border-bottom: 2px solid #e5e7eb; padding-bottom: 5px; margin-bottom: 25px;">Global File Status</h3>
<table class="matrixTable" style="width:100%; margin-bottom: 25px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
$fileSections
</table>

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