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

function New-SettingsOverviewHtmlHC {
    param(
        [array]$MatrixRows, 
        [hashtable]$Html
    )

    $rowsHtml = foreach ($S in $MatrixRows | Sort-Object ExcelID) {
        $types = @($S.Check.Type)
        $cls = Get-HtmlClassProbTypeHC ($types | Select-Object -First 1)

        $dur = if ($S.JobTime.Duration) {
            '{0:00}:{1:00}:{2:00}' -f $S.JobTime.Duration.Hours, $S.JobTime.Duration.Minutes, $S.JobTime.Duration.Seconds
        }
        else { 'NA' }

        $safeId = if ($S.ExcelID) { $S.ExcelID } else { 'Unknown' }
        
        # We can link the row directly to the consolidated log folder we made earlier!
        $link = if ($S.FileContext.LogFolder) { "$($S.FileContext.LogFolder)\00 - Execution Report.html" } else { '#' }

        "<tr>
            <td class='$cls'></td>
            <td><a href='$link'>$safeId</a></td>
            <td><a href='$link'>$([System.Net.WebUtility]::HtmlEncode($S.EnabledSetting.Raw.ComputerName))</a></td>
            <td><a href='$link'>$([System.Net.WebUtility]::HtmlEncode($S.EnabledSetting.Raw.Path))</a></td>
            <td><a href='$link'>$([System.Net.WebUtility]::HtmlEncode($S.EnabledSetting.Raw.Action))</a></td>
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
    $matricesByFile = $AllMatrices | Group-Object -Property { $_.File.Item.FullName } | Sort-Object Name

    foreach ($fileGroup in $matricesByFile) {
        $firstMatrix = $fileGroup.Group[0]

        # 1. Metadata & Excel Info Header
        $file = [System.Net.WebUtility]::HtmlEncode($firstMatrix.File.Item.Name)
        $modBy = [System.Net.WebUtility]::HtmlEncode($firstMatrix.File.ExcelInfo.LastModifiedBy ?? 'Unknown')
        $modDt = if ($firstMatrix.File.ExcelInfo.Modified -is [datetime]) {
            $firstMatrix.File.ExcelInfo.Modified.ToString('dd/MM/yyyy HH:mm:ss')
        }
        else { 'Unknown' }

        # 2. Global File/Sheet Checks
        $globalSections = @(
            New-HtmlSectionHC 'File Checks' $firstMatrix.File.Check
            if ($firstMatrix.FileContext.Sheets.FormData.Check) {
                New-HtmlSectionHC 'FormData Checks' $firstMatrix.FileContext.Sheets.FormData.Check
            }
            if ($firstMatrix.FileContext.Sheets.Permissions.Check) {
                New-HtmlSectionHC 'Permissions Checks' $firstMatrix.FileContext.Sheets.Permissions.Check
            }
        ) -join ''

        # 3. Settings Overview Table (Calling our updated function)
        $settingsOverview = New-SettingsOverviewHtmlHC -MatrixRows $fileGroup.Group -Html $Html

        # 4. Settings Detailed Results (This adds the checks below the overview!)
        $settingsDetails = ''
        foreach ($matrixRow in ($fileGroup.Group | Sort-Object ExcelID)) {
            if ($matrixRow.Check -and $matrixRow.Check.Count -gt 0) {
                $safeId = if ($matrixRow.ExcelID) { $matrixRow.ExcelID } else { 'Unknown' }
                $header = "Settings Details (ID: $safeId) - $($matrixRow.EnabledSetting.Raw.ComputerName)"
                
                $settingsDetails += New-HtmlSectionHC $header $matrixRow.Check
            }
        }

        # 5. Assemble the HTML block for this specific Excel file
        $saveLink = $firstMatrix.File.SaveFullName ?? '#' 
        
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
    $fileSections = @(
        New-HtmlSectionHC 'File Checks' $firstMatrix.File.Check
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
    foreach ($matrix in $FileMatrices) {
        $safeId = if ($matrix.ID) { $matrix.ID } else { 'Unknown' }
        $header = "Settings Row (ID: $safeId) - $($matrix.Setting.Raw.ComputerName)"
        
        $settingsSections += New-HtmlSectionHC $header $matrix.Check
    }

    # 3. BUILD THE FINAL HTML
    $htmlOut = @"
<!DOCTYPE html>
<html><head>
$($Html.Style)
$($Html.TroubleshootingStyle)
</head><body>
<h1>Execution & Troubleshooting Report</h1>
<h2>File: $($firstMatrix.File.Item.Name)</h2>

<h3>Global File Status</h3>
<table class="matrixTable" style="width:100%;">
$fileSections
</table>

<br>
<h3>Settings Execution Status</h3>
<table class="matrixTable" style="width:100%;">
$settingsSections
</table>

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