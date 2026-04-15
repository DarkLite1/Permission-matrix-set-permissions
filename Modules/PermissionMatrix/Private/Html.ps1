<#
    Html.ps1
    Consolidated HTML rendering logic for Toolbox.PermissionMatrixHC
#>

#region Initialize-HtmlStructureHC
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
#endregion

#region Helper functions
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
    param([string]$Title, [object]$Checks)

    if (-not $Checks) { return '' }

    $rows = $Checks |
    ConvertTo-StructuredObjectHC |
    ForEach-Object { New-HtmlCheckRowHC $_ }

    "<tr><th class='matrixHeader' colspan='8'>$Title</th></tr>$($rows -join '')"
}

function New-SettingsTableHtmlHC {
    param($MatrixItem, $Html)

    if (-not $MatrixItem.Settings) { return '' }

    $fatalFile = @($MatrixItem.File.Check?.Type) -contains 'FatalError'
    $fatalPerms = @($MatrixItem.Permissions.Check?.Type) -contains 'FatalError'
    if ($fatalFile -or $fatalPerms) { return '' }

    $rows = foreach ($S in $MatrixItem.Settings | Sort-Object ID) {

        if (-not $S.Check) { continue }

        $types = @($S.Check.Type)
        $cls = Get-HtmlClassProbTypeHC ($types | Select-Object -First 1)

        $dur = if ($S.JobTime.Duration) {
            '{0:00}:{1:00}:{2:00}' -f
            $S.JobTime.Duration.Hours,
            $S.JobTime.Duration.Minutes,
            $S.JobTime.Duration.Seconds
        }
        else { 'NA' }

        $link = $MatrixItem.TroubleshootingLogPath ?? '#'

        "<tr>
            <td class='$cls'></td>
            <td><a href='$link'>$($S.ID)</a></td>
            <td><a href='$link'>$([System.Net.WebUtility]::HtmlEncode($S.Import.ComputerName))</a></td>
            <td><a href='$link'>$([System.Net.WebUtility]::HtmlEncode($S.Import.Path))</a></td>
            <td><a href='$link'>$([System.Net.WebUtility]::HtmlEncode($S.Import.Action))</a></td>
            <td><a href='$link'>$dur</a></td>
        </tr>"
    }

    if (-not $rows) { return '' }

    $Html.Templates.SettingsHeader + ($rows -join '')
}
#endregion

#region Build-MatrixEmailHtmlHC
function Build-MatrixEmailHtmlHC {
    param(
        [array]$ImportedMatrix,
        [object]$Html
    )

    $output = ''

    foreach ($Item in $ImportedMatrix | Sort-Object { $_.File.Item.Name }) {

        $sections = @(
            New-HtmlSectionHC 'File' $Item.File.Check
            New-HtmlSectionHC 'FormData' $Item.FormData.Check
            New-HtmlSectionHC 'Permissions' $Item.Permissions.Check
        ) -join ''

        $settings = New-SettingsTableHtmlHC $Item $Html

        $file = [System.Net.WebUtility]::HtmlEncode($Item.File.Item.Name)
        $modBy = [System.Net.WebUtility]::HtmlEncode($Item.File.ExcelInfo.LastModifiedBy ?? 'Unknown')
        $modDt = if ($Item.File.ExcelInfo.Modified -is [datetime]) {
            $Item.File.ExcelInfo.Modified.ToString('dd/MM/yyyy HH:mm:ss')
        }
        else { 'Unknown' }

        $output += @"
<table class="matrixTable">
<tr><th class="matrixTitle" colspan="8"><a href="$($Item.File.SaveFullName)">$file</a></th></tr>
<tr><th class="matrixFileInfo" colspan="8">Last change: $modBy @ $modDt</th></tr>
$sections
$settings
</table><br><br>
"@
    }

    return $output
}
#endregion

function Write-MatrixTroubleshootingLogHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Matrix,
        [Parameter(Mandatory)][hashtable]$Html
    )

    $folder = $Matrix.File.LogFolder
    if ([string]::IsNullOrWhiteSpace($folder) -or -not (Test-Path -LiteralPath $folder -PathType Container)) { 
        return $null 
    }

    $sections = @(
        New-HtmlSectionHC 'File' $Matrix.File.Check
        New-HtmlSectionHC 'FormData' $Matrix.FormData.Check
        New-HtmlSectionHC 'Permissions' $Matrix.Permissions.Check
    ) -join ''

    $htmlOut = @"
<!DOCTYPE html>
<html><head>
$($Html.Style)
$($Html.TroubleshootingStyle)
</head><body>
<h1>Troubleshooting Log</h1>
<table class="matrixTable" style="width:100%;">
$sections
</table>
$($Html.Templates.LegendTable)
</body></html>
"@

    $path = Join-Path $folder '00 - Troubleshooting Log.html'
    $htmlOut | Out-File -FilePath $path -Encoding UTF8 -Force
    
    return $path
}

#region Build-ErrorWarningTableHC
function Build-ErrorWarningTableHC {
    param($CounterData, $SystemErrors)

    if ($CounterData.TotalErrors -eq 0 -and $CounterData.TotalWarnings -eq 0) {
        return ''
    }

    '<p><b>Detected issues:</b></p>'
}
#endregion

#region Generate-MailBodyHtmlHC
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
#endregion