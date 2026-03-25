function Build-ErrorWarningTable {
    param(
        [Parameter(Mandatory)][object]$CounterData,
        [Parameter(Mandatory)][object]$SystemErrors
    )

    #
    # Helper: build a table row in a single consistent way
    #
    function New-ErrorRow {
        param(
            [string]$CssClass,
            [string]$Label,
            [int]   $Count
        )
        return "<tr class='$CssClass'><th>$Label</th><td>$Count</td></tr>"
    }

    $rows = @()

    #
    # 1. System errors
    #
    if ($SystemErrors.Count -gt 0) {
        $rows += New-ErrorRow -CssClass 'probTextError' -Label 'System errors' -Count $SystemErrors.Count
    }

    #
    # 2. Matrix errors (excluding system errors)
    #
    $matrixErrors = $CounterData.TotalErrors - $SystemErrors.Count
    if ($matrixErrors -gt 0) {
        $rows += New-ErrorRow -CssClass 'probTextError' -Label 'Matrix errors' -Count $matrixErrors
    }

    #
    # 3. Matrix warnings
    #
    if ($CounterData.TotalWarnings -gt 0) {
        $rows += New-ErrorRow -CssClass 'probTextWarning' -Label 'Matrix warnings' -Count $CounterData.TotalWarnings
    }

    #
    # 4. If no rows, return empty string
    #
    if (-not $rows) {
        return ''
    }

    #
    # 5. Wrap rows in the final table
    #
    return "<p><b>Detected issues:</b></p><table class='errorWarningTable'>" +
    ($rows -join '') +
    '</table>'
}
function ConvertTo-HtmlValueHC {
    param(
        [Parameter(Mandatory)]
        $ErrorObj,
        [Parameter(Mandatory)]
        $SettingId,
        [Parameter(Mandatory)]
        [string]$LogFolderPath
    )

    if (-not $ErrorObj.Value) {
        return $null
    }
    elseif (
        ($ErrorObj.Value.Count -le 5) -and 
        (-not ($ErrorObj.Value -is [hashtable]))
    ) {
        return '<ul>{0}</ul>' -f $(@($ErrorObj.Value).ForEach({ "<li>$_</li>" }))
    }
    else {
        $safeName = "ID $SettingId - $($ErrorObj.Type) - $($ErrorObj.Name).txt".Split([IO.Path]::GetInvalidFileNameChars()) -join '_'

        $OutParams = @{
            LiteralPath = Join-Path -Path $LogFolderPath -ChildPath $safeName
            Encoding    = 'utf8'
            NoClobber   = $true
        }
        $ErrorObj | ConvertTo-Json -Depth 100 | ForEach-Object {
            [System.Text.RegularExpressions.Regex]::Unescape($_)
        } | Out-File @OutParams

        return '<ul><li><a href="{0}">{1} items</a></li></ul>' -f $OutParams.LiteralPath, $ErrorObj.Value.Count
    }
}
function Generate-MailBodyHtml {
    param(
        [Parameter(Mandatory)][object]$Settings,
        [Parameter(Mandatory)][object]$Html,
        [Parameter()][object]$ExportedFiles,
        [Parameter()][string]$AttNote,
        [Parameter()][string]$DurStr,
        [Parameter()][datetime]$ScriptStartTime,
        [Parameter()][string]$LogFolder
    )

    #
    # Helper: Create exported file links
    #
    function New-ExportListHtml {
        param([object]$Files)

        if (-not $Files -or $Files.Count -eq 0) {
            return ''
        }

        $items = $Files.GetEnumerator() |
        ForEach-Object {
            "<li>$($_.Value)$($_.Key)</a></li>"
        }

        return "<p><b>Exported $($Files.Count) file$(if($Files.Count-ne 1){'s'}):</b></p><ul>$($items -join '')</ul>"
    }

    #
    # Helper: Build the metadata table
    #
    function New-MetadataTable {
        param(
            [datetime]$Start,
            [string]$Duration,
            [string]$LogFolder
        )

        $startStr = $Start.ToString('dd/MM/yyyy HH:mm (dddd)')
        $logHtml = if ($LogFolder) {
            "<tr><th>Log files</th><td>$LogFolderOpen log folder</a></td></tr>"
        }

        return @"
<table class="aboutTable">
    <tr><th>Start time</th><td>$startStr</td></tr>
    <tr><th>Duration</th><td>$Duration</td></tr>
    $logHtml
    <tr><th>Host</th><td>$($host.Name)</td></tr>
    <tr><th>Computer</th><td>$env:COMPUTERNAME</td></tr>
    <tr><th>Account</th><td>$($env:USERDNSDOMAIN)\$($env:USERNAME)</td></tr>
</table>
"@
    }

    #
    # Compose sections
    #
    $exportHtml = New-ExportListHtml -Files $ExportedFiles
    $metaTable = New-MetadataTable -Start $ScriptStartTime -Duration $DurStr -LogFolder $LogFolder

    #
    # Main HTML document
    #
    return @"
<!DOCTYPE html>
<html>
<head>
    $($Html.Style)
</head>
<body>

<h1>$($Settings.ScriptName)</h1>
<hr size="2" color="#06cc7a">

$($Settings.SendMail.Body)
$($Html.ErrorWarningTable)
$exportHtml
$($Html.MatrixTables)
$AttNote

<hr size="2" color="#06cc7a">

$metaTable

</body>
</html>
"@
}