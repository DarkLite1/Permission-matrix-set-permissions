function Build-MatrixEmailHtml {
    param(
        [Parameter(Mandatory)][array]$ImportedMatrix,
        [Parameter(Mandatory)][object]$Html
    )

    function New-SectionHtml {
        param([string]$Name, [object]$Checks)

        if (-not $Checks) { return '' }

        $rows = ($Checks | ConvertTo-StructuredObjectHC | ForEach-Object {
                New-HtmlCheckRow -CheckObject $_
            }) -join ''

        return "<tr><th class='matrixHeader' colspan='8'>$Name</th></tr>$rows"
    }

    function New-SettingsTableHtml {
        param([object]$MatrixItem, [object]$Html)

        $fatalFile = @($MatrixItem.File.Check?.Type) -contains 'FatalError'
        $fatalPerms = @($MatrixItem.Permissions.Check?.Type) -contains 'FatalError'

        if (-not $MatrixItem.Settings -or $fatalFile -or $fatalPerms) {
            return ''    # Suppress table when File or Permissions have fatal errors
        }

        $rows = foreach ($S in $MatrixItem.Settings | Sort-Object ID) {

            $types = @($S.Check?.Type).Where({ $_ })
            if (-not $types) { continue }

            $class = if ($types -contains 'FatalError') { 'probTypeError' }
            elseif ($types -contains 'Warning') { 'probTypeWarning' }
            elseif ($types -contains 'Information') { 'probTypeInfo' }
            else { '' }

            $duration = if ($S.JobTime.Duration) {
                '{0:00}:{1:00}:{2:00}' -f $S.JobTime.Duration.Hours,
                $S.JobTime.Duration.Minutes,
                $S.JobTime.Duration.Seconds
            }
            else {
                'NA'
            }

            $link = $MatrixItem.TroubleshootingLogPath ?? '#'

            $encComp = [System.Net.WebUtility]::HtmlEncode($S.Import.ComputerName)
            $encPath = [System.Net.WebUtility]::HtmlEncode($S.Import.Path)
            $encAction = [System.Net.WebUtility]::HtmlEncode($S.Import.Action)

            "<tr>
                <td class='$class'></td>
                <td><a href='$link'>$($S.ID)</a></td>
                <td><a href='$link'>$encComp</a></td>
                <td><a href='$link'>$encPath</a></td>
                <td><a href='$link'>$encAction</a></td>
                <td><a href='$link'>$duration</a></td>
             </tr>"
        }

        if (-not $rows) { return '' }

        return $Html.Templates.SettingsHeader + ($rows -join '')
    }

    $resultHtml = ''

    foreach ($Item in $ImportedMatrix | Sort-Object { $_.File.Item.Name }) {

        # Build the 3 built‑in sections
        $sectionHtml = @(
            New-SectionHtml -Name 'File' -Checks $Item.File.Check
            New-SectionHtml -Name 'FormData' -Checks $Item.FormData.Check
            New-SectionHtml -Name 'Permissions' -Checks $Item.Permissions.Check
        ) -join ''

        # Build settings table (only if no fatal file/permissions)
        $settingsHtml = New-SettingsTableHtml -MatrixItem $Item -Html $Html

        # Metadata
        $encFileName = [System.Net.WebUtility]::HtmlEncode($Item.File.Item.Name)

        $modBy = [System.Net.WebUtility]::HtmlEncode(
            $Item.File.ExcelInfo.LastModifiedBy ??
            'Unknown'
        )

        $modDate = $Item.File.ExcelInfo.Modified
        if ($modDate -is [datetime]) {
            $modDate = $modDate.ToString('dd/MM/yyyy HH:mm:ss')
        }
        elseif ($modDate) {
            $modDate = [System.Net.WebUtility]::HtmlEncode("$modDate")
        }
        else {
            $modDate = 'Unknown'
        }

        # Assemble full table
        $resultHtml += @"
<table class="matrixTable">
    <tr>
        <th class="matrixTitle" colspan="8">
            <a href="$($Item.File.SaveFullName)">$encFileName</a>
        </th>
    </tr>
    <tr>
        <th class="matrixFileInfo" colspan="8">
            Last change: $modBy @ $modDate
        </th>
    </tr>
    $sectionHtml
    $settingsHtml
</table>
<br><br>
"@
    }

    return $resultHtml
}
function Process-MatrixObjects {
    param(
        [Parameter(Mandatory)][array]$ImportedMatrix,
        [Parameter(Mandatory)][object]$Html
    )

    #
    # Process each matrix item:
    #   - Generate its troubleshooting log
    #   - Attach TroubleshootingLogPath property
    #

    foreach (
        $matrixItem in
        $ImportedMatrix | Sort-Object { $_.File.Item.Name }
    ) {

        $logPath = $null

        try {
            $logPath = Write-MatrixTroubleshootingLog `
                -Matrix $matrixItem `
                -Html $Html
        }
        catch {
            Write-Warning "Failed to build troubleshooting log for '$($matrixItem.File.Item.Name)': $_"
        }

        #
        # Add or update TroubleshootingLogPath on the matrix item
        #
        $matrixItem |
        Add-Member -NotePropertyName TroubleshootingLogPath `
            -NotePropertyValue $logPath `
            -Force
    }

    return $ImportedMatrix
}
function Write-MatrixTroubleshootingLog {
    param(
        [Parameter(Mandatory)][object]$Matrix,
        [Parameter(Mandatory)][object]$Html
    )

    try {
        #
        # Validate log folder
        #
        $logFolder = $Matrix.File.LogFolder
        if (-not (Test-Path -LiteralPath $logFolder -PathType Container)) {
            return $null
        }

        #
        # File metadata (safe encoded)
        #
        $fileName = [System.Net.WebUtility]::HtmlEncode($Matrix.File.Item.Name)

        $modifiedBy = $Matrix.File.ExcelInfo.LastModifiedBy
        $modBy = if ($modifiedBy) {
            [System.Net.WebUtility]::HtmlEncode($modifiedBy.ToString().Trim())
        }
        else {
            'Unknown'
        }

        $modifiedTime = $Matrix.File.ExcelInfo.Modified
        $modDate = if ($modifiedTime -is [datetime]) {
            $modifiedTime.ToString('dd/MM/yyyy HH:mm:ss')
        }
        elseif ($modifiedTime) {
            [System.Net.WebUtility]::HtmlEncode("$modifiedTime")
        }
        else {
            'Unknown'
        }

        #
        # Function: Render a section (File / FormData / Permissions)
        #
        function New-SectionHtml {
            param(
                [string]$SectionName,
                [object]$Checks
            )

            if (-not $Checks) { return '' }

            $rows = ($Checks | ConvertTo-StructuredObjectHC | ForEach-Object {
                    New-HtmlCheckRow -CheckObject $_
                }) -join ''

            return "<tr><th class='matrixHeader' colspan='8'>$SectionName</th></tr>$rows"
        }

        #
        # Build 3 main sections
        #
        $sectionsHtml = @(
            New-SectionHtml -SectionName 'File' -Checks $Matrix.File.Check
            New-SectionHtml -SectionName 'FormData' -Checks $Matrix.FormData.Check
            New-SectionHtml -SectionName 'Permissions' -Checks $Matrix.Permissions.Check
        ) -join ''

        #
        # Settings checks (if any)
        #
        $settingsHtml = ''

        if ($Matrix.Settings) {

            $settingsRows = foreach ($S in $Matrix.Settings | Sort-Object ID) {
                if (-not $S.Check) { continue }

                #
                # Heading row for each setting entry
                #
                $encComp = [System.Net.WebUtility]::HtmlEncode($S.Import.ComputerName)
                $encPath = [System.Net.WebUtility]::HtmlEncode($S.Import.Path)

                "<tr><td colspan='8' style='background-color:#eee;'>
                    <b>Setting ID: $($S.ID)</b> ($encComp - $encPath)
                 </td></tr>" +

                #
                # Individual checks rendered via the shared row builder
                #
                (
                    $S.Check | ConvertTo-StructuredObjectHC | ForEach-Object {
                        New-HtmlCheckRow -CheckObject $_
                    }
                ) -join ''
            }

            if ($settingsRows) {
                $settingsHtml = "<tr><th class='matrixHeader' colspan='8'>Settings Checks</th></tr>$($settingsRows -join '')"
            }
        }

        #
        # Combine full table
        #
        $tableHtml = @"
<table class="matrixTable" style="width: 100%;">
    <tr><th colspan="8" class="matrixHeader">Troubleshooting details</th></tr>
    <tr><td colspan="8"><strong>Last change:</strong> $modBy @ $modDate</td></tr>
    $sectionsHtml
    $settingsHtml
</table>
<br>
$($Html.Templates.LegendTable)
"@

        #
        # Final HTML document
        #
        $fullHtml = @"
<!DOCTYPE html>
<html>
<head>
    $($Html.Style)
    $($Html.TroubleshootingStyle)
</head>
<body>
    <h1>Troubleshooting Log: $fileName</h1>
    $tableHtml
</body>
</html>
"@

        #
        # Write file
        #
        $filePath = Join-Path -Path $logFolder -ChildPath '00 - Troubleshooting Log.html'
        $fullHtml | Out-File -LiteralPath $filePath -Encoding UTF8 -Force

        return $filePath
    }
    catch {
        Write-Warning "Troubleshooting log failed for '$($Matrix.File.Item.Name)': $_"
        return $null
    }
}
function New-HtmlCheckRow {
    param(
        [Parameter(Mandatory)]
        [object]$CheckObject
    )

    # Determine CSS class based on type (Error / Warning / Info)
    $cssClass = Get-HtmlClassProbTypeHC -Name $CheckObject.Type

    # HTML-encode dynamic fields
    $name = [System.Net.WebUtility]::HtmlEncode($CheckObject.Name)
    $desc = [System.Net.WebUtility]::HtmlEncode($CheckObject.Description)

    # Optional list of values
    $listHtml = Format-HtmlList -Value $CheckObject.Value

    # Output final row
    return @"
<tr>
    <td class="$cssClass"></td>
    <td colspan="7">
        <p class="probTitle">$name</p>
        <p>$desc</p>
        $listHtml
    </td>
</tr>
"@
}
function Get-HtmlClassProbTypeHC {
    [OutputType([string])]
    param (
        [Parameter(Mandatory)]
        [ValidateSet('FatalError', 'Warning', 'Information')]
        [string]$Name
    )

    switch ($Name) {
        'FatalError' { return 'probTypeError' }
        'Warning' { return 'probTypeWarning' }
        'Information' { return 'probTypeInfo' }
    }
}
function Initialize-HtmlStructure {
    $style = @'
        <style type="text/css">
            a { color: black; text-decoration: underline; }
            a:hover { color: blue; }
            body { font-family:verdana; font-size:14px; background-color:white; }
            h1, h2, h3 { margin-bottom: 0; }
            p.italic { font-style: italic; font-size: 12px; }
            table { border-collapse:collapse; border:0px none; padding:3px; text-align:left; }
            td, th { border-collapse:collapse; border:1px none; padding:3px; text-align:left; }
            .matrixTable { border: 1px solid Black; border-collapse: separate; border-spacing: 0px 0.6em; width: 600px; }
            .matrixTitle { border: none; background-color: lightgrey; text-align: center; padding: 6px; }
            .matrixHeader { font-weight: normal; letter-spacing: 5pt; font-style: italic; }
            .matrixFileInfo { font-weight: normal; font-size: 12px; font-style: italic; text-align: center; }
            .legendTable { border-collapse: collapse; border: 1px solid Black; table-layout: fixed; }
            .legendTable td { text-align: center; }
            .probTitle { font-weight: bold; }
            .probTypeWarning { background-color: orange; }
            .probTextWarning { color: orange; font-weight: bold; }
            .probTypeError { background-color: red; }
            .probTextError { color: red; font-weight: bold; }
            .probTypeInfo { background-color: lightgrey; }
            table tbody tr td a { display: block; width: 100%; height: 100%; }
            .aboutTable th, .aboutTable td { color: rgb(143, 140, 140); font-weight: normal; }
            base { target="_blank" }
        </style>
'@
    $troubleshootingStyle = @'
        <style type="text/css">
            body { margin: 20px; }
        </style>
'@
    return @{
        Style                = $style
        TroubleshootingStyle = $troubleshootingStyle
        Templates            = @{
            SettingsHeader = '<th class="matrixHeader" colspan="8">Settings</th><tr><td></td><td>ID</td><td>ComputerName</td><td>Path</td><td>Action</td><td>Duration</td></tr>'
            LegendTable    = '<table class="legendTable"><tr><td class="probTypeError" style="width:150px;">Error</td><td class="probTypeWarning" style="width:150px;">Warning</td><td class="probTypeInfo" style="width:150px;">Information</td></tr></table>'
        }
    }
}
function Format-HtmlList {
    param([object]$Value)
    if (-not $Value) { return '' }
    if ($Value.Count -le 5 -and $Value -isnot [hashtable]) {
        $encodedItems = @($Value).ForEach(
            { "<li>$([System.Net.WebUtility]::HtmlEncode($_))</li>" }
        ) -join ''
        return "<ul>$encodedItems</ul>"
    }
    return '<p><i>Check JSON dump for multiple items.</i></p>'
}
