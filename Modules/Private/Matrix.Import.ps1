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
function Validate-JsonSchema {
    param(
        [Parameter(Mandatory)]
        [object]$JsonObject
    )

    $errors = [System.Collections.Generic.List[object]]::new()

    function Add-SchemaError {
        param([string]$Message)
        $errors.Add(
            [PSCustomObject]@{ 
                Message = $Message 
            }
        )
    }

    # --- 1. Required top-level objects ---
    foreach (
        $prop in 
        @('Matrix', 'Export', 'ServiceNow', 'MaxConcurrent', 'PSSessionConfiguration', 'Settings')
    ) {
        if ($null -eq $JsonObject.$prop) {
            Add-SchemaError "Property '$prop' not found"
        }
    }

    # If Settings missing, bail out (prevents deeper checks)
    if ($null -eq $JsonObject.Settings) {
        return $errors
    }

    # --- 2. Validate Settings structure ---
    if ($null -eq $JsonObject.Settings.SaveLogFiles.Where.Folder) {
        Add-SchemaError "Property 'Settings.SaveLogFiles.Where.Folder' not found"
    }

    if ($null -eq $JsonObject.Settings.SaveLogFiles.Detailed) {
        Add-SchemaError "Property 'Settings.SaveLogFiles.Detailed' not found"
    }
    elseif ($JsonObject.Settings.SaveLogFiles.Detailed -isnot [bool] ) {
        Add-SchemaError 'Settings.SaveLogFiles.Detailed must be boolean'
    }

    if ($null -eq $JsonObject.Settings.SendMail) {
        Add-SchemaError "Property 'Settings.SendMail' not found"
    }
    else {
        if (-not $JsonObject.Settings.SendMail.From) {
            Add-SchemaError "Property 'Settings.SendMail.From' not found"
        }

        if ($JsonObject.Settings.SendMail.To -and
            ($JsonObject.Settings.SendMail.To -isnot [string] -and
            $JsonObject.Settings.SendMail.To -isnot [array])) {
            Add-SchemaError "Property 'Settings.SendMail.To' not found"
        }

        if ($null -eq $JsonObject.Settings.SendMail.Body) {
            Add-SchemaError "Property 'Settings.SendMail.Body' not found"
        }
    }

    # --- 3. Validate Matrix structure ---
    if ($null -ne $JsonObject.Matrix) {
        if (-not $JsonObject.Matrix.FolderPath) {
            Add-SchemaError "Property 'Matrix.FolderPath' not found"
        }

        if (-not $JsonObject.Matrix.DefaultsFile) {
            Add-SchemaError "Property 'Matrix.DefaultsFile' not found"
        }

        if ($JsonObject.Matrix.ExcludedSamAccountName -and
            $JsonObject.Matrix.ExcludedSamAccountName -isnot [array]) {
            Add-SchemaError "Property 'Matrix.ExcludedSamAccountName' must be an array"
        }

        if ($null -eq $JsonObject.Matrix.Archive) {
            Add-SchemaError "Property 'Matrix.Archive' not found"
        }
        elseif ($JsonObject.Matrix.Archive -isnot [bool]) {
            Add-SchemaError 'Matrix.Archive must be boolean'
        }
    }
    else {
        Add-SchemaError 'Matrix required'
    }

    # --- 4. Validate MaxConcurrent structure ---
    if ($null -ne $JsonObject.MaxConcurrent) {
        foreach (
            $prop in
            @('Computers', 'FoldersPerMatrix', 'JobsPerRemoteComputer')
        ) {
            $val = $JsonObject.MaxConcurrent.$prop

            if ($null -eq $val) {
                Add-SchemaError "Property 'MaxConcurrent.$prop' not found"
                continue
            }

            if ($val -notmatch '^\d+$') {
                Add-SchemaError "MaxConcurrent.$prop must be an integer"
            }
        }
    }
     
    # --- 5. Validate Export structure (file extensions only) ---
    if ($null -ne $JsonObject.Export) {
        if ($JsonObject.Export.PermissionsExcelFile -and
            ($JsonObject.Export.PermissionsExcelFile -isnot [string]) -and
            ($JsonObject.Export.PermissionsExcelFile -notmatch '\.xlsx$')) {
            Add-SchemaError 'Export.PermissionsExcelFile must be a string ending with .xlsx'
        }

        if ($JsonObject.Export.OverviewHtmlFile -and
            ($JsonObject.Export.OverviewHtmlFile -isnot [string]) -and
            ($JsonObject.Export.OverviewHtmlFile -notmatch '\.html?$')) {
            Add-SchemaError 'Export.OverviewHtmlFile must be a string ending with .html'
        }

        if ($JsonObject.Export.ServiceNowFormDataExcelFile -and
            ($JsonObject.Export.ServiceNowFormDataExcelFile -isnot [string]) -and
            ($JsonObject.Export.ServiceNowFormDataExcelFile -notmatch '\.xlsx$')) {
            Add-SchemaError 'Export.ServiceNowFormDataExcelFile must be a string ending with .xlsx'
        }

        if ($JsonObject.Export.ServiceNowFormDataExcelFile -and
            -not $JsonObject.ServiceNow) {
            Add-SchemaError 'ServiceNow must be defined when ServiceNowFormDataExcelFile is used'
        }
    }
    else {
        Add-SchemaError 'Export required'
    }

    return $errors
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
