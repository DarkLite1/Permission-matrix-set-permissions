function Build-AccessList {
    param(
        [array]$SamAccountNames, [hashtable]$AdObjectHash, [string]$FileName
    )
    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($S in $SamAccountNames) {
        $adData = $AdObjectHash[$S]
        if (-not $adData?.adObject) { continue }
            
        if (-not $adData.adGroupMember) {
            $list.Add(
                [PSCustomObject]@{ 
                    MatrixFileName       = $FileName
                    SamAccountName       = $S
                    Name                 = $adData.adObject.Name
                    Type                 = $adData.adObject.ObjectClass
                    MemberName           = $null 
                    MemberSamAccountName = $null 
                }
            )
        }
        else {
            foreach ($member in $adData.adGroupMember) {
                $list.Add(
                    [PSCustomObject]@{
                        MatrixFileName       = $FileName 
                        SamAccountName       = $S 
                        Name                 = $adData.adObject.Name
                        Type                 = $adData.adObject.ObjectClass
                        MemberName           = $member.Name 
                        MemberSamAccountName = $member.SamAccountName 
                    }
                )
            }
        }
    }
    return $list
}
function Build-ExportData {
    param(
        [Parameter(Mandatory)][array]$ImportedMatrix,
        [hashtable]$AdObjectHash,
        [hashtable]$GroupManagerHash
    )

       
    # Start with an empty result
    $export = @{}

    # Temp lists only if needed
    $accessList = [System.Collections.Generic.List[object]]::new()
    $adObjects = [System.Collections.Generic.List[object]]::new()
    $formData = [System.Collections.Generic.List[object]]::new()
    $groupManagers = [System.Collections.Generic.List[object]]::new()

    function Get-UniqueMatrixSamAccountNames {
        param([object]$MatrixItem)

        return $MatrixItem.Settings?.AdObjects?.Values |
        ForEach-Object { "$($_.SamAccountName)".Trim() } |
        Where-Object { $_ } |
        Sort-Object -Unique
    }

    function Convert-AdObjectExport {
        param([string]$MatrixName, [object]$Entry)

        return [PSCustomObject]@{
            MatrixFileName = $MatrixName
            SamAccountName = $Entry.SamAccountName
            GroupName      = $Entry.Converted.Begin
            SiteCode       = $Entry.Converted.Middle
            Name           = $Entry.Converted.End
        }
    }

    foreach ($Matrix in $ImportedMatrix) {

        $matrixName = $Matrix.File.Item.BaseName
        $uniqueSams = Get-UniqueMatrixSamAccountNames -MatrixItem $Matrix


        # 1. AccessList
        $access = Build-AccessList `
            -SamAccountNames $uniqueSams `
            -AdObjectHash $AdObjectHash `
            -FileName $matrixName

        if ($access) { $accessList.AddRange($access) }


        # 2. GroupManagers
        $gm = Build-GroupManagerList `
            -SamAccountNames $uniqueSams `
            -AdObjectHash $AdObjectHash `
            -GroupManagerHash $GroupManagerHash `
            -FileName $matrixName

        if ($gm) { $groupManagers.AddRange($gm) }


        # 3. AD Objects
        if ($Matrix.Settings?.AdObjects) {
            $adConverted =
            $Matrix.Settings.AdObjects.GetEnumerator() |
            ForEach-Object {
                Convert-AdObjectExport -MatrixName $matrixName -Entry $_.Value
            } |
            Group-Object SamAccountName |
            ForEach-Object { $_.Group[0] }

            if ($adConverted) {
                $adObjects.AddRange($adConverted)
            }
        }


        # 4. FormData
        if ($Matrix.FormData?.Import) {
            $formData.AddRange($Matrix.FormData.Import)
        }
    }

    #
    # Add only non-empty collections
    #
    if ($accessList.Count -gt 0) { $export.AccessList = $accessList }
    if ($adObjects.Count -gt 0) { $export.AdObjects = $adObjects }
    if ($formData.Count -gt 0) { $export.FormData = $formData }
    if ($groupManagers.Count -gt 0) { $export.GroupManagers = $groupManagers }

    #
    # If nothing was added → return $null
    #
    if ($export.Count -eq 0) { return $null }

    return $export
}
function Build-GroupManagerList {
    param([array]$SamAccountNames, [hashtable]$AdObjectHash, [hashtable]$GroupManagerHash, [string]$FileName)
    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($S in $SamAccountNames) {
        $adData = $AdObjectHash[$S]
        if (-not $adData?.adObject -or $adData.adObject.ObjectClass -ne 'group') { continue }
            
        $managedBy = $adData.adObject.PSObject.Properties['ManagedBy']?.Value
        if ([string]::IsNullOrWhiteSpace($managedBy)) { continue }

        $gm = $GroupManagerHash[$managedBy]
        if (-not $gm?.adObject) { 
            $list.Add(
                [PSCustomObject]@{ 
                    MatrixFileName    = $FileName 
                    GroupName         = $adData.adObject.Name 
                    ManagerName       = $null
                    ManagerType       = $null 
                    ManagerMemberName = $null 
                }
            )
        }
        elseif (-not $gm.adGroupMember) { 
            $list.Add(
                [PSCustomObject]@{ 
                    MatrixFileName    = $FileName
                    GroupName         = $adData.adObject.Name 
                    ManagerName       = $gm.adObject.Name 
                    ManagerType       = $gm.adObject.ObjectClass 
                    ManagerMemberName = $null 
                }
            )
        }
        else { 
            foreach ($user in $gm.adGroupMember) { 
                $list.Add(
                    [PSCustomObject]@{ 
                        MatrixFileName    = $FileName 
                        GroupName         = $adData.adObject.Name 
                        ManagerName       = $gm.adObject.Name
                        ManagerType       = $gm.adObject.ObjectClass
                        ManagerMemberName = $user.Name 
                    }
                ) 
            } 
        }
    }
    return $list
}
function Export-Files {
    param(
        [Parameter(Mandatory)][object]$DataToExport,
        [Parameter(Mandatory)][object]$ExportConfig,
        [Parameter(Mandatory)][object]$ServiceNowConfig,
        [Parameter(Mandatory)][string]$ExportLogFolder,
        [Parameter(Mandatory)][object]$ScriptPathItem,
        [Parameter(Mandatory)][ref]$SystemErrors
    )

    #
    # Return structure
    #
    $results = @{}


    #
    # Helper: run a task safely and track failures
    #
    function Invoke-Safe {
        param(
            [scriptblock]$Action,
            [string]$ErrorMessage
        )
        try { & $Action }
        catch { 
            $SystemErrors.Value.Add(
                [PSCustomObject]@{
                    DateTime = Get-Date
                    Message  = "$ErrorMessage`: $_"
                }
            )
            return $false
        }
        return $true
    }


    #
    # 1. Export Permissions Excel
    #
    if ($ExportConfig.PermissionsExcelFile) {

        $ok = Invoke-Safe `
            -ErrorMessage 'Export-PermissionsFile failed' `
            -Action {
            Export-PermissionsFile `
                -DataToExport $DataToExport `
                -OutputPath $ExportConfig.PermissionsExcelFile `
                -LogFolder $ExportLogFolder `
                -SystemErrors $SystemErrors
        }

        if ($ok -and (Test-Path -LiteralPath $ExportConfig.PermissionsExcelFile -PathType Leaf)) {
            $results['PermissionsExcelFile'] = $ExportConfig.PermissionsExcelFile
        }
    }


    #
    # 2. Export + Upload ServiceNow form data
    #
    if ($ExportConfig.ServiceNowFormDataExcelFile -and $DataToExport.FormData) {

        $hasData = Invoke-Safe `
            -ErrorMessage 'Export-ServiceNowFormData failed' `
            -Action {
            Export-ServiceNowFormData `
                -DataToExport $DataToExport `
                -OutputPath $ExportConfig.ServiceNowFormDataExcelFile `
                -ExportLogFolder $ExportLogFolder `
                -SystemErrors $SystemErrors
        }

        if ($hasData -and (Test-Path -LiteralPath $ExportConfig.ServiceNowFormDataExcelFile -PathType Leaf)) {

            Invoke-Safe `
                -ErrorMessage 'Upload-ServiceNowFormData failed' `
                -Action {
                Upload-ServiceNowFormData `
                    -OutputPath $ExportConfig.ServiceNowFormDataExcelFile `
                    -ServiceNowConfig $ServiceNowConfig `
                    -ScriptPathItem $ScriptPathItem `
                    -SystemErrors $SystemErrors
            }

            $results['ServiceNowFormDataExcelFile'] = $ExportConfig.ServiceNowFormDataExcelFile
        }
    }


    #
    # 3. Export Overview HTML
    #
    if ($ExportConfig.OverviewHtmlFile -and $DataToExport.FormData) {

        $ok = Invoke-Safe `
            -ErrorMessage 'Export-OverviewHtml failed' `
            -Action {
            Export-OverviewHtml `
                -DataToExport $DataToExport `
                -OutputPath $ExportConfig.OverviewHtmlFile `
                -ExportLogFolder $ExportLogFolder `
                -SystemErrors $SystemErrors
        }

        if ($ok -and (Test-Path -LiteralPath $ExportConfig.OverviewHtmlFile -PathType Leaf)) {
            $results['OverviewHtmlFile'] = $ExportConfig.OverviewHtmlFile
        }
    }


    return $results
}
function Export-OverviewHtml {
    param(
        [Parameter(Mandatory)][object]$DataToExport,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$ExportLogFolder,
        [Parameter(Mandatory)][ref]$SystemErrors
    )

    try {
        #
        # 1. Remove old output file
        #
        Remove-FileHC -FilePath $OutputPath

        #
        # 2. Build table rows
        #
        function New-OverviewRow {
            param([object]$FormData)

            $category = $FormData.MatrixCategoryName
            $subcat = $FormData.MatrixSubCategoryName
            $folderPath = $FormData.MatrixFolderDisplayName
            $filePath = $FormData.MatrixFilePath
            $fileName = $FormData.MatrixFileName

            # Build mailto: list safely
            $emails = ($FormData.MatrixResponsible -split ',') |
            ForEach-Object {
                $trimmed = $_.Trim()
                if ($trimmed) { "mailto:$trimmed$trimmed</a>" }
            } |
            Join-String -Separator ', '

            return @"
<tr>
    <td>$category</td>
    <td>$subcat</td>
    <td>$folderPath$folderPath</a></td>
    <td>$filePath$fileName</a></td>
    <td>$emails</td>
</tr>
"@
        }

        $rows = $DataToExport.FormData |
        Sort-Object MatrixCategoryName, MatrixSubCategoryName, MatrixFolderDisplayName |
        ForEach-Object { New-OverviewRow -FormData $_ } |
        Join-String

        #
        # 3. Build full HTML document
        #
        $html = @"
<html>
<head>
<style>
    body { font-family:Arial; }
    table { width:100%; border-collapse:collapse; }
    th, td { padding:10px; border-bottom:1px solid #ddd; text-align:left; }
</style>
</head>
<body>
<h1>Matrix files overview</h1>

<table>
    <tr>
        <th>Category</th>
        <th>Subcategory</th>
        <th>Folder</th>
        <th>Link to matrix</th>
        <th>Responsible</th>
    </tr>
    $rows
</table>

</body>
</html>
"@

        #
        # 4. Save HTML output
        #
        $html | Out-File -LiteralPath $OutputPath -Encoding UTF8 -Force

        #
        # 5. Copy to log folder
        #
        $logCopy = Join-Path $ExportLogFolder 'Overview.html'
        Copy-Item -LiteralPath $OutputPath -Destination $logCopy -Force
    }
    catch {
        $SystemErrors.Value.Add(
            [PSCustomObject]@{
                DateTime = Get-Date
                Message  = "Export-OverviewHtml failed: $_"
            }
        )
    }
}
function Export-PermissionsFile {
    param(
        [Parameter(Mandatory)][object]$DataToExport,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$LogFolder,
        [Parameter(Mandatory)][ref]$SystemErrors
    )

    try {
        #
        # 1. Remove existing output files
        #
        Remove-FileHC -FilePath $OutputPath

        $logTempPath = Join-Path $LogFolder 'Permissions.xlsx'
        Remove-FileHC -FilePath $logTempPath

        #
        # 2. Export each collection into its corresponding worksheet
        #
        foreach ($entry in $DataToExport.GetEnumerator()) {

            if (-not $entry.Value) {
                continue
            }

            $params = @{
                Path          = $logTempPath
                WorksheetName = $entry.Name
                TableName     = $entry.Name
                AutoSize      = $true
                FreezeTopRow  = $true
            }

            try {
                $entry.Value | Export-Excel @params
            }
            catch {
                $SystemErrors.Value.Add(
                    [PSCustomObject]@{
                        DateTime = Get-Date
                        Message  = "Export-PermissionsFile failed for sheet '$($entry.Name)': $_"
                    }
                )
            }
        }

        #
        # 3. Copy the final result to the target output path
        #
        if (Test-Path -LiteralPath $logTempPath -PathType Leaf) {
            Copy-Item -LiteralPath $logTempPath -Destination $OutputPath -Force
        }
    }
    catch {
        $SystemErrors.Value.Add(
            [PSCustomObject]@{
                DateTime = Get-Date
                Message  = "Export-PermissionsFile failed: $_"
            }
        )
    }
}
function Export-ServiceNowFormData {
    param(
        [Parameter(Mandatory)][object]$DataToExport,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$ExportLogFolder,
        [Parameter(Mandatory)][ref]$SystemErrors
    )

    try {
        #
        # 1. Remove existing output file (if any)
        #
        Remove-FileHC -FilePath $OutputPath

        #
        # 2. Build a lookup of FormData keyed by MatrixFileName
        #
        $formDataHash = @{}
        foreach ($fd in $DataToExport.FormData) {
            if ($fd.MatrixFileName) {
                $formDataHash[$fd.MatrixFileName] = $fd
            }
        }

        #
        # 3. Build ServiceNow export rows
        #
        $serviceNowRows = foreach ($adObj in $DataToExport.AdObjects) {

            $fd = $formDataHash[$adObj.MatrixFileName]

            if (
                $fd -and
                $fd.MatrixFormStatus -eq 'Enabled'
            ) {
                [PSCustomObject]@{
                    u_matrixfilename        = $adObj.MatrixFileName
                    u_matrixfolderpath      = $fd.MatrixFolderPath
                    u_matrixcategoryname    = $fd.MatrixCategoryName
                    u_matrixsubcategoryname = $fd.MatrixSubCategoryName
                    u_matrixresponsible     = $fd.MatrixResponsible
                    u_adobjectname          = $adObj.SamAccountName
                }
            }
        }

        #
        # 4. Nothing to export?
        #
        if (-not $serviceNowRows) {
            return $false
        }

        #
        # 5. Export to Excel (SnowFormData sheet)
        #
        $xlsxParams = @{
            Path          = $OutputPath
            WorksheetName = 'SnowFormData'
            TableName     = 'SnowFormData'
            AutoSize      = $true
            FreezeTopRow  = $true
        }

        $serviceNowRows | Export-Excel @xlsxParams

        #
        # 6. Copy file to export log folder
        #
        $logCopyPath = Join-Path $ExportLogFolder 'ServiceNowFormData.xlsx'

        if (Test-Path -LiteralPath $OutputPath) {
            Copy-Item -LiteralPath $OutputPath -Destination $logCopyPath -Force
            return $true
        }

        return $false
    }
    catch {
        $SystemErrors.Value.Add(
            [PSCustomObject]@{
                DateTime = Get-Date
                Message  = "Export-ServiceNowFormData failed: $_"
            }
        )
        return $false
    }
}
function Upload-ServiceNowFormData {
    param(
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][object]$ServiceNowConfig,
        [Parameter(Mandatory)][object]$ScriptPathItem,
        [Parameter(Mandatory)][ref]$SystemErrors
    )

    try {
        #
        # 1. Validate required ServiceNow parameters
        #
        $credPath = $ServiceNowConfig.CredentialsFilePath
        $env = $ServiceNowConfig.Environment
        $table = $ServiceNowConfig.TableName

        if (-not $credPath -or -not $env -or -not $table) {
            return   # Silent skip 
        }

        #
        # 2. Execute uploader script
        #
        & $ScriptPathItem.UpdateServiceNow `
            -CredentialsFilePath $credPath `
            -Environment $env `
            -TableName $table `
            -FormDataExcelFilePath $OutputPath `
            -ExcelFileWorksheetName 'SnowFormData'
    }
    catch {
        $SystemErrors.Value.Add(
            [PSCustomObject]@{
                DateTime = Get-Date
                Message  = "Upload-ServiceNowFormData failed: $_"
            }
        )
    }
}