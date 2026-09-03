function Build-ExportDataHC {
    <#
    .SYNOPSIS
        Builds aggregated export data for permissions and ServiceNow form data.

    .DESCRIPTION
        Iterates the flattened matrix objects ($Context.AllMatrices) and emits
        one Permissions row per matrix object, plus one FormData row and its
        ServiceNow rows per matrix FILE.

        Because AllMatrices is flattened per Settings row, several matrix
        objects can share one source file. FormData is therefore deduplicated on
        the FileContext (keyed on Item.FullName, falling back to Item.Name)
        while permissions rows are emitted for every matrix object.

    .NOTES
        - The dedupe means the ServiceNow rows are built from whichever matrix
          object reached the file FIRST; its Matrix.AdNames alone supply the AD
          objects for that file.
        - A file whose FormData is $null produces no FormData and no ServiceNow
          rows, but is still marked as seen.
        - The FormData row is MUTATED: 'MatrixFileName' is added to the file's
          formatted FormData in place.

    .PARAMETER ImportedMatrix
        Matrix objects, each exposing Setting.Formatted.{ComputerName, Path,
        Action}, Check, FileContext.Item.{Name, FullName} and
        FileContext.Sheets.FormData.Formatted.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$ImportedMatrix,
        [string[]]$AdGroupPlaceHolders = @()
    )

    $permissionsRows = [System.Collections.Generic.List[pscustomobject]]::new()
    $formDataRows = [System.Collections.Generic.List[pscustomobject]]::new()
    $serviceNowData = [System.Collections.Generic.List[pscustomobject]]::new()

    # Tracks the files whose FormData row has already been emitted, so a file
    # with several enabled Settings rows still yields a single FormData row
    $seenFiles = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )

    foreach ($matrixObj in $ImportedMatrix) {

        $fileContext = $matrixObj.FileContext

        #region Permissions row (one per matrix object / enabled Settings row)
        if ($matrixObj.Setting) {
            $setting = $matrixObj.Setting.Formatted

            $permissionsRows.Add(
                [pscustomobject]@{
                    MatrixFile = $fileContext.Item.Name
                    Computer   = $setting.ComputerName
                    Path       = $setting.Path
                    Action     = $setting.Action
                    Errors     = @(
                        $matrixObj.Check |
                        Where-Object { $_.Type -eq 'FatalError' }).Count
                    Incorrect  = @(
                        $matrixObj.Check |
                        Where-Object { $_.Type -eq 'Incorrect' }).Count
                    Warnings   = @(
                        $matrixObj.Check |
                        Where-Object { $_.Type -eq 'Warning' }).Count
                    Fixed      = @(
                        $matrixObj.Check |
                        Where-Object { $_.Type -eq 'Fixed' }).Count
                }
            )
        }
        #endregion

        #region FormData row (one per file, deduplicated on the FileContext)
        if ($fileContext) {
            $fileKey = $fileContext.Item.FullName
            if (-not $fileKey) { $fileKey = $fileContext.Item.Name }

            if ($fileKey -and $seenFiles.Add($fileKey)) {
                $formData = $fileContext.Sheets.FormData.Formatted
                if ($formData) {
                    $formData | Add-Member -NotePropertyMembers @{
                        MatrixFileName = $fileContext.Item.Name
                    } -Force
                    
                    $formDataRows.Add([pscustomobject]$formData)

                    #region Create ServiceNow upload data
                    $adObjects = @(
                        $matrixObj.Matrix.AdNames.Values | 
                        Sort-Object -Unique
                    )

                    $emailsResponsible = (
                        Resolve-ResponsibleEmailHC `
                            -Responsible $formData.MatrixResponsible `
                            -AdGroupPlaceHolders $AdGroupPlaceHolders
                    ).Emails -join ','

                    foreach ($adObject in $adObjects) {
                        $serviceNowData.Add(
                            [pscustomobject]@{
                                u_matrixfilename        = $formData.MatrixFileName
                                u_matrixfolderpath      = $formData.MatrixFolderPath
                                u_matrixcategoryname    = $formData.MatrixCategoryName
                                u_matrixsubcategoryname = $formData.MatrixSubCategoryName
                                u_matrixresponsible     = $emailsResponsible
                                u_adobjectname          = $adObject
                            }
                        )
                    }
                    #endregion
                }
            }
        }
        #endregion
    }

    return [pscustomobject]@{
        Permissions    = $permissionsRows.ToArray()
        FormData       = $formDataRows.ToArray()
        ServiceNowData = $serviceNowData.ToArray()
    }
}

function Export-FilesHC {
    <#
    .SYNOPSIS
        Executes all export operations based on settings.

    .DESCRIPTION
        Writes the configured export artifacts to disk, each one skipped when
        its setting is absent:

        - Permissions Excel: one consolidated workbook holding the 'AccessList',
          'GroupManagers', 'AdObjects' and 'FormData' worksheets aggregated
          across every matrix file.
        - ServiceNow FormData Excel.
        - The standalone overview HTML page.

        The email summary body is a separate artifact built by EndHC and is not
        used here.

    .NOTES
        EndHC also writes the same per-file rows into the per-matrix copies in
        the log folder. To avoid resolving group managers in AD twice,
        Build-ConsolidatedExportDataHC caches the rows built here on each file
        result (.LogSheets) for EndHC to reuse.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$ImportedMatrix,
        [Parameter(Mandatory)]      $ExportSettings,
        [array]$FileResults = @(),
        [array]$AdObjectDetails = @(),
        [string[]]$AdGroupPlaceHolders = @()
    )

    $exportData = Build-ExportDataHC `
        -ImportedMatrix $ImportedMatrix `
        -AdGroupPlaceHolders $AdGroupPlaceHolders

    $results = [ordered]@{
        Permissions  = $null
        FormData     = $null
        OverviewHtml = $null
    }

    # 1. Consolidated Permissions workbook
    #    (AccessList / GroupManagers / AdObjects / FormData)
    if ($ExportSettings.PermissionsExcelFile) {
        $consolidated = Build-ConsolidatedExportDataHC `
            -FileResults $FileResults `
            -AdObjectDetails $AdObjectDetails

        $results.Permissions = Export-ConsolidatedPermissionsFileHC `
            -AccessList $consolidated.AccessList `
            -GroupManagers $consolidated.GroupManagers `
            -AdObjects $consolidated.AdObjects `
            -FormData $consolidated.FormData `
            -AdGroupPlaceHolders $AdGroupPlaceHolders `
            -Path $ExportSettings.PermissionsExcelFile
    }

    # 2. ServiceNow FormData Excel
    if ($ExportSettings.ServiceNowFormDataExcelFile) {
        $results.FormData = Export-ServiceNowFormDataHC `
            -FormDataRows $exportData.FormData `
            -ServiceNowDataRows $exportData.ServiceNowData `
            -Path $ExportSettings.ServiceNowFormDataExcelFile
    }

    # 3. Overview HTML (built from FormData rows; independent of the email body)
    if ($ExportSettings.OverviewHtmlFile) {
        $html = New-OverviewHtmlHC -FormData $exportData.FormData
        $results.OverviewHtml = Export-OverviewHtmlHC `
            -Html $html `
            -Path $ExportSettings.OverviewHtmlFile
    }

    return $results
}

function Export-ServiceNowFormDataHC {
    <#
        Writes ServiceNow FormData into an Excel file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$FormDataRows,
        [Parameter(Mandatory)][array]$ServiceNowDataRows,
        [Parameter(Mandatory)][string]$Path
    )

    try {
        $params = @{
            Path     = $Path
            AutoSize = $true
        }

        $FormDataRows | 
        Export-Excel @params -WorksheetName 'FormData'-TableName 'FormData'

        $ServiceNowDataRows | 
        Export-Excel @params -WorksheetName 'ServiceNowData'-TableName 'ServiceNowData'

        return $Path
    }
    catch {
        throw "Failed exporting ServiceNow FormData Excel: $_"
    }
}

function Export-OverviewHtmlHC {
    <#
        Writes the generated HTML overview page to a file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Html,
        [Parameter(Mandatory)][string]$Path
    )

    try {
        $Html | Out-File -LiteralPath $Path -Encoding utf8 -Force
        return $Path
    }
    catch {
        throw "Failed exporting Overview HTML file: $_"
    }
}

function Build-ConsolidatedExportDataHC {
    <#
    .SYNOPSIS
        Aggregates the log-sheet rows of every matrix file into one set of rows
        for the consolidated Permissions workbook.

    .DESCRIPTION
        For each file result, builds the per-file 'AccessList', 'GroupManagers'
        and 'AdObjects' rows with Build-MatrixLogSheetRowsHC and combines them:

        - AccessList    : prefixed with a 'MatrixFileName' column so rows from
                          different files can be told apart
        - GroupManagers : likewise prefixed; 'MemberEnabled' is preserved
        - AdObjects     : taken as-is (already carries 'MatrixFileName')
        - FormData      : one row per file, from the file's formatted FormData

    .NOTES
        The unmodified per-file row sets are cached on each file result as a
        'LogSheets' property, so EndHC can reuse them for the per-matrix log
        folder copy instead of resolving group managers in AD a second time.

    .OUTPUTS
        PSCustomObject with 'AccessList', 'GroupManagers', 'AdObjects' and
        'FormData'.
    #>
    [CmdletBinding()]
    param(
        [array]$FileResults = @(),
        [array]$AdObjectDetails = @()
    )

    $accessListRows = [System.Collections.Generic.List[pscustomobject]]::new()
    $groupManagerRows = [System.Collections.Generic.List[pscustomobject]]::new()
    $adObjectRows = [System.Collections.Generic.List[pscustomobject]]::new()
    $formDataRows = [System.Collections.Generic.List[pscustomobject]]::new()

    foreach ($fileResult in $FileResults) {

        # Reuse the cached per-file rows when present, otherwise build them
        # once and cache them so EndHC's log-folder copy can reuse them
        $logSheets = if (
            $fileResult.PSObject.Properties['LogSheets'] -and
            $fileResult.LogSheets
        ) {
            $fileResult.LogSheets
        }
        else {
            $sheets = Build-MatrixLogSheetRowsHC `
                -FileResult $fileResult `
                -AdObjectDetails $AdObjectDetails

            $fileResult | Add-Member `
                -NotePropertyName 'LogSheets' `
                -NotePropertyValue $sheets -Force

            $sheets
        }

        $matrixFileName = $fileResult.Item.Name

        foreach ($row in $logSheets.AccessList) {
            # Prefix the file name so rows from different files are
            # distinguishable in the combined sheet
            $accessListRows.Add(
                [pscustomobject]@{
                    MatrixFileName       = $matrixFileName
                    SamAccountName       = $row.SamAccountName
                    Name                 = $row.Name
                    Type                 = $row.Type
                    MemberName           = $row.MemberName
                    MemberSamAccountName = $row.MemberSamAccountName
                    MemberEnabled        = $row.MemberEnabled
                }
            )
        }

        foreach ($row in $logSheets.GroupManagers) {
            # Prefix the file name; keep MemberEnabled
            $groupManagerRows.Add(
                [pscustomobject]@{
                    MatrixFileName    = $matrixFileName
                    GroupName         = $row.GroupName
                    ManagerName       = $row.ManagerName
                    ManagerType       = $row.ManagerType
                    ManagerMemberName = $row.ManagerMemberName
                    MemberEnabled     = $row.MemberEnabled
                }
            )
        }

        foreach ($row in $logSheets.AdObjects) {
            $adObjectRows.Add($row)
        }

        $formData = $fileResult.Sheets.FormData.Formatted
        if ($formData) {
            $formDataRows.Add([pscustomobject]$formData)
        }
    }

    return [pscustomobject]@{
        AccessList    = $accessListRows.ToArray()
        GroupManagers = $groupManagerRows.ToArray()
        AdObjects     = $adObjectRows.ToArray()
        FormData      = $formDataRows.ToArray()
    }
}

function Get-PlaceHolderFilterValueHC {
    <#
    .SYNOPSIS
        Builds the value list used to filter placeholder accounts out of the
        matrix Excel log file.

    .DESCRIPTION
        Placeholder accounts are configured as SamAccountNames
        ('Matrix.AdGroupPlaceHolders', in both the main and the audit report
        configuration). 'AccessList' has a
        'MemberSamAccountName' column and can be matched directly, but
        'GroupManagers' only carries a display name in 'ManagerMemberName'.

        AccessList rows already pair both spellings of every member, so this
        walks them once to translate each placeholder SamAccountName into its
        display name and returns both. A value occurring in neither column never
        matches, so the combined list can be handed to Set-DefaultSheetFilterHC
        for both worksheets at once.

    .NOTES
        Returns an empty array when no placeholders are configured, without
        reading AccessListRow at all.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [string[]]$AdGroupPlaceHolders = @(),
        [array]$AccessListRow = @()
    )

    $samAccountName = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]@($AdGroupPlaceHolders | Where-Object { $_ }),
        [System.StringComparer]::OrdinalIgnoreCase
    )

    if ($samAccountName.Count -eq 0) { return @() }

    # Kept separate from the lookup set so a resolved display name can never
    # be treated as a placeholder SamAccountName on a later row
    $result = [System.Collections.Generic.HashSet[string]]::new(
        $samAccountName,
        [System.StringComparer]::OrdinalIgnoreCase
    )

    foreach ($row in $AccessListRow) {
        if (-not $row) { continue }

        $sam = [string]$row.MemberSamAccountName

        if ((-not $sam) -or (-not $samAccountName.Contains($sam.Trim()))) {
            continue
        }

        $name = [string]$row.MemberName

        if ($name -and $name.Trim()) { $null = $result.Add($name.Trim()) }
    }

    return [string[]]$result
}

function Export-ConsolidatedPermissionsFileHC {
    <#
    .SYNOPSIS
        Writes the consolidated Permissions workbook with the 'AccessList',
        'GroupManagers', 'AdObjects' and 'FormData' worksheets.

    .DESCRIPTION
        Always creates all four worksheets, even when a row set is empty
        (header-only for the sheets with a fixed column layout), so the workbook
        always has the same structure. Any pre-existing file at the target path
        is replaced, so re-runs don't stack stale worksheets.

    .NOTES
        FormData columns depend on the matrix template, so no fixed headers are
        written when that set is empty.

    .OUTPUTS
        System.String - the path that was written.
    #>
    [CmdletBinding()]
    param(
        [array]$AccessList = @(),
        [array]$GroupManagers = @(),
        [array]$AdObjects = @(),
        [array]$FormData = @(),
        [string[]]$AdGroupPlaceHolders = @(),
        [Parameter(Mandatory)][string]$Path
    )

    try {
        # Start from a clean file so re-runs don't stack stale worksheets
        if (Test-Path -LiteralPath $Path) {
            Remove-Item -LiteralPath $Path -Force -ErrorAction Stop
        }

        $worksheets = @(
            @{
                Name    = 'AccessList'
                Rows    = $AccessList
                Headers = @(
                    'MatrixFileName', 'SamAccountName', 'Name', 'Type',
                    'MemberName', 'MemberSamAccountName', 'MemberEnabled'
                )
            }
            @{
                Name    = 'GroupManagers'
                Rows    = $GroupManagers
                Headers = @(
                    'MatrixFileName', 'GroupName', 'ManagerName',
                    'ManagerType', 'ManagerMemberName', 'MemberEnabled'
                )
            }
            @{
                Name    = 'AdObjects'
                Rows    = $AdObjects
                Headers = @(
                    'MatrixFileName', 'SamAccountName',
                    'GroupName', 'SiteCode', 'Name', 'Enabled'
                )
            }
            @{
                Name    = 'FormData'
                Rows    = $FormData
                Headers = @()
            }
        )

        foreach ($ws in $worksheets) {
            if ($ws.Rows -and @($ws.Rows).Count -gt 0) {
                $ws.Rows | Export-Excel -Path $Path `
                    -WorksheetName $ws.Name -TableName $ws.Name `
                    -AutoSize -FreezeTopRow
            }
            else {
                # Always create the worksheet, even without data, so the
                # workbook structure stays stable across runs
                $excelPackage = Open-ExcelPackage -Path $Path -Create
                try {
                    $sheet = Add-Worksheet `
                        -ExcelPackage $excelPackage `
                        -WorksheetName $ws.Name

                    for ($i = 0; $i -lt $ws.Headers.Count; $i++) {
                        $sheet.Cells[1, ($i + 1)].Value = $ws.Headers[$i]
                        $sheet.Cells[1, ($i + 1)].Style.Font.Bold = $true
                    }
                }
                finally {
                    Close-ExcelPackage -ExcelPackage $excelPackage
                }
            }
        }

        $placeHolderValue = Get-PlaceHolderFilterValueHC `
            -AdGroupPlaceHolders $AdGroupPlaceHolders `
            -AccessListRow $AccessList

        Set-DefaultSheetFilterHC -Path $Path `
            -WorksheetName 'AccessList', 'GroupManagers' `
            -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank `
            -ExcludeColumnName 'MemberSamAccountName', 'ManagerMemberName' `
            -ExcludeValue $placeHolderValue

        return $Path
    }
    catch {
        throw "Failed exporting consolidated Permissions Excel file '$Path': $_"
    }
}

function Copy-MatrixFileToLogFolderHC {
    <#
    .SYNOPSIS
        Copies the original matrix Excel file to the log folder and appends the
        worksheets 'AccessList', 'GroupManagers' and 'AdObjects'.

    .DESCRIPTION
        Copies the processed source .xlsx into the dated log folder of that
        matrix file. The copy always contains the three extra worksheets, even
        when no rows are available (header row only), so every archived matrix
        file has the same structure.

        Expected row shapes (extra properties become extra columns):
        - AccessList   : SamAccountName, Name, Type, MemberName,
                         MemberSamAccountName, MemberEnabled
        - GroupManagers: GroupName, ManagerName, ManagerType,
                         ManagerMemberName, MemberEnabled
        - AdObjects    : MatrixFileName, SamAccountName, GroupName, SiteCode,
                         Name, Enabled

    .PARAMETER DestinationFileName
        Optional file name (with extension) for the copy. Defaults to the source
        file's own name. Lets callers add a date-stamped name, as the audit
        report's per-run history files do.

    .NOTES
        The read-only attribute is stripped from the copy: source matrix files
        are often stored read-only and ImportExcel cannot write to them.

    .OUTPUTS
        System.String - the absolute path of the created copy.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFilePath,
        [Parameter(Mandatory)][string]$LogFolder,
        [array]$AccessListRows,
        [array]$GroupManagerRows,
        [array]$AdObjectRows,
        [hashtable]$DefaultsAcl,
        [string[]]$AdGroupPlaceHolders = @(),
        [string]$DestinationFileName
    )

    try {
        if (-not (Test-Path -LiteralPath $SourceFilePath -PathType Leaf)) {
            throw "Source matrix file '$SourceFilePath' not found"
        }

        $leaf = if ($DestinationFileName) { $DestinationFileName }
        else { Split-Path -Path $SourceFilePath -Leaf }

        $destinationPath = Join-Path -Path $LogFolder -ChildPath $leaf

        Copy-Item -LiteralPath $SourceFilePath `
            -Destination $destinationPath -Force -ErrorAction Stop

        # Source matrix files are frequently read-only; the copy inherits
        # that attribute and Export-Excel would fail to open the package
        Set-ItemProperty -LiteralPath $destinationPath `
            -Name IsReadOnly -Value $false -ErrorAction SilentlyContinue

        #region Convert hashtable to array
        if (-not $DefaultsAcl) { $DefaultsAcl = @{} }

        $DefaultsAclRows = $DefaultsAcl.GetEnumerator() | Sort-Object Key | ForEach-Object {
            [pscustomobject]@{
                SamAccountName = $_.Key
                Permission     = $_.Value
            }
        }
        #endregion

        $worksheets = @(
            @{
                Name    = 'AccessList'
                Rows    = $AccessListRows
                Headers = @(
                    'SamAccountName', 'Name', 'Type',
                    'MemberName', 'MemberSamAccountName', 'MemberEnabled'
                )
            }
            @{
                Name    = 'GroupManagers'
                Rows    = $GroupManagerRows
                Headers = @(
                    'GroupName', 'ManagerName', 'ManagerType',
                    'ManagerMemberName', 'MemberEnabled'
                )
            }
            @{
                Name    = 'AdObjects'
                Rows    = $AdObjectRows
                Headers = @(
                    'MatrixFileName', 'SamAccountName',
                    'GroupName', 'SiteCode', 'Name', 'Enabled'
                )
            }
            @{
                Name    = 'DefaultPermissions'
                Rows    = $DefaultsAclRows
                Headers = @(
                    'SamAccountName', 'Permission'
                )
            }
        )

        foreach ($ws in $worksheets) {
            if ($ws.Rows -and @($ws.Rows).Count -gt 0) {
                $ws.Rows | Export-Excel -Path $destinationPath `
                    -WorksheetName $ws.Name -TableName $ws.Name `
                    -AutoSize -FreezeTopRow
            }
            else {
                # Always create the worksheet, even without data,
                # so every archived matrix file has the same structure
                $excelPackage = Open-ExcelPackage -Path $destinationPath
                try {
                    $sheet = Add-Worksheet `
                        -ExcelPackage $excelPackage `
                        -WorksheetName $ws.Name

                    for ($i = 0; $i -lt $ws.Headers.Count; $i++) {
                        $sheet.Cells[1, ($i + 1)].Value = $ws.Headers[$i]
                        $sheet.Cells[1, ($i + 1)].Style.Font.Bold = $true
                    }
                }
                finally {
                    Close-ExcelPackage -ExcelPackage $excelPackage
                }
            }
        }

        $placeHolderValue = Get-PlaceHolderFilterValueHC `
            -AdGroupPlaceHolders $AdGroupPlaceHolders `
            -AccessListRow $AccessListRows

        Set-DefaultSheetFilterHC -Path $destinationPath `
            -WorksheetName 'AccessList', 'GroupManagers' `
            -ColumnName 'MemberEnabled' -VisibleValue 'TRUE' -IncludeBlank `
            -ExcludeColumnName 'MemberSamAccountName', 'ManagerMemberName' `
            -ExcludeValue $placeHolderValue

        return $destinationPath
    }
    catch {
        throw "Failed copying matrix file '$SourceFilePath' to log folder '$LogFolder': $_"
    }
}

function Build-MatrixLogSheetRowsHC {
    <#
    .SYNOPSIS
        Builds the row sets for the 'AccessList', 'GroupManagers' and
        'AdObjects' worksheets in the matrix file copy saved to the log folder.

    .DESCRIPTION
        Transforms the resolved AD details of one matrix file into three flat
        row collections for Copy-MatrixFileToLogFolderHC:

        - AccessList   : one row per group member, with 'MemberEnabled' holding
                         the member's AD account status (blank for nested
                         groups). A group without members still gets one row
                         with empty member columns. AD objects of type 'user'
                         used directly in the matrix are listed with themselves
                         as member, so their status is visible too.
        - GroupManagers: one row per group. A manager ('managedBy') is resolved
                         against AD; if the manager is itself a GROUP, one row
                         per manager-group member is written. For a single user
                         manager, 'MemberEnabled' holds that manager's own
                         status, so a disabled managing account is visible.
        - AdObjects    : one row per unique AD object in the file, with
                         'Enabled'. 'GroupName', 'SiteCode' and 'Name' are
                         derived by matching the AD object name against the
                         Settings rows of this file; names that don't follow the
                         naming convention leave these blank.

    .NOTES
        AD object names come from the per-folder 'AdNames' maps built during the
        SID rewrite in the BEGIN stage, falling back to the raw ACL keys when
        that rewrite was skipped. This includes default permissions merged into
        the folder ACLs.

    .PARAMETER AdObjectDetails
        $Context.AdObjectDetails, populated by the BEGIN stage from
        Get-ADObjectDetailHC (objects with 'SamAccountName', 'adObject' and
        'adGroupMember').

    .OUTPUTS
        PSCustomObject with 'AccessList', 'GroupManagers' and 'AdObjects'.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][pscustomobject]$FileResult,
        [array]$AdObjectDetails = @(),
        [int]$MaxThreads = 7
    )

    $accessListRows = [System.Collections.Generic.List[pscustomobject]]::new()
    $groupManagerRows = [System.Collections.Generic.List[pscustomobject]]::new()
    $adObjectRows = [System.Collections.Generic.List[pscustomobject]]::new()

    #region Collect the unique AD object names used in this matrix file
    $fileAdNames = [System.Collections.Generic.List[string]]::new()

    foreach ($matrixObj in $FileResult.Matrices) {
        foreach ($folder in $matrixObj.Matrix) {
            if (
                $folder.PSObject.Properties['AdNames'] -and
                $folder.AdNames.Count
            ) {
                # SID rewrite done: AdNames holds SID -> original name
                $fileAdNames.AddRange([string[]]@($folder.AdNames.Values))
            }
            elseif ($folder.ACL -and $folder.ACL.Count) {
                # SID rewrite skipped (matrix flagged): keys are still names
                $fileAdNames.AddRange([string[]]@($folder.ACL.Keys))
            }
        }
    }

    $uniqueAdNames = @($fileAdNames | Sort-Object -Unique)
    #endregion

    #region Index the resolved AD details by input name
    $detailMap = @{}

    foreach ($detail in $AdObjectDetails) {
        if ($detail.SamAccountName) {
            $detailMap[$detail.SamAccountName] = $detail
        }
    }
    #endregion

    #region Build name part prefixes from the Settings rows
    # Full prefixes ('GroupName SiteCode') are tried before GroupName-only
    # prefixes, longest first, so the most specific match always wins
    $fullPrefixes = [System.Collections.Generic.List[pscustomobject]]::new()
    $groupPrefixes = [System.Collections.Generic.List[string]]::new()

    foreach ($matrixObj in $FileResult.Matrices) {
        $groupName = [string]$matrixObj.Setting.Formatted.GroupName
        $siteCode = [string]$matrixObj.Setting.Formatted.SiteCode

        if ($groupName) {
            if ($siteCode) {
                $fullPrefixes.Add(
                    [pscustomobject]@{
                        GroupName = $groupName
                        SiteCode  = $siteCode
                        Prefix    = ('{0} {1} ' -f $groupName, $siteCode)
                    }
                )
            }
            $groupPrefixes.Add($groupName)
        }
    }

    $fullPrefixes = @(
        $fullPrefixes | Sort-Object -Property Prefix -Unique |
        Sort-Object -Property { $_.Prefix.Length } -Descending
    )
    $groupPrefixes = @(
        $groupPrefixes | Sort-Object -Unique |
        Sort-Object -Property Length -Descending
    )

    $getNameParts = {
        param([string]$adName)

        foreach ($fp in $fullPrefixes) {
            if ($adName.StartsWith($fp.Prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                return [pscustomobject]@{
                    GroupName = $fp.GroupName
                    SiteCode  = $fp.SiteCode
                    Name      = $adName.Substring($fp.Prefix.Length)
                }
            }
        }

        foreach ($gp in $groupPrefixes) {
            if ($adName.StartsWith("$gp ", [System.StringComparison]::OrdinalIgnoreCase)) {
                return [pscustomobject]@{
                    GroupName = $gp
                    SiteCode  = $null
                    Name      = $adName.Substring($gp.Length + 1)
                }
            }
        }

        # Name doesn't follow the 'GroupName [SiteCode] Name' convention
        return [pscustomobject]@{
            GroupName = $null
            SiteCode  = $null
            Name      = $null
        }
    }
    #endregion

    #region Resolve all group managers in one batch
    $managerMap = @{}

    $managerDNs = @(
        foreach ($adName in $uniqueAdNames) {
            $detail = $detailMap[$adName]
            if (
                $detail.adObject.ObjectClass -eq 'group' -and
                $detail.adObject.ManagedBy
            ) {
                $detail.adObject.ManagedBy
            }
        }
    ) | Sort-Object -Unique

    if ($managerDNs) {
        $resolvedManagers = Get-ADObjectDetailHC `
            -ADObjectName $managerDNs `
            -Type 'DistinguishedName' `
            -MaxThreads $MaxThreads

        foreach ($rm in $resolvedManagers) {
            if ($rm.DistinguishedName) {
                $managerMap[$rm.DistinguishedName] = $rm
            }
        }
    }
    #endregion

    foreach ($adName in $uniqueAdNames) {
        $detail = $detailMap[$adName]
        $adObject = $detail.adObject

        #region AdObjects row
        $nameParts = & $getNameParts (
            $(if ($adObject.SamAccountName) { $adObject.SamAccountName }
                else { $adName })
        )

        $adObjectRows.Add(
            [pscustomobject]@{
                MatrixFileName = $FileResult.Item.BaseName
                SamAccountName = if ($adObject.SamAccountName) {
                    $adObject.SamAccountName
                }
                else { $adName }
                GroupName      = $nameParts.GroupName
                SiteCode       = $nameParts.SiteCode
                Name           = $nameParts.Name
                Enabled        = $adObject.Enabled
            }
        )
        #endregion

        if (-not $adObject) {
            # Not found in AD: visible in 'AdObjects', nothing to expand
            continue
        }

        if ($adObject.ObjectClass -eq 'group') {
            #region AccessList rows for group members
            if ($detail.adGroupMember) {
                foreach ($member in $detail.adGroupMember) {
                    $accessListRows.Add(
                        [pscustomobject]@{
                            SamAccountName       = $adObject.SamAccountName
                            Name                 = $adObject.Name
                            Type                 = 'group'
                            MemberName           = $member.Name
                            MemberSamAccountName = $member.SamAccountName
                            MemberEnabled        = $member.Enabled
                        }
                    )
                }
            }
            else {
                # Group without members: keep it visible in the sheet
                $accessListRows.Add(
                    [pscustomobject]@{
                        SamAccountName       = $adObject.SamAccountName
                        Name                 = $adObject.Name
                        Type                 = 'group'
                        MemberName           = $null
                        MemberSamAccountName = $null
                        MemberEnabled        = $null
                    }
                )
            }
            #endregion

            #region GroupManagers rows
            $manager = if (
                $adObject.ManagedBy -and
                $managerMap.ContainsKey($adObject.ManagedBy)
            ) {
                $managerMap[$adObject.ManagedBy]
            }
            else { $null }

            if ($manager -and $manager.adObject) {
                if (
                    $manager.adObject.ObjectClass -eq 'group' -and
                    $manager.adGroupMember
                ) {
                    # Manager is a group: one row per manager group member
                    foreach ($mgrMember in $manager.adGroupMember) {
                        $groupManagerRows.Add(
                            [pscustomobject]@{
                                GroupName         = $adObject.Name
                                ManagerName       = $manager.adObject.Name
                                ManagerType       = 'group'
                                ManagerMemberName = $mgrMember.Name
                                MemberEnabled     = $mgrMember.Enabled
                            }
                        )
                    }
                }
                else {
                    $groupManagerRows.Add(
                        [pscustomobject]@{
                            GroupName         = $adObject.Name
                            ManagerName       = $manager.adObject.Name
                            ManagerType       = $manager.adObject.ObjectClass
                            ManagerMemberName = $null
                            # The manager's own AD account status, so a
                            # disabled managing account is visible
                            MemberEnabled     = $manager.adObject.Enabled
                        }
                    )
                }
            }
            else {
                # No manager set: keep the group visible in the sheet
                $groupManagerRows.Add(
                    [pscustomobject]@{
                        GroupName         = $adObject.Name
                        ManagerName       = $null
                        ManagerType       = $null
                        ManagerMemberName = $null
                        MemberEnabled     = $null
                    }
                )
            }
            #endregion
        }
        else {
            #region AccessList row for a user used directly in the matrix
            $accessListRows.Add(
                [pscustomobject]@{
                    SamAccountName       = $adObject.SamAccountName
                    Name                 = $adObject.Name
                    Type                 = 'user'
                    MemberName           = $adObject.Name
                    MemberSamAccountName = $adObject.SamAccountName
                    MemberEnabled        = $adObject.Enabled
                }
            )
            #endregion
        }
    }

    return [pscustomobject]@{
        AccessList    = @(
            $accessListRows | Sort-Object SamAccountName, MemberName
        )
        GroupManagers = @(
            $groupManagerRows | Sort-Object GroupName, ManagerMemberName
        )
        AdObjects     = $adObjectRows.ToArray()
    }
}