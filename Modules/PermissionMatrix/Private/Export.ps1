function Build-ExportDataHC {
    <#
    .SYNOPSIS
        Builds aggregated export data for permissions and ServiceNow form data.

    .DESCRIPTION
        Iterates the flattened matrix objects (one per enabled Settings row,
        as found in $Context.AllMatrices) and extracts:

        - one Permissions row per matrix object, built from the formatted
          Settings values and the matrix object's own check list
        - one FormData row per matrix file, taken from the file's formatted
          FormData

        Because $Context.AllMatrices is flattened per Settings row, several
        matrix objects can share the same source file. The FormData row is
        therefore emitted only once per file (deduplicated on the shared
        FileContext), while permissions rows are emitted for every matrix
        object.

        This output is specifically formatted to be fed directly into the HTML
        and Excel reporting functions.

    .PARAMETER ImportedMatrix
        An array of matrix objects (typically $Context.AllMatrices). Each
        object exposes:
        - Setting.Formatted.{ComputerName, Path, Action}
        - Check (list of objects with a 'Type' of 'FatalError' / 'Warning')
        - FileContext.Item.{Name, FullName}
        - FileContext.Sheets.FormData.Formatted (a single formatted row, or
          $null when no ServiceNow / overview export is configured)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$ImportedMatrix
    )

    $permissionsRows = [System.Collections.Generic.List[pscustomobject]]::new()
    $formDataRows = [System.Collections.Generic.List[pscustomobject]]::new()

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
                    Warnings   = @(
                        $matrixObj.Check |
                        Where-Object { $_.Type -eq 'Warning' }).Count
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
                    $formDataRows.Add([pscustomobject]$formData)
                }
            }
        }
        #endregion
    }

    return [pscustomobject]@{
        Permissions = $permissionsRows.ToArray()
        FormData    = $formDataRows.ToArray()
    }
}

function Export-FilesHC {
    <#
    .SYNOPSIS
        Executes all export operations based on settings.

    .DESCRIPTION
        Writes the configured export artifacts to disk:

        - Permissions Excel: a single consolidated workbook holding the
          'AccessList', 'GroupManagers', 'AdObjects' and 'FormData'
          worksheets aggregated across every matrix file. The same per-file
          rows are still written into the per-matrix copies in the log folder
          by EndHC; to avoid resolving group managers in AD twice, the rows
          built here are cached on each file result (.LogSheets) so EndHC can
          reuse them.
        - ServiceNow FormData Excel
        - the standalone overview HTML page (generated internally)

        The email summary body is a separate artifact built by EndHC and is
        not used here.

    .PARAMETER FileResults
        The per-file result objects ($Context.FileResults). Required to build
        the consolidated Permissions workbook and the log-folder sheet rows.

    .PARAMETER AdObjectDetails
        The resolved AD objects of the run ($Context.AdObjectDetails), used to
        expand group members and managers for the AccessList / GroupManagers /
        AdObjects sheets.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$ImportedMatrix,
        [Parameter(Mandatory)]      $ExportSettings,
        [array]$FileResults = @(),
        [array]$AdObjectDetails = @()
    )

    $exportData = Build-ExportDataHC -ImportedMatrix $ImportedMatrix

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
            -Path $ExportSettings.PermissionsExcelFile
    }

    # 2. ServiceNow FormData Excel
    if ($ExportSettings.ServiceNowFormDataExcelFile) {
        $results.FormData = Export-ServiceNowFormDataHC `
            -Rows $exportData.FormData `
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

function Export-PermissionsFileHC {
    <#
        Writes a permissions Excel export using ImportExcel module.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$Rows,
        [Parameter(Mandatory)][string]$Path
    )

    try {
        $Rows | Export-Excel -Path $Path -WorksheetName 'Permissions' -AutoSize
        return $Path
    }
    catch {
        throw "Failed exporting Permissions Excel file: $_"
    }
}

function Export-ServiceNowFormDataHC {
    <#
        Writes ServiceNow FormData into an Excel file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$Rows,
        [Parameter(Mandatory)][string]$Path
    )

    try {
        $Rows | Export-Excel -Path $Path -WorksheetName 'FormData' -AutoSize
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
        Aggregates the log-sheet rows of every matrix file into one set of
        rows for the consolidated Permissions workbook.

    .DESCRIPTION
        For each file result, builds the per-file 'AccessList',
        'GroupManagers' and 'AdObjects' rows with Build-MatrixLogSheetRowsHC
        and combines them across all files:

        - AccessList   : each row is prefixed with a 'MatrixFileName' column
                         so rows from different files can be told apart in the
                         combined sheet
        - GroupManagers: each row is prefixed with a 'MatrixFileName' column
                         so rows from different files can be told apart in the
                         combined sheet; 'MemberEnabled' is preserved
        - AdObjects    : taken as-is (already carries 'MatrixFileName')
        - FormData     : one row per file, from the file's formatted FormData

        The unmodified per-file row sets are cached on each file result as a
        'LogSheets' property, so EndHC can reuse them when writing the
        per-matrix copy to the log folder instead of resolving group managers
        in AD a second time.

    .PARAMETER FileResults
        The per-file result objects ($Context.FileResults).

    .PARAMETER AdObjectDetails
        The resolved AD objects of the run ($Context.AdObjectDetails).

    .OUTPUTS
        PSCustomObject with the properties 'AccessList', 'GroupManagers',
        'AdObjects' and 'FormData'.
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

function Export-ConsolidatedPermissionsFileHC {
    <#
    .SYNOPSIS
        Writes the consolidated Permissions workbook with the 'AccessList',
        'GroupManagers', 'AdObjects' and 'FormData' worksheets.

    .DESCRIPTION
        Always creates all four worksheets, even when a row set is empty
        (header-only for the sheets with a fixed column layout), so the
        workbook always has the same structure. Any pre-existing file at the
        target path is replaced, so re-runs don't stack stale worksheets.

    .PARAMETER AccessList
        Aggregated AccessList rows.

    .PARAMETER GroupManagers
        Aggregated GroupManagers rows (including the 'MatrixFileName' column).

    .PARAMETER AdObjects
        Aggregated AdObjects rows.

    .PARAMETER FormData
        Aggregated FormData rows (one per file). Columns depend on the matrix
        template, so no fixed headers are written when this set is empty.

    .PARAMETER Path
        The target .xlsx path ($Context.Config.Export.PermissionsExcelFile).

    .OUTPUTS
        System.String - the path that was written.
    #>
    [CmdletBinding()]
    param(
        [array]$AccessList = @(),
        [array]$GroupManagers = @(),
        [array]$AdObjects = @(),
        [array]$FormData = @(),
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

        return $Path
    }
    catch {
        throw "Failed exporting consolidated Permissions Excel file '$Path': $_"
    }
}

function Copy-MatrixFileToLogFolderHC {
    <#
    .SYNOPSIS
        Copies the original matrix Excel file to the log folder and appends
        the worksheets 'AccessList', 'GroupManagers' and 'AdObjects'.

    .DESCRIPTION
        Creates a copy of the processed source matrix .xlsx file inside the
        dated log folder of that matrix file. The copy always contains the
        three extra worksheets, even when no rows are available (in that
        case only the header row is written), so users always find the
        same structure in every archived matrix file.

        The read-only attribute is stripped from the copy, because source
        matrix files are often opened/stored read-only and ImportExcel
        cannot write to a read-only file.

        Expected row shapes (extra properties are exported as extra columns):
        - AccessList   : SamAccountName, Name, Type, MemberName,
                         MemberSamAccountName, MemberEnabled
        - GroupManagers: GroupName, ManagerName, ManagerType,
                         ManagerMemberName, MemberEnabled
        - AdObjects    : MatrixFileName, SamAccountName, GroupName,
                         SiteCode, Name, Enabled

    .PARAMETER SourceFilePath
        The absolute path of the original matrix Excel file that was
        processed (e.g. $fileResult.Item.FullName).

    .PARAMETER LogFolder
        The absolute path of the log folder for this matrix file
        (e.g. $fileResult.LogFolder).

    .PARAMETER AccessListRows
        Rows for the 'AccessList' worksheet.

    .PARAMETER GroupManagerRows
        Rows for the 'GroupManagers' worksheet.

    .PARAMETER AdObjectRows
        Rows for the 'AdObjects' worksheet.

    .PARAMETER DestinationFileName
        Optional file name (with extension) for the copy inside LogFolder. When
        omitted, the source file's own name is used. Lets callers add a
        date-stamped name (e.g. the audit report's per-run history files).

    .OUTPUTS
        System.String
        The absolute path of the created copy.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFilePath,
        [Parameter(Mandatory)][string]$LogFolder,
        [array]$AccessListRows,
        [array]$GroupManagerRows,
        [array]$AdObjectRows,
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
        'AdObjects' worksheets in the matrix file copy saved to the log
        folder.

    .DESCRIPTION
        Transforms the resolved Active Directory details of one matrix file
        into three flat row collections, ready to be passed to
        Copy-MatrixFileToLogFolderHC:

        - AccessList   : one row per group member, including the new
                         'MemberEnabled' column (AD account status of the
                         member; blank for nested groups). Groups without
                         members still get one row with empty member
                         columns. AD objects of type 'user' that are used
                         directly in the matrix are listed with themselves
                         as member, so their account status is visible too.
        - GroupManagers: one row per group. When the group has a manager
                         ('managedBy'), the manager is resolved against AD.
                         If the manager is itself a group, one row per
                         manager group member is written with that member's
                         AD account status in 'MemberEnabled'. When the
                         manager is a single user, 'MemberEnabled' holds
                         the manager's own account status, so a disabled
                         managing account is visible.
        - AdObjects    : one row per unique AD object used in the matrix
                         file, including the new 'Enabled' column. The
                         'GroupName', 'SiteCode' and 'Name' columns are
                         derived by matching the AD object name against
                         the 'GroupName'/'SiteCode' values of the Settings
                         rows of this matrix file. Names that don't follow
                         the naming convention keep these columns blank.

        The AD object names used by this matrix file are collected from the
        per-folder 'AdNames' maps created during the SID rewrite in the
        BEGIN stage (falling back to the raw ACL keys when the rewrite was
        skipped). This also includes default permissions that were merged
        into the folder ACLs.

    .PARAMETER FileResult
        One matrix file result object from $Context.FileResults, containing
        the 'Item' (FileInfo) and 'Matrices' properties.

    .PARAMETER AdObjectDetails
        The resolved AD objects of the whole pipeline run:
        $Context.AdObjectDetails, populated by the BEGIN stage from
        Get-ADObjectDetailHC (objects with the properties 'SamAccountName',
        'adObject' and 'adGroupMember').

    .PARAMETER MaxThreads
        Maximum number of concurrent AD queries used to resolve group
        managers. (Default: 7)

    .OUTPUTS
        PSCustomObject with the properties 'AccessList', 'GroupManagers'
        and 'AdObjects'.
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