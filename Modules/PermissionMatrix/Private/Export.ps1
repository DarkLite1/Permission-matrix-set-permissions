function Build-ExportDataHC {
    <#
    .SYNOPSIS
        Builds aggregated export data for permissions and ServiceNow form data.

    .DESCRIPTION
        Iterates through the processed matrices and extracts the execution
        results (Errors, Warnings, Paths, Actions) and form data into flat,
        structured lists.
        This output is specifically formatted to be fed directly into the HTML
        and Excel reporting functions.

    .PARAMETER ImportedMatrix
        An array of processed matrix file objects (typically from $Context.
        FileResults) containing the settings, execution checks, and associated
        form data.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$ImportedMatrix
    )

    $permissionsRows = [System.Collections.Generic.List[pscustomobject]]::new()
    $formDataRows = [System.Collections.Generic.List[pscustomobject]]::new()

    foreach ($I in $ImportedMatrix) {

        # Permissions export rows
        if ($I.Settings) {
            foreach ($S in $I.Settings) {
                $permissionsRows.Add([pscustomobject]@{
                        MatrixFile = $I.File.Item.Name
                        Computer   = $S.Import.ComputerName
                        Path       = $S.Import.Path
                        Action     = $S.Import.Action
                        Errors     = @(
                            $S.Check |
                            Where-Object { $_.Type -eq 'FatalError' }).Count
                        Warnings   = @(
                            $S.Check |
                            Where-Object { $_.Type -eq 'Warning' }).Count
                    })
            }
        }

        # FormData sheet export rows
        if ($I.FormData.Import) {
            $formDataRows.AddRange([pscustomobject[]]@($I.FormData.Import))
        }
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
        Builds export data from imported matrices, then writes the configured
        export artifacts to disk: Permissions Excel, ServiceNow FormData Excel,
        and the standalone overview HTML page.

        The overview HTML is generated internally — callers no longer pass it
        in. The email summary body is a separate artifact built by EndHC and
        is not used here.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$ImportedMatrix,
        [Parameter(Mandatory)]      $ExportSettings
    )

    $exportData = Build-ExportDataHC -ImportedMatrix $ImportedMatrix

    $results = [ordered]@{
        Permissions  = $null
        FormData     = $null
        OverviewHtml = $null
    }

    # 1. Permissions Excel
    if ($ExportSettings.PermissionsExcelFile) {
        $results.Permissions = Export-PermissionsFileHC `
            -Rows $exportData.Permissions `
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
                         ManagerMemberName
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
        [array]$AdObjectRows
    )

    try {
        if (-not (Test-Path -LiteralPath $SourceFilePath -PathType Leaf)) {
            throw "Source matrix file '$SourceFilePath' not found"
        }

        $destinationPath = Join-Path `
            -Path $LogFolder `
            -ChildPath (Split-Path -Path $SourceFilePath -Leaf)

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
                    'GroupName', 'ManagerName',
                    'ManagerType', 'ManagerMemberName'
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
                         manager group member is written.
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