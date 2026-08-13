function Test-MatrixPermissionsHC {
    <#
    .SYNOPSIS
        Verify input for the Excel sheet 'Permissions'.

    .DESCRIPTION
        Verify if all input in the Excel sheet 'Permissions' is correct. When
        incorrect input is detected an object is returned containing all the
        details about the issue.
        This test is best run before expanding the matrix as it will save time.

    .PARAMETER Permissions
        The objects coming from the Excel sheet 'Permissions', as retrieved by
        Import-Excel.
    #>

    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param (
        [parameter(Mandatory)]
        [PSCustomObject[]]$Permissions
    )

    $checks = [System.Collections.Generic.List[pscustomobject]]::new()

    try {
        $Props = $Permissions[0].PSObject.Properties.Name
        $FirstProperty = $Props[0]

        #region Structural Validation (Fatal - Exits Immediately)
        if ($Permissions.Count -lt 4) {
            $checks.Add(
                (New-ValidationCheckHC `
                    -Type 'FatalError' `
                    -Name 'Missing rows' `
                    -Description 'At least 4 rows are required: 3 header rows and 1 row for the parent folder.' `
                    -Value "$($Permissions.Count) rows")
            )
            return $checks
        }

        if ($Props.Count -lt 2) {
            $checks.Add(
                (New-ValidationCheckHC `
                    -Type 'FatalError' `
                    -Name 'Missing columns' `
                    -Description 'At least 2 columns are required: 1 for the folder names and 1 where the permissions are defined.' `
                    -Value "$($Props.Count) column")
            )
            return $checks
        }
        #endregion

        #region Missing header SamAccountName
        $missingSamAccountNames = [System.Collections.Generic.List[string]]::new()

        foreach ($col in $Props) {
            if ($col -eq $FirstProperty) { continue }

            if ([string]::IsNullOrWhiteSpace($Permissions[0].$col) -and
                [string]::IsNullOrWhiteSpace($Permissions[1].$col) -and
                [string]::IsNullOrWhiteSpace($Permissions[2].$col)) {
                $missingSamAccountNames.Add($col)
            }
        }

        if ($missingSamAccountNames.Count -gt 0) {
            $checks.Add(
                (New-ValidationCheckHC `
                    -Type 'FatalError' `
                    -Name 'Missing AD object name' `
                    -Description 'The first 3 rows of the Permissions sheet are reserved for header information. Please provide the SamAccountName of the AD object in at least one of these rows for each column.' `
                    -Value "Columns: $($missingSamAccountNames -join ', ')")
            )
        }
        #endregion

        # Separate Headers from Data
        $NonHeaderRows = $Permissions | Select-Object -Skip 3

        <# The sub-folder rows: everything after the 3 header rows and the
        parent folder row on row 4. This slice was previously built twice,
        here and again as '$Permissions | Select-Object -Skip 4' further
        down, which is the same set of rows by a different route. #>
        $FolderRows = $NonHeaderRows | Select-Object -Skip 1

        #region Permission character unknown
        $InvalidChars = [System.Collections.Generic.List[string]]::new()

        foreach ($Row in $NonHeaderRows) {
            $PermColumns = $Row.PSObject.Properties.Where({ $_.Name -ne $FirstProperty })
            foreach ($Col in $PermColumns) {
                $Ace = $Col.Value
                if (
                    -not [string]::IsNullOrWhiteSpace($Ace) -and
                    $Ace -notmatch '^(L|R|W|I|F)$'
                ) {
                    $InvalidChars.Add($Ace)
                }
            }
        }

        if ($InvalidChars.Count -gt 0) {
            $checks.Add(
                (New-ValidationCheckHC `
                    -Type 'FatalError' `
                    -Name 'Invalid permission character' `
                    -Description "Supported characters are 'F', 'W', 'R', 'L', 'I' or blank." `
                    -Value "Characters: $(($InvalidChars | Select-Object -Unique) -join ', ')")
            )
        }
        #endregion

        #region Folder name missing
        $MissingFolders = $FolderRows.Where(
            { [string]::IsNullOrWhiteSpace($_.$FirstProperty) }
        )

        if ($MissingFolders.Count -gt 0) {
            $checks.Add(
                (New-ValidationCheckHC `
                    -Type 'FatalError' `
                    -Name 'Missing folder name' `
                    -Description 'Each row needs a folder name in the first column.' `
                    -Value "$($MissingFolders.Count) missing folder name(s) in column 1")
            )
        }
        #endregion

        #region Duplicate folder name
        $NotUniqueFolder = $FolderRows.$FirstProperty | Group-Object | Where-Object Count -GE 2
        if ($NotUniqueFolder) {
            $checks.Add(
                (New-ValidationCheckHC `
                    -Type 'FatalError' `
                    -Name 'Duplicate folder name' `
                    -Description 'Folder names in the first column need to be unique.' `
                    -Value (($NotUniqueFolder.Name) -join ', '))
            )
        }
        #endregion

        #region Deepest folder has only List permissions or none at all
        <#
         Normalize paths before comparing: trim surrounding whitespace and
         trailing backslashes, so a cell typed as 'BEL\L&D\Certificates\' is
         still recognized as the parent of 'BEL\L&D\Certificates\AGG'.
         Child detection uses String.StartsWith instead of '-like' so
         wildcard characters in folder names ('[', ']', '*', '?') cannot
         break the match. NTFS paths are case-insensitive, so all
         comparisons use OrdinalIgnoreCase.
        #>
        $normalizePath = { 
            param($p) 
            ([string]$p).Trim().TrimEnd('\', '/') -replace '/', '\' 
        }

        $Paths = @(
            $FolderRows.$FirstProperty.Where({
                    -not [string]::IsNullOrWhiteSpace($_)
                }).ForEach({ & $normalizePath $_ })
        )

        # Faster check for deepest folders
        $DeepestFolders = [System.Collections.Generic.HashSet[string]]::new(
            [System.StringComparer]::OrdinalIgnoreCase
        )

        foreach ($P in $Paths) {
            $childPrefix = '{0}\' -f $P
            $hasChild = $false

            foreach ($other in $Paths) {
                if (
                    ($other.Length -gt $childPrefix.Length) -and
                    $other.StartsWith(
                        $childPrefix,
                        [System.StringComparison]::OrdinalIgnoreCase
                    )
                ) {
                    $hasChild = $true
                    break
                }
            }

            if (-not $hasChild) { [void]$DeepestFolders.Add($P) }
        }

        # Parent folder permissions (Row index 3)
        $ParentFolderPermissions = $Permissions[3].PSObject.Properties.Where({
                $_.Name -ne $FirstProperty -and -not [string]::IsNullOrWhiteSpace($_.Value)
            }).Value

        $ParentFolderHasPermission = [bool]($ParentFolderPermissions.Where({ $_ -ne 'L' }))
        $inAccessibleFolders = [System.Collections.Generic.List[string]]::new()

        <#
         Folders marked with 'I' (Ignore) are not managed by the matrix.
         Exclude them, and every subfolder beneath them, from the
         inaccessible check. Ignored rows do still count as children when
         determining the deepest folders, so a parent whose only children
         are ignored is not treated as a leaf.
        #>
        $ignoredRoots = [System.Collections.Generic.List[string]]::new()

        foreach ($Row in $FolderRows) {
            $isIgnored = [bool]$Row.PSObject.Properties.Where({
                    $_.Name -ne $FirstProperty -and $_.Value -eq 'I'
                }, 'First').Count

            if ($isIgnored) {
                $p = & $normalizePath $Row.$FirstProperty
                if (-not [string]::IsNullOrWhiteSpace($p)) {
                    $ignoredRoots.Add($p)
                }
            }
        }

        $isInIgnoredSubtree = {
            param($p)
            foreach ($root in $ignoredRoots) {
                if (
                    $p.Equals(
                        $root, [System.StringComparison]::OrdinalIgnoreCase
                    ) -or
                    $p.StartsWith(
                        ('{0}\' -f $root),
                        [System.StringComparison]::OrdinalIgnoreCase
                    )
                ) {
                    return $true
                }
            }
            return $false
        }

        foreach ($Row in $FolderRows) {
            $rowPath = & $normalizePath $Row.$FirstProperty

            if (
                [string]::IsNullOrWhiteSpace($rowPath) -or
                (-not $DeepestFolders.Contains($rowPath)) -or
                (& $isInIgnoredSubtree $rowPath)
            ) {
                continue
            }

            $definedPermissions = $Row.PSObject.Properties.Where({
                    $_.Name -ne $FirstProperty -and
                    -not [string]::IsNullOrWhiteSpace($_.Value)
                }).Value

            $accessPermissions = $definedPermissions.Where({ $_ -ne 'L' })

            if (-not $accessPermissions) {
                <#
                 Two distinct situations for a deepest folder without any
                 access-granting (non 'L') permission on its row:
                 - The row defines only 'L': the folder gets an explicit
                   List-only ACL. Users can never read or write there, no
                   matter what the parent grants => always inaccessible.
                 - The row is completely blank: the folder simply inherits
                   the parent ACL => only inaccessible when the parent
                   does not grant access either.
                #>
                if ($definedPermissions -or (-not $ParentFolderHasPermission)) {
                    # Report the cell content as typed in Excel, so the user
                    # can easily find the row back in the worksheet.
                    $inAccessibleFolders.Add($Row.$FirstProperty)
                }
            }
        }

        if ($inAccessibleFolders.Count -gt 0) {
            $checks.Add(
                (New-ValidationCheckHC `
                    -Type 'Warning' `
                    -Name 'Inaccessible folders' `
                    -Description 'The deepest folders are not accessible: they either define only List permissions, or they define no permissions at all while the parent folder does not grant access. Users can list but never read or write content in these folders. Folders marked with ''I'' (Ignore) and their subfolders are excluded from this check.' `
                    -Value ($inAccessibleFolders -join ', '))
            )
        }
        #endregion

        # Output all collected errors at the end
        if ($checks.Count -gt 0) {
            return $checks
        }

    }
    catch {
        throw "Failed testing the Excel sheet 'Permissions' for incorrect data: $_"
    }
}

function Test-MatrixFormDataHC {
    <#
    .SYNOPSIS
        Verify input for the Excel sheet 'FormData'.

    .DESCRIPTION
        Verify if the Excel sheet 'FormData' contains the correct data.

    .PARAMETER FormData
        Represents the data coming from the Excel sheet 'FormData'. When no rows
        are supplied (null or empty) a non-fatal Warning is returned, consistent
        with the other Test-Matrix*HC validators. The parameter is intentionally
        not Mandatory so a missing sheet can be reported rather than rejected at
        parameter binding.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param (
        [PSCustomObject[]]$FormData
    )

    try {
        #region No FormData -> Warning
        if ((-not $FormData) -or ($FormData.Count -eq 0)) {
            return New-ValidationCheckHC `
                -Type 'Warning' `
                -Name 'Missing FormData' `
                -Description 'No FormData rows were found. ServiceNow form data will not be exported for this matrix file.' `
                -Category 'FormData'
        }
        #endregion

        if ($FormData.Count -ne 1) {
            return New-ValidationCheckHC `
                -Type 'FatalError' `
                -Name 'Incorrect row count' `
                -Description "Exactly one row of data is required. Found $($FormData.Count) row(s)." `
                -Value $FormData.Count `
                -Category 'FormData'
        }

        $Row = $FormData[0]
        $Properties = ($Row | Get-Member -MemberType NoteProperty).Name

        $MandatoryProperties = @(
            'MatrixFormStatus',
            'MatrixCategoryName',
            'MatrixSubCategoryName',
            'MatrixResponsible',
            'MatrixFolderDisplayName',
            'MatrixFolderPath'
        )

        #region Missing column headers
        $MissingProperties = $MandatoryProperties.Where({ $_ -notin $Properties })

        if ($MissingProperties) {
            return New-ValidationCheckHC `
                -Type 'FatalError' `
                -Name 'Missing column header' `
                -Description "The following column headers are mandatory: $($MandatoryProperties -join ', ')." `
                -Value ($MissingProperties -join ', ') `
                -Category 'FormData'
        }
        #endregion

        #region Mandatory property values (Only if Enabled)
        <# Compare a normalized copy. Format-FormDataStringsHC trims every
        string field, but it only runs after this validator has passed, so
        a MatrixFormStatus of 'Enabled ' with a trailing space skipped the
        mandatory-value check entirely. The row was still written to the
        ServiceNow export, which does not filter on status, so blank
        mandatory values reached ServiceNow unreported. #>
        if ("$($Row.MatrixFormStatus)".Trim() -eq 'Enabled') {

            $MandatoryPropertyValues = $MandatoryProperties.Where({ $_ -ne 'MatrixFormStatus' })

            $BlankProperties = $MandatoryPropertyValues.Where({
                    [string]::IsNullOrWhiteSpace($Row.$_)
                })

            if ($BlankProperties) {
                return New-ValidationCheckHC `
                    -Type 'FatalError' `
                    -Name 'Missing value' `
                    -Description "Values for the following columns are mandatory when status is Enabled: $($MandatoryPropertyValues -join ', ')." `
                    -Value ($BlankProperties -join ', ') `
                    -Category 'FormData'
            }
        }
        #endregion
    }
    catch {
        throw "Failed testing the Excel sheet 'FormData': $_"
    }
}

function Test-MatrixSettingRowHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$SettingRow,
        [Parameter()][bool]$RequireGroupName = $false,
        [Parameter()][bool]$RequireSiteCode = $false
    )

    $checks = [System.Collections.Generic.List[pscustomobject]]::new()

    $validActions = @('Fix', 'New', 'Check')

    if ([string]::IsNullOrWhiteSpace($SettingRow.Action)) {
        $checks.Add(
            (New-ValidationCheckHC `
                -Type 'FatalError' `
                -Name 'Missing Action' `
                -Description "The column 'Action' cannot be empty." `
                -Value $null)
        )
    }
    elseif ($SettingRow.Action -notin $validActions) {
        $checks.Add(
            (New-ValidationCheckHC `
                -Type 'FatalError' `
                -Name 'Invalid Action' `
                -Description "Supported Action values are '$($validActions -join "', '")'." `
                -Value "Found: '$($SettingRow.Action)'")
        )
    }

    if ([string]::IsNullOrWhiteSpace($SettingRow.Path)) {
        $checks.Add(
            (New-ValidationCheckHC `
                -Type 'FatalError' `
                -Name 'Missing Path' `
                -Description "The column 'Path' cannot be empty." `
                -Value $null)
        )
    }

    # Validate ApplyDefaultPermissions
    if (
        $SettingRow.PSObject.Properties.Match('ApplyDefaultPermissions').Count -gt 0 -and
        -not [string]::IsNullOrWhiteSpace($SettingRow.ApplyDefaultPermissions)
    ) {
        $parsed = $false
        # If the value cannot be parsed strictly as a boolean, flag it as a FatalError
        if (-not [bool]::TryParse($SettingRow.ApplyDefaultPermissions.ToString(), [ref]$parsed)) {
            $checks.Add(
                (New-ValidationCheckHC `
                    -Type 'FatalError' `
                    -Name 'Invalid ApplyDefaultPermissions' `
                    -Description "The column 'ApplyDefaultPermissions' must be a valid boolean ('True' or 'False') or left blank." `
                    -Value "Found: '$($SettingRow.ApplyDefaultPermissions)'")
            )
        }
    }

    if ([string]::IsNullOrWhiteSpace($SettingRow.ComputerName)) {
        $checks.Add(
            (New-ValidationCheckHC `
                -Type 'FatalError' `
                -Name 'Missing ComputerName' `
                -Description "The column 'ComputerName' cannot be empty." `
                -Value $null)
        )
    }

    if (
        $RequireSiteCode -and
        [string]::IsNullOrWhiteSpace($SettingRow.SiteCode)
    ) {
        $checks.Add(
            (New-ValidationCheckHC `
                -Type 'FatalError' `
                -Name 'Missing SiteCode' `
                -Description "The column 'SiteCode' cannot be empty because it is used as a placeholder in the Permissions sheet." `
                -Value $null)
        )
    }

    if (
        $RequireGroupName -and
        [string]::IsNullOrWhiteSpace($SettingRow.GroupName)
    ) {
        $checks.Add(
            (New-ValidationCheckHC `
                -Type 'FatalError' `
                -Name 'Missing GroupName' `
                -Description "The column 'GroupName' cannot be empty because it is used as a placeholder in the Permissions sheet." `
                -Value $null)
        )
    }

    $applyDefaults = $SettingRow.ApplyDefaultPermissions
    if ([string]::IsNullOrWhiteSpace($applyDefaults)) {
        $checks.Add(
            (New-ValidationCheckHC `
                -Type 'FatalError' `
                -Name 'Missing ApplyDefaultPermissions' `
                -Description "The column 'ApplyDefaultPermissions' cannot be empty." `
                -Value $null)
        )
    }
    else {
        # Safely test if the value can be evaluated as a true/false boolean
        $parsedBool = $false
        if (-not [bool]::TryParse($applyDefaults.ToString(), [ref]$parsedBool)) {
            $checks.Add(
                (New-ValidationCheckHC `
                    -Type 'FatalError' `
                    -Name 'Invalid ApplyDefaultPermissions' `
                    -Description "The column 'ApplyDefaultPermissions' must be a valid boolean (True or False)." `
                    -Value "Found: '$applyDefaults'")
            )
        }
    }

    return $checks
}

function Test-AdObjectInMatrixHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$Matrix,
        [Parameter(Mandatory)]$ADObject,
        [string[]]$ExcludedSamAccountName = @()
    )

    $checks = @()

    $matrixAdObjects = @($Matrix.ACL.Keys) | Select-Object -Unique

    if (-not $matrixAdObjects) { return $checks }

    #region Index the resolved AD details once for O(1) lookups
    $detailLookup = @{}

    foreach ($detail in $ADObject) {
        if ($detail.SamAccountName) {
            $detailLookup[[string]$detail.SamAccountName] = $detail
        }
    }
    #endregion

    #region Placeholder accounts, case-insensitive
    $placeHolders = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]@($ExcludedSamAccountName | Where-Object { $_ }),
        [System.StringComparer]::OrdinalIgnoreCase
    )
    #endregion

    #region Unknown AD objects (fatal)
    $missingAdObjects = $matrixAdObjects | Where-Object {
        $null -eq $detailLookup[[string]$_].adObject
    }

    if ($missingAdObjects) {
        $checks += New-ValidationCheckHC `
            -Type 'FatalError' `
            -Name 'Unknown AD Objects in Matrix' `
            -Description 'One or more AD objects referenced in the matrix were not found in Active Directory. Please check the SamAccountName values in the Permissions sheet and ensure they exist in AD.' `
            -Value @($missingAdObjects | Sort-Object)
    }
    #endregion

    #region Groups without effective members (warning)
    <#
     A group grants access to nobody when, after discarding the placeholder
     accounts from 'Matrix.ExcludedSamAccountName' and the disabled accounts,
     no members are left. This covers three cases in one check:
     - the group is completely empty
     - the group only holds the placeholder account(s)
     - the group only holds disabled accounts
     ...and the combination of the last two.

     Members with an Enabled value of $null (nested groups, and the synthetic
     'All users' member of 'Domain Users') are counted as real members.

     The reason is reported per group because the remedy differs: an empty
     group needs members, a placeholder-only group needs the placeholder
     replaced, and a disabled-only group needs the accounts re-enabled or
     swapped out.
    #>
    $emptyGroups = [System.Collections.Generic.List[psobject]]::new()

    <#
     User accounts referenced directly in the matrix that are disabled in AD.

     Collected in the same pass as the empty-group check to avoid a second walk
     over the matrix objects.
    #>
    $disabledUsers = [System.Collections.Generic.List[psobject]]::new()

    foreach ($name in $matrixAdObjects) {
        $detail = $detailLookup[[string]$name]

        # Already reported above as an unknown AD object
        if (-not $detail.adObject) { continue }

        if ($detail.adObject.ObjectClass -ne 'group') {
            <#
             A disabled user granted permissions directly is dead access: the
             ACE stays on the folder but nobody can use it. Unlike a group, it
             will not start working again when someone is added, so it needs
             either removing from the matrix or the account re-enabling.

             'Enabled' is compared to $false explicitly, not tested for
             truthiness. Get-ADObjectDetailHC leaves it $null when
             'useraccountcontrol' was not returned, and an unknown state must
             not be reported as disabled.

             Placeholder accounts are skipped: they exist precisely to occupy a
             matrix slot without granting access, so reporting them every run
             would be noise. This matches how the empty-group check below treats
             them.
            #>
            if (
                ($detail.adObject.Enabled -eq $false) -and
                (-not $placeHolders.Contains([string]$name))
            ) {
                $disabledUsers.Add(
                    [PSCustomObject]@{
                        Name              = [string]$name
                        DisplayName       = [string]$detail.adObject.Name
                        DistinguishedName = [string]$detail.adObject.DistinguishedName
                    }
                )
            }

            continue
        }

        # 'adGroupMember' is $null when the group could not be expanded
        $members = @($detail.adGroupMember | Where-Object { $_ })

        $effectiveMembers = @(
            $members | Where-Object {
                (-not $placeHolders.Contains([string]$_.SamAccountName)) -and
                ($_.Enabled -ne $false)
            }
        )

        if ($effectiveMembers.Count -ne 0) { continue }

        #region Explain why the group grants access to nobody
        $placeHolderMemberCount = @(
            $members | Where-Object {
                $placeHolders.Contains([string]$_.SamAccountName)
            }
        ).Count

        $disabledMemberCount = @(
            $members | Where-Object {
                ($_.Enabled -eq $false) -and
                (-not $placeHolders.Contains([string]$_.SamAccountName))
            }
        ).Count

        $reason = if ($members.Count -eq 0) {
            'no members'
        }
        elseif ($disabledMemberCount -eq 0) {
            'only placeholder accounts'
        }
        elseif ($placeHolderMemberCount -eq 0) {
            'only disabled accounts'
        }
        else {
            'only placeholder and disabled accounts'
        }
        #endregion

        $emptyGroups.Add(
            [PSCustomObject]@{
                Name   = [string]$name
                Reason = $reason
            }
        )
    }

    if ($emptyGroups.Count -gt 0) {
        <#
         Emit one entry per group, sorted by name, as structured objects
         rather than preformatted text.

         These were previously rendered into strings ("'<name>' : <reason>",
         the name padded to the widest name so the reasons lined up). That
         put presentation into the data: the detail JSON stored padding as
         real characters, and anything that wanted the name or the reason
         back had to parse it out with a regex. Emitting the objects lets
         ConvertTo-Json write proper Name/Reason fields, so the JSON is both
         readable and queryable (Where-Object Reason -eq 'no members'), and
         any future alignment becomes a rendering decision at the point of
         display.

         $emptyGroups already holds objects of exactly this shape, so this
         is the list sorted, not a rebuild.
        #>
        $emptyGroupList = @($emptyGroups | Sort-Object -Property 'Name')

        #region Resolve the configured placeholder accounts by name
        $placeHolderNames = @($placeHolders) | Sort-Object

        $placeHolderText = if ($placeHolderNames.Count -gt 0) {
            "Placeholder accounts configured in 'Matrix.ExcludedSamAccountName': $($placeHolderNames -join ', ')."
        }
        else {
            "No placeholder accounts are configured in 'Matrix.ExcludedSamAccountName'."
        }
        #endregion

        $checks += New-ValidationCheckHC `
            -Type 'Information' `
            -Name 'AD groups without members' `
            -Description "One or more AD groups in the matrix have no effective members: they are empty, they only contain placeholder accounts, or they only contain disabled accounts. No one has access to the folders granted to these groups. $placeHolderText" `
            -Value $emptyGroupList
    }
    #endregion

    #region Disabled user accounts used directly in the matrix (information)
    if ($disabledUsers.Count -gt 0) {
        $checks += New-ValidationCheckHC `
            -Type 'Information' `
            -Name 'Disabled AD user accounts' `
            -Description 'One or more user accounts granted permissions directly in the matrix are disabled in Active Directory. The permissions are still applied to the folders, but the accounts cannot use them. Remove the accounts from the matrix or re-enable them in Active Directory.' `
            -Value @($disabledUsers | Sort-Object -Property 'Name')
    }
    #endregion

    return $checks
}

function Test-ConfigurationStructureHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Json,
        [Parameter(Mandatory)][ref]   $SystemErrors
    )

    #region Top-Level properties
    foreach ($prop in @(
            'Matrix', 'Export', 'ServiceNow', 'MaxConcurrent', 'PSSessionConfiguration', 'Settings'
        )) {
        if ($null -eq $Json.$prop) {
            Add-JsonSchemaErrorHC `
                -Type 'FatalError' `
                -Name "Missing '$prop'" `
                -Message "Property '$prop' not found in JSON." `
                -SystemErrors $SystemErrors
        }
    }
    #endregion

    #region Settings
    if ($Json.Settings) {
        #region SaveInEventLog
        if ($Json.Settings.SaveInEventLog.Save -isnot [bool]) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Incorrect 'Settings.SaveInEventLog.Save'" `
                -Message 'Must be boolean.' `
                -SystemErrors $SystemErrors
        }
        #endregion

        #region SaveLogFiles
        if (-not $Json.Settings.SaveLogFiles.Where.Folder) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Missing 'Settings.SaveLogFiles.Where.Folder'" `
                -Message 'Folder is required.' `
                -SystemErrors $SystemErrors
        }

        if ($null -eq $Json.Settings.SaveLogFiles.Detailed) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Missing 'Settings.SaveLogFiles.Detailed'" `
                -Message 'Detailed is required.' `
                -SystemErrors $SystemErrors
        }
        elseif ($Json.Settings.SaveLogFiles.Detailed -isnot [bool]) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Incorrect 'Settings.SaveLogFiles.Detailed'" `
                -Message 'Must be boolean.' `
                -SystemErrors $SystemErrors
        }
        #endregion

        #region SendMail
        if ( $Json.Settings.SendMail) {
            if (-not $Json.Settings.SendMail.From) {
                Add-JsonSchemaErrorHC -Type 'FatalError' `
                    -Name "Missing 'Settings.SendMail.From'" `
                    -Message 'From is required.' `
                    -SystemErrors $SystemErrors
            }
            if ($Json.Settings.SendMail.To -and
                ($Json.Settings.SendMail.To -isnot [string] -and
                $Json.Settings.SendMail.To -isnot [array])) {
                Add-JsonSchemaErrorHC -Type 'FatalError' `
                    -Name "Incorrect 'Settings.SendMail.To'" `
                    -Message 'Must be string or array.' `
                    -SystemErrors $SystemErrors
            }
            if ($null -eq $Json.Settings.SendMail.Body) {
                Add-JsonSchemaErrorHC -Type 'FatalError' `
                    -Name "Missing 'Settings.SendMail.Body'" `
                    -Message 'Body is required.' `
                    -SystemErrors $SystemErrors
            }
            if (-not $Json.Settings.SendMail.Smtp.Port -or $Json.Settings.SendMail.Smtp.Port -notmatch '^\d+$') {
                Add-JsonSchemaErrorHC -Type 'FatalError' -Name "Incorrect 'SendMail.Smtp.Port'" `
                    -Message 'Port must be numeric.' `
                    -SystemErrors $SystemErrors
            }

            $validConn = @('None', 'Auto', 'SslOnConnect', 'StartTls', 'StartTlsWhenAvailable')
            if ($Json.Settings.SendMail.Smtp.ConnectionType -notin $validConn) {
                Add-JsonSchemaErrorHC -Type 'FatalError' -Name "Incorrect 'Settings.SendMail.Smtp.ConnectionType'" `
                    -Message 'Invalid connection type.' `
                    -SystemErrors $SystemErrors
            }
        }
        else {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Missing 'Settings.SendMail'" `
                -Message 'SendMail block is mandatory.' `
                -SystemErrors $SystemErrors

        }
        #endregion

        if (-not $json.Settings.ScriptName) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Missing 'Settings.ScriptName'" `
                -Message 'ScriptName is required.' `
                -SystemErrors $SystemErrors
        }
    }
    #endregion

    #region Matrix
    if ($Json.Matrix) {
        if (-not $Json.Matrix.FolderPath) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Missing 'Matrix.FolderPath'" `
                -Message "Property 'Matrix.FolderPath' not found" `
                -SystemErrors $SystemErrors
        }
        elseif (-not (Test-Path -LiteralPath $Json.Matrix.FolderPath -PathType Container)) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Incorrect 'Matrix.FolderPath'" `
                -Message "Property 'Matrix.FolderPath' path '$($Json.Matrix.FolderPath)' not found" `
                -SystemErrors $SystemErrors
        }

        if (-not $Json.Matrix.DefaultsFile) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Missing 'Matrix.DefaultsFile'" `
                -Message "Property 'Matrix.DefaultsFile' not found" `
                -SystemErrors $SystemErrors
        }
        elseif (-not (Test-Path -LiteralPath $Json.Matrix.DefaultsFile -PathType Leaf)) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Incorrect 'Matrix.DefaultsFile'" `
                -Message "Property 'Matrix.DefaultsFile' path '$($Json.Matrix.DefaultsFile)' not found" `
                -SystemErrors $SystemErrors
        }

        if ($Json.Matrix.AdGroupPlaceHolders -and
            $Json.Matrix.AdGroupPlaceHolders -isnot [array]) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Incorrect 'Matrix.AdGroupPlaceHolders'" `
                -Message "Property 'Matrix.AdGroupPlaceHolders' must be an array." `
                -SystemErrors $SystemErrors
        }

        if ($null -eq $Json.Matrix.Archive -or $Json.Matrix.Archive -isnot [bool]) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Incorrect 'Matrix.Archive'" `
                -Message "Property 'Matrix.Archive' must be boolean." `
                -SystemErrors $SystemErrors
        }
    }
    #endregion

    #region MaxConcurrent
    if ($Json.MaxConcurrent) {
        foreach ($prop in 'JobsTotal', 'JobsPerComputer', 'FoldersPerMatrix') {
            $val = $Json.MaxConcurrent.$prop
            if ($null -eq $val -or $val -notmatch '^\d+$') {
                Add-JsonSchemaErrorHC -Type 'FatalError' `
                    -Name "Incorrect 'MaxConcurrent.$prop'" `
                    -Message "Property 'MaxConcurrent.$prop' must be numeric." `
                    -SystemErrors $SystemErrors
            }
            elseif ([int]$val -lt 1) {
                <# Zero is numeric but not a usable throttle: these values end
                up as -ThrottleLimit on ForEach-Object -Parallel, which needs
                at least 1. Rejecting it here gives a clear message instead of
                a parameter binding failure part way through a run. #>
                Add-JsonSchemaErrorHC -Type 'FatalError' `
                    -Name "Incorrect 'MaxConcurrent.$prop'" `
                    -Message "Property 'MaxConcurrent.$prop' must be 1 or higher." `
                    -SystemErrors $SystemErrors
            }
        }

        #region JobsPerComputer cannot exceed JobsTotal
        # A per-computer cap larger than the total is unreachable: the throttle
        # would stop the run before a single computer ever reached it, so the
        # configured value would silently mean something else.
        if (
            ($Json.MaxConcurrent.JobsTotal -match '^\d+$') -and
            ($Json.MaxConcurrent.JobsPerComputer -match '^\d+$') -and
            ([int]$Json.MaxConcurrent.JobsPerComputer -gt [int]$Json.MaxConcurrent.JobsTotal)
        ) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Incorrect 'MaxConcurrent.JobsPerComputer'" `
                -Message "Property 'MaxConcurrent.JobsPerComputer' ($($Json.MaxConcurrent.JobsPerComputer)) cannot be greater than 'MaxConcurrent.JobsTotal' ($($Json.MaxConcurrent.JobsTotal))." `
                -SystemErrors $SystemErrors
        }
        #endregion
    }
    #endregion

    #region Export
    if ($Json.Export) {
        if ($Json.Export.PermissionsExcelFile -and $Json.Export.PermissionsExcelFile -notmatch '\.xlsx$') {
            Add-JsonSchemaErrorHC -Type 'FatalError' -Name "Incorrect 'Export.PermissionsExcelFile'" `
                -Message 'Must end with .xlsx' `
                -SystemErrors $SystemErrors
        }

        if ($Json.Export.OverviewHtmlFile -and $Json.Export.OverviewHtmlFile -notmatch '\.html?$') {
            Add-JsonSchemaErrorHC -Type 'FatalError' -Name "Incorrect 'Export.OverviewHtmlFile'" `
                -Message 'Must end with .html' `
                -SystemErrors $SystemErrors
        }

        if ($Json.Export.ServiceNowFormDataExcelFile) {

            if ($Json.Export.ServiceNowFormDataExcelFile -notmatch '\.xlsx$') {
                Add-JsonSchemaErrorHC -Type 'FatalError' `
                    -Name "Incorrect 'Export.ServiceNowFormDataExcelFile'" `
                    -Message 'Must end with .xlsx' `
                    -SystemErrors $SystemErrors
            }

            if (-not $Json.ServiceNow) {
                Add-JsonSchemaErrorHC -Type 'FatalError' `
                    -Name 'Incorrect configuration' `
                    -Message 'ServiceNow must be defined when using ServiceNowFormDataExcelFile.' `
                    -SystemErrors $SystemErrors
            }
            else {
                foreach ($p in 'CredentialsFilePath', 'TableName', 'Environment') {
                    if (-not $Json.ServiceNow.$p) {
                        Add-JsonSchemaErrorHC -Type 'FatalError' `
                            -Name "Missing 'ServiceNow.$p'" `
                            -Message "$p is required." `
                            -SystemErrors $SystemErrors
                    }
                }
            }
        }
    }
    #endregion

    #region SharePoint
    if ($Json.SharePoint -and $Json.SharePoint.SiteUrl) {
        if ($Json.SharePoint.SiteUrl -notmatch '^https://') {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Incorrect 'SharePoint.SiteUrl'" `
                -Message 'Must be a URL starting with https://' `
                -SystemErrors $SystemErrors
        }

        foreach ($p in 'DocumentLibraryName', 'ClientId', 'TenantId', 'CertificateThumbprint') {
            if (-not $Json.SharePoint.$p) {
                Add-JsonSchemaErrorHC -Type 'FatalError' `
                    -Name "Missing 'SharePoint.$p'" `
                    -Message "$p is required when 'SharePoint.SiteUrl' is used." `
                    -SystemErrors $SystemErrors
            }
        }

        if (-not $Json.Export.OverviewHtmlFile) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name 'Incorrect configuration' `
                -Message "'Export.OverviewHtmlFile' must be defined when 'SharePoint.SiteUrl' is used." `
                -SystemErrors $SystemErrors
        }
    }
    #endregion
}