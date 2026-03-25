function ConvertTo-AceHC {
    <#
    .SYNOPSIS
        Convert an AD Object name and a permission character to a valid ACE.

    .DESCRIPTION
        Convert an AD Object name and a permission character to a valid Access Control List Entry.

    .PARAMETER Type
        The permission character defining the access to the folder. Valid values: L, R, W, F, M.

    .PARAMETER Name
        Name of the AD object, used to identify the user or group within AD.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateSet('L', 'R', 'W', 'F', 'M')]
        [String]$Type,

        [Parameter(Mandatory)]
        [String]$Name
    )

    $Identity = if ($Name -match '\\') { $Name } else { "$env:USERDOMAIN\$Name" }

    switch ($Type) {
        'L' {
            return [System.Security.AccessControl.FileSystemAccessRule]::new(
                $Identity,
                [System.Security.AccessControl.FileSystemRights]::ReadAndExecute,
                [System.Security.AccessControl.InheritanceFlags]::ContainerInherit,
                [System.Security.AccessControl.PropagationFlags]::None,
                [System.Security.AccessControl.AccessControlType]::Allow
            )
        }
        'W' {
            # This folder only
            [System.Security.AccessControl.FileSystemAccessRule]::new(
                $Identity,
                [System.Security.AccessControl.FileSystemRights]'CreateFiles, AppendData, DeleteSubdirectoriesAndFiles, ReadAndExecute, Synchronize',
                [System.Security.AccessControl.InheritanceFlags]::None,
                [System.Security.AccessControl.PropagationFlags]::InheritOnly,
                [System.Security.AccessControl.AccessControlType]::Allow
            )
            # Subfolders and files only
            [System.Security.AccessControl.FileSystemAccessRule]::new(
                $Identity,
                [System.Security.AccessControl.FileSystemRights]'DeleteSubdirectoriesAndFiles, Modify, Synchronize',
                [System.Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit',
                [System.Security.AccessControl.PropagationFlags]::InheritOnly,
                [System.Security.AccessControl.AccessControlType]::Allow
            )
            return
        }
        'R' {
            return [System.Security.AccessControl.FileSystemAccessRule]::new(
                $Identity,
                [System.Security.AccessControl.FileSystemRights]::ReadAndExecute,
                [System.Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit',
                [System.Security.AccessControl.PropagationFlags]::None,
                [System.Security.AccessControl.AccessControlType]::Allow
            )
        }
        'F' {
            return [System.Security.AccessControl.FileSystemAccessRule]::new(
                $Identity,
                [System.Security.AccessControl.FileSystemRights]::FullControl,
                [System.Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit',
                [System.Security.AccessControl.PropagationFlags]::None,
                [System.Security.AccessControl.AccessControlType]::Allow
            )
        }
        'M' {
            return [System.Security.AccessControl.FileSystemAccessRule]::new(
                $Identity,
                [System.Security.AccessControl.FileSystemRights]::Modify,
                [System.Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit',
                [System.Security.AccessControl.PropagationFlags]::None,
                [System.Security.AccessControl.AccessControlType]::Allow
            )
        }
    }
}
function ConvertTo-MatrixAclHC {
    <#
    .SYNOPSIS
        Convert the Excel sheet 'Permissions' to permission objects.

    .DESCRIPTION
        Convert the Excel sheet 'Permissions' to permission objects, by using
        the 'GroupName' and 'SiteCode' defined in the Excel sheet 'Settings'.

        Each object will contain the complete 'SamAccountName', the folder
        'Path' on the local machine and the type of access ('ACE') to that
        folder.

    .PARAMETER NonHeaderRows
        The objects coming from the Excel sheet 'Permissions', as retrieved by
        Import-Excel, but without the header columns. The header columns are
        replaced with ADObjects.

    .PARAMETER ADObjects
        A hashtable containing the property name and the SamAccountName
        belonging to that column.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param (
        [Parameter(Mandatory)]
        [PSCustomObject[]]$NonHeaderRows,

        [Parameter(Mandatory)]
        [hashtable]$ADObjects
    )

    begin {
        try {
            # Cache the column names ONCE instead of evaluating them every row
            $AllProperties = $NonHeaderRows[0].PSObject.Properties.Name
            $FirstProperty = $AllProperties[0]
            
            # Create an array of just the columns that contain permissions
            $PermColumns = $AllProperties | Select-Object -Skip 1
        }
        catch {
            throw "Failed initializing ConvertTo-MatrixAclHC: $_"
        }
    }

    process {
        try {
            for ($i = 0; $i -lt $NonHeaderRows.Count; $i++) {
                $Row = $NonHeaderRows[$i]
                $Path = $Row.$FirstProperty

                $Obj = [PSCustomObject]@{
                    Path   = $Path
                    Parent = ($i -eq 0)
                    Ignore = $false
                    ACL    = @{}
                }

                foreach ($ColName in $PermColumns) {
                    $Ace = $Row.$ColName

                    if ([string]::IsNullOrWhiteSpace($Ace)) { continue }

                    # If we hit an 'i' or 'I', set Ignore, clear any ACLs, and stop checking this row
                    if ($Ace -eq 'i' -or $Ace -eq 'I') {
                        $Obj.Ignore = $true
                        $Obj.ACL.Clear()
                        break
                    }

                    $SamAccountName = $ADObjects.($ColName).SamAccountName

                    if ([string]::IsNullOrWhiteSpace($SamAccountName)) {
                        throw "Missing AD Object for column $($ColName.TrimStart('P')) on folder path '$Path'."
                    }

                    if ($Obj.ACL.ContainsKey($SamAccountName)) {
                        throw "The AD object name '$SamAccountName' is not unique on folder path '$Path'."
                    }

                    $Obj.ACL.Add($SamAccountName, $Ace)
                }

                $Obj
            }
        }
        catch {
            throw "Failed converting to matrix ACL: $_"
        }
    }
}
function ConvertTo-MatrixADNamesHC {
    <#
    .SYNOPSIS
        Generate AD SamAccountNames from the first three rows in the Excel file.

    .DESCRIPTION
        Generate AD SamAccountNames from the first three rows in worksheet
        'Permissions' by replacing strings with the correct values.

        In case the value in A2 and B2 are equal, they are replaced by the
        string defined in 'Middle'. In case the value in A3 and B3 are equal,
        they are replaced by the value in 'Begin'.

        The template name to replace is always defined in cell A2 and A3 for
        their respective row.

    .PARAMETER ColumnHeaders
        The first 3 rows (objects) of the worksheet 'Permissions'. These objects
        contain the values to create the correct SamAccountNames.

    .PARAMETER Begin
        The value of the first part of the newly generated string. Usually this
        is the beginning of an AD GroupName like 'BEL ROL-AGG-SAGREX'.

    .PARAMETER Middle
        The value of the middle part of the newly generated string. Usually
        this is something like 'North'.
    #>

    [CmdletBinding()]
    [OutputType([hashtable])]
    param (
        [Parameter(Mandatory)]
        [ValidateCount(3, [int]::MaxValue)]
        [PSCustomObject[]]$ColumnHeaders,
        [String]$Begin,
        [String]$Middle,
        [String]$BeginReplace = 'GroupName',
        [String]$MiddleReplace = 'SiteCode'
    )

    process {
        try {
            Write-Verbose 'Converting to matrix AD object names'

            $Properties = $ColumnHeaders[0].PSObject.Properties.Name
            $FirstProperty = $Properties[0]
            $Result = @{}

            foreach ($Prop in $Properties) {
                # Skip the first column (usually the folder path/row headers)
                if ($Prop -eq $FirstProperty) { continue }

                Write-Verbose "Processing Property: '$Prop'"

                #region Get original values
                $EndVal = $ColumnHeaders[0].$Prop
                $MiddleVal = $ColumnHeaders[1].$Prop
                $BeginVal = $ColumnHeaders[2].$Prop

                $Original = [ordered]@{
                    Begin  = $BeginVal
                    Middle = $MiddleVal
                    End    = $EndVal
                }
                Write-Verbose "Original value begin '$BeginVal' middle '$MiddleVal' end '$EndVal'"
                #endregion

                #region Convert placeholder to proper values
                $ConvBegin = if ($BeginVal -eq $BeginReplace -and $Begin) { $Begin } else { $BeginVal }
                $ConvMiddle = if ($MiddleVal -eq $MiddleReplace -and $Middle) { $Middle } else { $MiddleVal }
                $ConvEnd = $EndVal

                $Converted = [ordered]@{
                    Begin  = $ConvBegin
                    Middle = $ConvMiddle
                    End    = $ConvEnd
                }
                Write-Verbose "Converted value begin '$ConvBegin' middle '$ConvMiddle' end '$ConvEnd'"
                #endregion

                #region Create SamAccountName
                # Filter out nulls/spaces and join
                $SamAccountName = (
                    $ConvBegin, $ConvMiddle, $ConvEnd | Where-Object { 
                        -not [string]::IsNullOrWhiteSpace($_) 
                    }
                ) -join ' '
                
                Write-Verbose "SamAccountName '$SamAccountName'"
                #endregion

                $Result.$Prop = @{
                    SamAccountName = $SamAccountName
                    Original       = $Original
                    Converted      = $Converted
                }
            }

            return $Result
        }
        catch {
            throw "Failed generating the correct AD object name for begin '$Begin' and middle '$Middle': $_"
        }
    }
}
function Get-DefaultAclHC {
    <#
    .SYNOPSIS
        Get the ACL from the default settings.

    .DESCRIPTION
        Retrieve the 'ADObjectName' and the 'Permission' properties and combine them into a hash table.
        Also tests if the permission characters are correct and throws an error in case they're not. The
        ADObjectName is not tested and needs to be checked afterwards.

    .PARAMETER Sheet
        The Excel worksheet containing the permission parameters.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param (
        [Parameter(Mandatory)]
        [PSCustomObject[]]$Sheet
    )

    process {
        try {
            $ACL = @{}

            foreach ($Row in $Sheet) {
                $ADObjectName = $Row.ADObjectName
                $Permission = $Row.Permission

                $HasName = -not [string]::IsNullOrWhiteSpace($ADObjectName)
                $HasPerm = -not [string]::IsNullOrWhiteSpace($Permission)

                if ((-not $HasName ) -and (-not $HasPerm)) {
                    continue
                }

                if (-not $HasPerm) {
                    throw "AD object name '$ADObjectName' has no permission."
                }

                if (-not $HasName) {
                    throw "Permission '$Permission' has no AD object name."
                }

                if ($Permission -notmatch '^(L|R|W|M|F)$') {
                    throw "Permission character '$Permission' is unknown."
                }

                if ($ACL.ContainsKey($ADObjectName)) {
                    throw "AD Object name '$ADObjectName' is not unique."
                }

                $ACL.Add($ADObjectName, $Permission)
            }

            return $ACL
        }
        catch {
            throw "Failed retrieving the ACL from the default settings file: $_"
        }
    }
}
function Format-PermissionsStringsHC {
    <#
    .SYNOPSIS
        String manipulations on values in the 'Permissions' sheet.

    .DESCRIPTION
        Remove leading and trailing spaces from strings, remove leading and
        trailing slashes from the path locations, change lower case permission
        characters to upper case, ...

    .PARAMETER Permissions
        Content of the Excel worksheet 'Permissions'.
    #>

    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [PSCustomObject[]]$Permissions
    )

    begin {
        $RowIndex = 0
        $FirstPropertyName = $null
    }

    process {
        foreach ($Row in $Permissions) {
            if ($null -eq $FirstPropertyName) {
                $FirstPropertyName = @($Row.PSObject.Properties.Name)[0]
            }

            foreach ($P in $Row.PSObject.Properties) {
                if (-not [string]::IsNullOrWhiteSpace($P.Value)) {
                    
                    $CleanValue = $P.Value.ToString().Trim()

                    if ($P.Name -eq $FirstPropertyName) {
                        $P.Value = $CleanValue.Trim('\')
                    } 
                    elseif ($RowIndex -ge 3) {
                        $P.Value = $CleanValue.ToUpper()
                    } 
                    else {
                        $P.Value = $CleanValue
                    }
                }
            }

            $Row
            $RowIndex++
        }
    }
}
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

    $ValidationErrors = [System.Collections.Generic.List[PSCustomObject]]::new()

    try {
        $Props = $Permissions[0].PSObject.Properties.Name
        $FirstProperty = $Props[0]

        #region Structural Validation (Fatal - Exits Immediately)
        if ($Permissions.Count -lt 4) {
            return [PSCustomObject]@{
                Type        = 'FatalError'
                Name        = 'Missing rows'
                Description = 'At least 4 rows are required: 3 header rows and 1 row for the parent folder.'
                Value       = "$($Permissions.Count) rows"
            }
        }

        if ($Props.Count -lt 2) {
            return [PSCustomObject]@{
                Type        = 'FatalError'
                Name        = 'Missing columns'
                Description = 'At least 2 columns are required: 1 for the folder names and 1 where the permissions are defined.'
                Value       = "$($Props.Count) column"
            }
        }
        #endregion

        #region Missing header SamAccountName
        foreach ($col in $Props) {
            if ([string]::IsNullOrWhiteSpace($Permissions[0].$col) -and 
                [string]::IsNullOrWhiteSpace($Permissions[1].$col) -and 
                [string]::IsNullOrWhiteSpace($Permissions[2].$col)) {
                
                $ValidationErrors.Add([PSCustomObject]@{
                        Type        = 'FatalError'
                        Name        = 'SamAccountName missing'
                        Description = 'Missing SamAccountName in the header row'
                        Value       = "Column number $($col.TrimStart('P'))"
                    })
            }
        }
        #endregion

        # Separate Headers from Data
        $NonHeaderRows = $Permissions | Select-Object -Skip 3
        $FolderNames = $NonHeaderRows | Select-Object -Skip 1

        #region Permission character unknown
        $InvalidChars = [System.Collections.Generic.List[string]]::new()
        
        foreach ($Row in $NonHeaderRows) {
            $PermColumns = $Row.PSObject.Properties.Where({ $_.Name -ne $FirstProperty })
            foreach ($Col in $PermColumns) {
                $Ace = $Col.Value
                if (-not [string]::IsNullOrWhiteSpace($Ace) -and $Ace -notmatch '^(L|R|W|I|C|F)$') {
                    $InvalidChars.Add($Ace)
                }
            }
        }

        if ($InvalidChars.Count -gt 0) {
            $ValidationErrors.Add([PSCustomObject]@{
                    Type        = 'FatalError'
                    Name        = 'Permission character unknown'
                    Description = "Supported characters are 'F', 'W', 'R', 'L', 'I', 'C', or blank."
                    Value       = ($InvalidChars | Select-Object -Unique) -join ', '
                })
        }
        #endregion

        #region Folder name missing
        $MissingFolders = $FolderNames.Where({ [string]::IsNullOrWhiteSpace($_.$FirstProperty) })
        if ($MissingFolders.Count -gt 0) {
            $ValidationErrors.Add([PSCustomObject]@{
                    Type        = 'FatalError'
                    Name        = 'Folder name missing'
                    Description = 'Missing folder name in the first column.'
                    Value       = "$($MissingFolders.Count) missing folder name(s)"
                })
        }
        #endregion

        #region Duplicate folder name
        $NotUniqueFolder = $FolderNames.$FirstProperty | Group-Object | Where-Object Count -GE 2
        if ($NotUniqueFolder) {
            $ValidationErrors.Add([PSCustomObject]@{
                    Type        = 'FatalError'
                    Name        = 'Duplicate folder name'
                    Description = 'Every folder name in the first column needs to be unique.'
                    Value       = ($NotUniqueFolder.Name) -join ', '
                })
        }
        #endregion

        #region Deepest folder has only List permissions or none at all
        $FolderRows = $Permissions | Select-Object -Skip 4
        $Paths = @($FolderRows.$FirstProperty)

        # Faster check for deepest folders
        $DeepestFolders = foreach ($P in $Paths) {
            if (-not ($Paths.Where({ $_ -ne $P -and $_ -like "$P\*" }))) {
                $P
            }
        }

        # Parent folder permissions (Row index 3)
        $ParentFolderPermissions = $Permissions[3].PSObject.Properties.Where({ 
                $_.Name -ne $FirstProperty -and -not [string]::IsNullOrWhiteSpace($_.Value) 
            }).Value

        $ParentFolderHasPermission = [bool]($ParentFolderPermissions.Where({ $_ -ne 'L' }))
        $inAccessibleFolders = [System.Collections.Generic.List[string]]::new()

        foreach ($Row in $FolderRows.Where({ $_.$FirstProperty -in $DeepestFolders })) {
            $Perms = $Row.PSObject.Properties.Where({
                    $_.Name -ne $FirstProperty -and 
                    -not [string]::IsNullOrWhiteSpace($_.Value) -and 
                    $_.Value -ne 'L'
                }).Value

            if ((-not $Perms) -and (-not $ParentFolderHasPermission)) {
                $inAccessibleFolders.Add($Row.$FirstProperty)
            }
        }

        if ($inAccessibleFolders.Count -gt 0) {
            $ValidationErrors.Add([PSCustomObject]@{
                    Type        = 'Warning'
                    Name        = 'Matrix design flaw'
                    Description = 'All folders need to be accessible by the end user. Please define at least (R)ead or (W)rite on the deepest folder.'
                    Value       = $inAccessibleFolders -join ', '
                })
        }
        #endregion

        # Output all collected errors at the end
        if ($ValidationErrors.Count -gt 0) {
            return $ValidationErrors
        }

    }
    catch {
        throw "Failed testing the Excel sheet 'Permissions' for incorrect data: $_"
    }
}