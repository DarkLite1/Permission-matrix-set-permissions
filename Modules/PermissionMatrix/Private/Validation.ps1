function ConvertTo-StructuredObjectHC {
    <#
    .SYNOPSIS
        Normalize mixed pipeline input into structured records, wrapping strings
        and unknown objects and passing structured objects through.

    .DESCRIPTION
        Takes a stream of arbitrary objects and emits a structured record for
        each, so a mixed pipeline (strings, hashtables, custom objects, other
        types) becomes a uniform sequence of objects downstream code can handle
        consistently.

        Each input item is classified as follows:

        - $null: skipped, producing no output.
        - [string]: wrapped via New-ValidationCheckHC with Type 'Information'
          and Name 'Message', the string becoming the record's Description.
        - [hashtable] or [pscustomobject]: passed through unchanged, on the
          assumption it is already a structured record.
        - Anything else: stringified and wrapped via New-ValidationCheckHC with
          Type 'Information' and Name 'UnknownObject', the string form becoming
          the record's Description.

        The function processes pipeline input one item at a time and also
        iterates the items of any array passed as a single argument.

    .PARAMETER InputObject
        The object(s) to normalize. Accepts pipeline input. Each item is
        classified and emitted individually; $null items are dropped. Mandatory.

    .EXAMPLE
        'something happened' | ConvertTo-StructuredObjectHC

        Emits a validation-check record: Type 'Information', Name 'Message',
        Description 'something happened'.

    .EXAMPLE
        @(
            'a message',
            [pscustomobject]@{ Type = 'Warning'; Name = 'X' },
            42,
            $null
        ) | ConvertTo-StructuredObjectHC

        Emits three records: the string is wrapped as a 'Message', the
        PSCustomObject passes through unchanged, 42 is wrapped as an
        'UnknownObject' with Description '42', and the $null is skipped.

    .EXAMPLE
        Some-Step | ConvertTo-StructuredObjectHC | Where-Object Type -eq 'Information'

        Normalizes whatever Some-Step emits (free-form strings, ready-made
        records, or other values) so the downstream filter can rely on a
        consistent record shape.

    .OUTPUTS
        System.Management.Automation.PSCustomObject
        For strings and unrecognized types, a record from New-ValidationCheckHC.
        For hashtables and PSCustomObjects, the original object unchanged. No
        output is produced for $null items.

    .NOTES
        - $null items are silently dropped.
        - Strings and unknown types are wrapped with Type 'Information'; the
          difference is the Name ('Message' vs 'UnknownObject'). Note an unknown
          object is recorded as 'Information', not as a warning, even though it
          was an unexpected type.
        - Hashtables are passed through as-is and are not converted to
          PSCustomObjects or validated; downstream code receiving a [hashtable]
          alongside [pscustomobject] records should be ready for both shapes.
        - The wrapped records carry only Description (no Value); their Value and
          Category fields are $null.

    .LINK
        New-ValidationCheckHC
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true)] 
        $InputObject
    )

    process {
        foreach ($obj in $InputObject) {
            
            if ($null -eq $obj) { continue }

            if ($obj -is [string]) {
                New-ValidationCheckHC `
                    -Type 'Information' `
                    -Name 'Message' `
                    -Description $obj
                continue
            }

            if ($obj -is [hashtable] -or $obj -is [pscustomobject]) {
                $obj
                continue
            }

            New-ValidationCheckHC `
                -Type 'Information' `
                -Name 'UnknownObject' `
                -Description "$obj"
        }
    }
}

function Test-MatrixFileHC {
    [CmdletBinding()]
    param([Parameter(Mandatory)] $MatrixObject)

    $checks = @()

    if (-not $MatrixObject.Settings -or $MatrixObject.Settings.Count -eq 0) {
        $checks += New-ValidationCheckHC -Type 'Warning' -Name 'Matrix disabled' `
            -Description 'No Settings rows found.' -Category 'File'
    }

    if (-not $MatrixObject.Permissions -or $MatrixObject.Permissions.Count -eq 0) {
        $checks += New-ValidationCheckHC `
            -Type 'FatalError' `
            -Name 'Missing Permissions sheet' `
            -Description 'Permissions sheet missing or empty.' `
            -Category 'File'
    }

    return $checks
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
        $FolderNames = $NonHeaderRows | Select-Object -Skip 1

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
        $MissingFolders = $FolderNames.Where(
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
        $NotUniqueFolder = $FolderNames.$FirstProperty | Group-Object | Where-Object Count -GE 2
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
            $checks.Add(
                (New-ValidationCheckHC `
                    -Type 'Warning' `
                    -Name 'Inaccessible folders' `
                    -Description 'The deepest folders have no permissions or only List permissions, and the parent folder does not have permissions that allow access. This means these folders will be inaccessible.' `
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

    process {
        try {
            #region No FormData -> Warning
            if ((-not $FormData) -or ($FormData.Count -eq 0)) {
                return [PSCustomObject]@{
                    Type        = 'Warning'
                    Name        = 'Missing FormData'
                    Description = 'No FormData rows were found. ServiceNow form data will not be exported for this matrix file.'
                    Value       = $null
                }
            }
            #endregion

            if ($FormData.Count -ne 1) {
                return [PSCustomObject]@{
                    Type        = 'FatalError'
                    Name        = 'Incorrect row count'
                    Description = "Exactly one row of data is required. Found $($FormData.Count) row(s)."
                    Value       = $FormData.Count
                }
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
                return [PSCustomObject]@{
                    Type        = 'FatalError'
                    Name        = 'Missing column header'
                    Description = "The following column headers are mandatory: $($MandatoryProperties -join ', ')."
                    Value       = $MissingProperties -join ', '
                }
            }
            #endregion

            #region Mandatory property values (Only if Enabled)
            if ($Row.MatrixFormStatus -eq 'Enabled') {

                $MandatoryPropertyValues = $MandatoryProperties.Where({ $_ -ne 'MatrixFormStatus' })

                $BlankProperties = $MandatoryPropertyValues.Where({
                        [string]::IsNullOrWhiteSpace($Row.$_)
                    })

                if ($BlankProperties) {
                    return [PSCustomObject]@{
                        Type        = 'FatalError'
                        Name        = 'Missing value'
                        Description = "Values for the following columns are mandatory when status is Enabled: $($MandatoryPropertyValues -join ', ')."
                        Value       = $BlankProperties -join ', '
                    }
                }
            }
            #endregion
        }
        catch {
            throw "Failed testing the Excel sheet 'FormData': $_"
        }
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

function Test-AdObjectsHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$ADObjects,
        [Parameter(Mandatory)]       $AdInfo
    )

    $checks = @()

    foreach ($obj in $ADObjects) {
        if ($obj -notin $AdInfo) {
            $checks += New-ValidationCheckHC `
                -Type 'Warning' `
                -Name 'Missing AD Object' `
                -Description "AD object '$obj' not found." `
                -Category 'AD'
        }
    }

    return $checks
}

function Test-AdObjectInMatrixHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$Matrix,
        [Parameter(Mandatory)]$ADObject
    )

    $checks = @()

    $matrixAdObjects = @($Matrix.ACL.Keys) | Select-Object -Unique

    if (-not $matrixAdObjects) { return $checks }

    $missingAdObjects = $matrixAdObjects | Where-Object { 
        $name = $_
        $match = $ADObject | Where-Object { $_.SamAccountName -eq $name }
        $null -eq $match.adObject 
    }

    if ($missingAdObjects) {
        $checks += New-ValidationCheckHC `
            -Type 'FatalError' `
            -Name ' Unknown AD Objects in Matrix' `
            -Description 'One or more AD objects referenced in the matrix were not found in Active Directory. Please check the SamAccountName values in the Permissions sheet and ensure they exist in AD.' `
            -Value "Not existing AD Objects: $($missingAdObjects -join ', ')"
    }

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
        if ($Json.Settings.SaveLogFiles.Detailed -isnot [bool]) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Incorrect 'Settings.SaveLogFiles.Detailed'" `
                -Message 'Must be boolean.' `
                -SystemErrors $SystemErrors
        }

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
        foreach ($prop in 'Computers', 'FoldersPerMatrix', 'JobsPerRemoteComputer') {
            $val = $Json.MaxConcurrent.$prop
            if ($null -eq $val -or $val -notmatch '^\d+$') {
                Add-JsonSchemaErrorHC -Type 'FatalError' `
                    -Name "Incorrect 'MaxConcurrent.$prop'" `
                    -Message "Property 'MaxConcurrent.$prop' must be numeric." `
                    -SystemErrors $SystemErrors
            }
        }
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
}