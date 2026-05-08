function Invoke-PermissionMatrixBeginHC {
    <#
    .SYNOPSIS
        BEGIN stage for the Permission Matrix pipeline.
    .DESCRIPTION
        1. Sequential: Validates JSON.
        2. Parallel: Reads, validates, and archives Matrix Excel files.
        3. Sequential: Checks for cross-matrix duplicates, loads Defaults, and performs bulk AD queries.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ConfigurationJsonFile,

        [Parameter(Mandatory)]
        [hashtable]$ScriptPath,

        [Parameter(Mandatory)]
        [ref]$SystemErrors
    )

    try {
        #region Get JSON content
        if (-not (Test-Path -LiteralPath $ConfigurationJsonFile -PathType Leaf)) {
            Add-ErrorHC `
                -Type 'FatalError' `
                -Name 'Configuration file not found' `
                -Message "File '$ConfigurationJsonFile' does not exist." `
                -Category 'RuntimeSettings' `
                -SystemErrors $SystemErrors
            return $null
        }

        $json = Get-Content -LiteralPath $ConfigurationJsonFile -Raw -Encoding UTF8 | ConvertFrom-Json -Depth 50
        #endregion

        $Context = [pscustomobject]@{
            JsonFileName  = [System.IO.Path]::GetFileNameWithoutExtension($ConfigurationJsonFile)
            Config        = $json
            ScriptPath    = $ScriptPath
            StartTime     = Get-Date
            Counter       = New-CounterObjectHC
            ExportedFiles = @{}
            FoundMatrices = $false
            FileResults   = @()
            AllMatrices   = @()
            Defaults      = $null
        }

        #region Validate Configuration Structure
        Validate-ConfigurationStructureHC `
            -Json $json `
            -SystemErrors $SystemErrors
        #endregion

        #region Validate Script Paths
        foreach ($key in $ScriptPath.Keys) {
            $path = $ScriptPath[$key]
            if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
                Add-ErrorHC `
                    -Type 'FatalError' `
                    -Name 'Missing Script File' `
                    -Message "The required script '$key' was not found at '$path'." `
                    -Category 'RuntimeSettings' `
                    -SystemErrors $SystemErrors
            }
        }
        #endregion

        if ($SystemErrors.Value.Count -gt 0) { return $Context }
        
        #region Get Matrix Files
        try {
            $matrixFiles = Get-ChildItem -Path $Context.Config.Matrix.FolderPath -Filter '*.xlsx' -File -ErrorAction Stop
        }
        catch {
            Add-ErrorHC `
                -Type 'FatalError' `
                -Name 'Matrix folder access failed' `
                -Message "Cannot access '$($Context.Config.Matrix.FolderPath)'." `
                -Category 'Matrix' `
                -SystemErrors $SystemErrors
            return $Context
        }
        
        $matrixFiles = $matrixFiles | Where-Object { 
            $_.FullName -ne $Context.Config.Matrix.DefaultsFile 
        }
        
        if (-not $matrixFiles -or $matrixFiles.Count -eq 0) {
            return $Context 
        }

        $Context.FoundMatrices = $true
        #endregion

        #region Read Defaults Excel file and validate (Placed here to save I/O)
        $defaults = Import-MatrixDefaultsHC `
            -Matrix $Context.Config.Matrix `
            -SystemErrors $SystemErrors

        if (Test-ItemHasFatalErrorHC -CheckList $SystemErrors.Value) { 
            return $Context 
        }

        $Context.Defaults = $defaults
        #endregion

        #region Create Archive Folder
        $archivePath = $null
        if ($Context.Config.Matrix.Archive) {
            $archivePath = Join-Path -Path $Context.Config.Matrix.FolderPath -ChildPath 'Archive'
            if (-not (Test-Path -LiteralPath $archivePath -PathType Container)) {
                $null = New-Item -ItemType Directory -Path $archivePath -Force -ErrorAction SilentlyContinue
            }
        }
        #endregion

        #region Import, validate and archive in Parallel   
        $throttle = $Context.Config.MaxConcurrent.FoldersPerMatrix ?? 4

        $parallelResults = Invoke-WithOptionalParallelismHC `
            -InputObject $matrixFiles `
            -ThrottleLimit $throttle `
            -ArgumentList $Context, $archivePath `
            -ScriptBlock {
            param($file, $context, $archiveFolder)

            #region Load module and helper functions inside runspace
            Import-Module ImportExcel -ErrorAction Stop

            $privateFolder = Join-Path `
                -Path (Split-Path $context.ScriptPath.PermissionMatrixModule) `
                -ChildPath 'Private'

            Get-ChildItem -Path $privateFolder -Recurse -Filter '*.ps1' | 
            ForEach-Object { . $_.FullName }
            #endregion
            
            try {
                #region Import & validate matrix
                $fileResult = Import-MatrixFileHC `
                    -MatrixFile $file `
                    -Context $context

                $reqGroupName = $false
                $reqSiteCode = $false

                if ($fileResult.Sheets.Permissions.Raw) {
                    #region Check if GroupName and SiteCode columns are required
                    $headerRows = $fileResult.Sheets.Permissions.Raw | 
                    Select-Object -First 3

                    foreach ($row in $headerRows) {
                        foreach ($p in $row.PSObject.Properties) {
                            Write-Verbose "Checking Permissions header: $($p.Name) = '$($p.Value)'"
                            if ($p.Value -is [string]) {
                                if ($p.Value -match 'GroupName') { 
                                    $reqGroupName = $true 
                                }
                                if ($p.Value -match 'SiteCode') { 
                                    $reqSiteCode = $true 
                                }
                            }
                        }
                    }
                    #endregion

                    #region Validate Permissions and add any errors to the file result
                    if ($fileResult.Sheets.Permissions.Formatted) {
                        $permErrors = Test-MatrixPermissionsHC `
                            -Permissions $fileResult.Sheets.Permissions.Formatted
    
                        if ($permErrors) {
                            $fileResult.Check.AddRange(
                                [pscustomobject[]]@($permErrors)
                            )
                        }
                    }
                    #endregion
                }


                if ($fileResult.Matrices) {
                    $permSheet = $fileResult.Sheets.Permissions.Formatted
                    $dataRows = if ($permSheet) { $permSheet | Select-Object -Skip 4 } else { @() }

                    foreach ($m in $fileResult.Matrices) {
                        $rowErrors = Test-MatrixSettingRowHC `
                            -SettingRow $m.Setting.Raw `
                            -RequireGroupName $reqGroupName `
                            -RequireSiteCode $reqSiteCode

                        if ($rowErrors) { 
                            $m.Check.AddRange([pscustomobject[]]@($rowErrors)) 
                        }

                        $isFileBroken = Test-ItemHasFatalErrorHC `
                            -CheckList $fileResult.Check
                        $isRowBroken = Test-ItemHasFatalErrorHC `
                            -CheckList $m.Check

                        if (
                            -not $isFileBroken -and 
                            -not $isRowBroken -and
                            $permSheet
                        ) {
                            # A. Extract and Map AD Objects
                            $adMap = Get-MatrixADObjectsMapHC `
                                -PermissionsSheet $permSheet `
                                -SettingRow $m.Setting.Formatted

                            # B. Build the Matrix ACLs
                            $m.Matrix = ConvertTo-MatrixAclHC `
                                -DataRows $dataRows `
                                -AdObjectsMap $adMap

                            # C. Merge Defaults per Folder
                            if ($context.Defaults.DefaultAcl) {
                                try {
                                    $applyDefaultPerms = [System.Convert]::ToBoolean($m.Setting.Formatted.ApplyDefaultPermissions ?? $false)
                                    
                                    foreach ($folder in $m.Matrix) {
                                        $folder.ACL = Merge-DefaultPermissionsHC `
                                            -Defaults $context.Defaults.DefaultAcl `
                                            -MatrixAcl $folder.ACL `
                                            -ApplyDefaultPermissions $applyDefaultPerms
                                    }
                                }
                                catch {
                                    $m.Check.Add(
                                        [PSCustomObject]@{
                                            Type        = 'FatalError'
                                            Name        = 'Defaults Conflict'
                                            Description = 'When ApplyDefaultPermissions is enabled, the matrix cannot explicitly define AD Objects already managed by defaults.'
                                            Value       = $_.Exception.Message
                                        }
                                    )
                                }
                            }
                        }
                    }
                }
                #endregion
            }
            catch {
                if (-not $fileResult) {
                    $fileResult = [pscustomobject]@{
                        File     = $file
                        Check    = [System.Collections.Generic.List[pscustomobject]]::new()
                        Matrices = [System.Collections.Generic.List[pscustomobject]]::new()
                    }
                }
                
                $fileResult.Check.Add(
                    [pscustomobject]@{
                        Type        = 'FatalError'
                        Name        = 'Runspace processing failed'
                        Description = 'An unexpected terminating error occurred during I/O or Validation.'
                        Value       = $_
                    }
                )
            }
            finally {
                #region Archive file
                if ($archiveFolder) {
                    try {
                        $destination = Join-Path -Path $archiveFolder -ChildPath $file.Name
                        Move-Item -LiteralPath $file.FullName -Destination $destination -Force -ErrorAction Stop
                    }
                    catch {
                        $fileResult.File.Check.Add(
                            [pscustomobject]@{
                                Type        = 'Warning' 
                                Name        = 'Archiving failed'
                                Description = 'File could not be moved to archive.'
                                Value       = $_
                            })
                    }
                }
                #endregion
    
                $fileResult
            }
        }
        #endregion

        #region Collect results and store in context
        $Context.FileResults = $parallelResults

        $importedMatrices = [System.Collections.Generic.List[pscustomobject]]::new()
        foreach ($res in $parallelResults) {
            if ($res.Matrices) {
                $importedMatrices.AddRange(
                    [pscustomobject[]]@($res.Matrices)
                )
            }
        }
        $Context.AllMatrices = $importedMatrices
        #endregion

        #region Duplicate ComputerName/Path Validation
        $duplicateMatrices = $Context.AllMatrices | 
        Group-Object -Property { $_.Setting.Formatted.ComputerName }, { $_.Setting.Formatted.Path } | 
        Where-Object Count -GE 2

        foreach ($DupGroup in $duplicateMatrices) {
            foreach ($MatrixObj in $DupGroup.Group) {
                $conflictingFiles = ($DupGroup.Group | ForEach-Object { $_.FileContext.Item.Name }) | Select-Object -Unique
                $fileListString = $conflictingFiles -join "', '"

                $MatrixObj.Check.Add(
                    [PSCustomObject]@{
                        Type        = 'FatalError'
                        Name        = 'Duplicate ComputerName/Path'
                        Description = "Multiple settings across the matrices have the same 'ComputerName' and 'Path' combination, which can lead to conflicts during permission application."
                        Value       = "File '$fileListString', ComputerName '$($MatrixObj.Setting.Formatted.ComputerName)', Path '$($MatrixObj.Setting.Formatted.Path)'"
                    }
                )
            }
        }
        #endregion

        #region Get all AD Objects from matrices and defaults
        $allAdObjects = [System.Collections.Generic.List[string]]::new()
        
        foreach ($matrixObj in $Context.AllMatrices) {
            foreach ($folder in $matrixObj.Matrix) {
                if ($folder.ACL) {
                    $allAdObjects.AddRange([string[]]@($folder.ACL.Keys))
                }
            }
        }
        
        if ($Context.Defaults.DefaultAcl) {
            $allAdObjects.AddRange(
                [string[]]@($Context.Defaults.DefaultAcl.Keys)
            )
        }

        $uniqueAdObjects = $allAdObjects | Sort-Object -Unique
        #endregion

        #region Bulk query AD for all unique objects and build a name → SID map
        if ($uniqueAdObjects.Count -gt 0) {
            #region Bulk AD Lookup
            $adObjectDetails = @()
            try {
                $adObjectDetails = @(
                    Get-ADObjectDetailHC `
                        -ADObjectName $uniqueAdObjects `
                        -Type 'SamAccountName' `
                        -ErrorAction Stop
                )
            }
            catch {
                Add-ErrorHC `
                    -Type 'Warning' `
                    -Name 'AD Bulk Lookup Failure' `
                    -Message "Failed during bulk AD lookup. Some AD objects may be marked as unknown. Error: $_" `
                    -Category 'ExpandedMatrix' `
                    -SystemErrors $SystemErrors
            }
            #endregion

            #region Build Name → SID map for quick lookup during ACL rewrite
            $nameToSid = @{}
            foreach ($detail in $adObjectDetails) {
                if ($detail.adObject -and $detail.adObject.ObjectSid) {
                    $nameToSid[$detail.SamAccountName] = $detail.adObject.ObjectSid
                }
            }
            #endregion

            #region Rewrite ACLs in all matrices to use SIDs instead of names
            foreach ($matrixObj in $Context.AllMatrices) {
                $isFileBroken = Test-ItemHasFatalErrorHC `
                    -CheckList $matrixObj.FileContext.Check
                $isRowBroken = Test-ItemHasFatalErrorHC `
                    -CheckList $matrixObj.Check

                if ($isFileBroken -or $isRowBroken) {
                    continue
                }

                $adObjectCheck = Test-AdObjectInMatrixHC `
                    -Matrix $matrixObj.Matrix `
                    -ADObject $adObjectDetails

                if ($adObjectCheck) {
                    $matrixObj.Check.AddRange(
                        [pscustomobject[]]@($adObjectCheck)
                    )
                    # If validation flagged a fatal error, skip the SID rewrite for this matrix
                    if (Test-ItemHasFatalErrorHC -CheckList $matrixObj.Check) {
                        continue
                    }
                }

                # Add SID rewrite as a final step after all checks to ensure we have the necessary details for accurate error reporting
                foreach ($folder in $matrixObj.Matrix) {
                    if (-not $folder.ACL -or $folder.ACL.Count -eq 0) { continue }

                    $newAcl = @{}
                    $adNames = @{}
                    foreach ($name in @($folder.ACL.Keys)) {
                        $sid = $nameToSid[$name]
                        if ($sid) {
                            $newAcl[$sid] = $folder.ACL[$name]
                            $adNames[$sid] = $name
                        }
                    }
                    $folder.ACL = $newAcl
                    $folder | Add-Member `
                        -NotePropertyName 'AdNames' `
                        -NotePropertyValue $adNames -Force
                }
            }
            #endregion
        }
        #endregion

        return $Context
    }
    catch {
        Add-ErrorHC `
            -Type 'FatalError' `
            -Category 'Runtime' `
            -Name 'BEGIN stage failure' `
            -Message "Unhandled exception: $_" `
            -SystemErrors $SystemErrors
        return $null
    }
}