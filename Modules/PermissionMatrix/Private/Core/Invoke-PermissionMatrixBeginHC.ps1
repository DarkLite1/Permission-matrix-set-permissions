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
        # =====================================================================
        # 1. SEQUENTIAL: Test input file & create context
        # =====================================================================
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

        # =====================================================================
        # 2. PARALLEL: Read, Validate, and Archive Matrix Files
        # =====================================================================
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

        #region Exclude Defaults file from the processing list
        $matrixFiles = $matrixFiles | Where-Object { 
            $_.FullName -ne $Context.Config.Matrix.DefaultsFile 
        }
        #endregion

        if (-not $matrixFiles -or $matrixFiles.Count -eq 0) {
            return $Context # No files found, exit BEGIN gracefully
        }
        
        $Context.FoundMatrices = $true

        #region Setup Archive Folder
        $archivePath = $null
        if ($Context.Config.Matrix.Archive) {
            $archivePath = Join-Path -Path $Context.Config.Matrix.FolderPath -ChildPath 'Archive'
            if (-not (Test-Path -LiteralPath $archivePath -PathType Container)) {
                $null = New-Item -ItemType Directory -Path $archivePath -Force -ErrorAction SilentlyContinue
            }
        }
        #endregion

        $throttle = $Context.Config.MaxConcurrent.FoldersPerMatrix ?? 4

        #region Import, validate and archive in Parallel
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
                    foreach ($m in $fileResult.Matrices) {
                        $rowErrors = Test-MatrixSettingRowHC `
                            -SettingRow $m.Setting.Raw `
                            -RequireGroupName $reqGroupName `
                            -RequireSiteCode $reqSiteCode

                        if ($rowErrors) { 
                            $m.Check.AddRange(
                                [pscustomobject[]]@($rowErrors)
                            ) 
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

        # =====================================================================
        # 3. SEQUENTIAL: Cross-Matrix Checks & AD Lookups
        # =====================================================================
        
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

        #region Read Defaults Excel file and validate
        $defaults = Import-MatrixDefaultsHC `
            -Matrix $Context.Config.Matrix `
            -SystemErrors $SystemErrors

        if (Test-HasFatalErrorsHC $SystemErrors) { return $Context }

        $Context.Defaults = $defaults
        #endregion

        # 3c. One AD query for all objects combined
        $allAdObjects = $Context.AllMatrices.Settings.Matrix.ACL.Keys | Sort-Object -Unique

        if ($allAdObjects.Count -gt 0) {
            $adObjectDetails = @(
                Get-ADObjectDetailHC `
                    -ADObjectName $allAdObjects `
                    -Type 'SamAccountName'
            )
            
            # 3d. Combine AD info with matrix data (Expanded Matrix Validation)
            foreach ($matrixObj in $Context.AllMatrices) {
                foreach ($S in $matrixObj.Settings) {
                    if (-not $S.Matrix) { continue }

                    if ($Context.Defaults.DefaultAcl) {
                        try {
                            $applyDefaultPerms = [System.Convert]::ToBoolean($S.Setting.Formatted.ApplyDefaultPermissions ?? $false)
                            
                            $S.Matrix.ACL = Merge-DefaultPermissionsHC `
                                -Defaults $Context.Defaults.DefaultAcl `
                                -Matrix $S.Matrix.ACL `
                                -ApplyDefaultPermissions $applyDefaultPerms
                        }
                        catch {
                            $S.Check += [PSCustomObject]@{
                                Type        = 'FatalError'
                                Name        = 'Defaults Conflict'
                                Description = 'When ApplyDefaultPermissions is enabled, the matrix cannot explicitly define AD Objects that are already managed by the defaults.'
                                Value       = $_.Exception.Message
                            }
                            
                            continue 
                        }
                    }

                    $expandedCheck = Test-ExpandedMatrixHC `
                        -Matrix $S.Matrix `
                        -ADObject $adObjectDetails `
                        -ExcludedSamAccountName $Context.Config.Matrix.ExcludedSamAccountName

                    if ($expandedCheck) {
                        $S.Check += $expandedCheck | 
                        ConvertTo-StructuredObjectHC
                    }
                }
            }
        }

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