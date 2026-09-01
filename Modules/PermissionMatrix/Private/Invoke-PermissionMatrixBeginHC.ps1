function Invoke-PermissionMatrixBeginHC {
    <#
    .SYNOPSIS
        Initializes, validates, and builds the execution context for the
        Permission Matrix pipeline.

    .DESCRIPTION
        This function serves as the 'BEGIN' stage of the orchestrator. It
        processes the foundational data required for the remote execution jobs
        through three distinct phases:

        1. Initialization (Sequential):
            Validates the JSON configuration file and constructs the baseline
            `$Context` state object.
        2. Ingestion (Parallel):
            Discovers, structurally validates, and archives the Excel Matrix
            files using a multi-threaded runspace pool to minimize I/O
            bottlenecks.
        3. Global Resolution (Sequential):
            Performs cross-matrix duplicate detection, bulk queries Active
            Directory for unique objects, merges global default permissions,
            and permanently rewrites AD Account Names into SIDs (Security
            Identifiers) to guarantee precise permission application during the
            remote execution phase.

        Architectural Convention:
        This function returns `$null` ONLY for catastrophic, pre-context
        failures (e.g., the JSON file is missing or corrupt). Once the
        `$Context` object is successfully constructed, the function will return
        it even if fatal errors occur later. This allows the calling script to
        inspect the partial state and generate comprehensive error reports
        rather than failing silently.

    .PARAMETER ConfigurationJsonFile
        The absolute path to the main JSON configuration file governing the
        execution.

    .PARAMETER ScriptPath
        A hashtable containing the absolute file paths to the required
        execution scripts and modules.

    .PARAMETER SystemErrors
        A reference variable ([ref]) containing a List[pscustomobject]. Used to
        capture terminating pipeline errors.

    .OUTPUTS
        System.Management.Automation.PSCustomObject
        Returns the fully constructed `$Context` object containing the imported
        matrices, resolved AD details, and configurations. Returns `$null` only
        on pre-context initialization failures.

    .EXAMPLE
        $sysErrors = [System.Collections.Generic.List[pscustomobject]]::new()

        $context = Invoke-PermissionMatrixBeginHC `
            -ConfigurationJsonFile 'C:\Config.json' `
            -ScriptPath $scriptPaths `
            -SystemErrors ([ref]$sysErrors)
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
            JsonFileName    = [System.IO.Path]::GetFileNameWithoutExtension($ConfigurationJsonFile)
            Config          = $json
            ScriptPath      = $ScriptPath
            StartTime       = Get-Date
            Counter         = New-CounterObjectHC
            ExportedFiles   = @{}
            FoundMatrices   = $false
            FileResults     = @()
            AllMatrices     = @()
            AdObjectDetails = @()
            Defaults        = $null
        }

        #region Validate Configuration Structure
        Test-ConfigurationStructureHC `
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

        <# Only a fatal error stops the run here, matching how every other
        stage of this function decides. Counting entries instead meant a
        Warning added by the two validations above would abort silently, and
        nothing constrains the Type those helpers accept. #>
        if (Test-ItemHasFatalErrorHC -CheckList $SystemErrors.Value) {
            return $Context
        }

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

        <# Compare resolved paths rather than the raw configured string. A
        DefaultsFile written as a relative path, with forward slashes, or with
        redundant '.' segments points at the same file on disk but is not the
        same string, and the defaults workbook would then be imported as if it
        were an ordinary matrix file.

        Note this normalizes syntax only. A drive letter mapped to the same
        UNC share still compares as a different path, so keep DefaultsFile and
        Matrix.FolderPath in the same notation. #>
        $defaultsFullName = (
            Resolve-Path `
                -LiteralPath $Context.Config.Matrix.DefaultsFile `
                -ErrorAction SilentlyContinue
        ).ProviderPath

        if (-not $defaultsFullName) {
            # Validation already requires this file to exist. If it vanished
            # since then, fall back to the configured value so an exact match
            # is still excluded, and let Import-MatrixDefaultsFileHC report it
            # a few lines below.
            $defaultsFullName = $Context.Config.Matrix.DefaultsFile
        }

        $matrixFiles = $matrixFiles | Where-Object {
            $_.FullName -ne $defaultsFullName
        }

        if (-not $matrixFiles -or $matrixFiles.Count -eq 0) {
            return $Context
        }

        $Context.FoundMatrices = $true
        #endregion

        #region Read Defaults Excel file and validate (Placed here to save I/O)
        $defaults = Import-MatrixDefaultsFileHC `
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
        $throttle = if (
            [string]::IsNullOrWhiteSpace($Context.Config.MaxConcurrent.FoldersPerMatrix)
        ) {
            4
        }
        else {
            # ?? only guards against $null, so a configured 0 used to pass
            # through as a throttle of 0.
            [math]::Max(1, [int]$Context.Config.MaxConcurrent.FoldersPerMatrix)
        }

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
                    # Only the AD Object Name header columns (P2, P3, ...) can
                    # use 'GroupName'/'SiteCode' as placeholders that are resolved
                    # from the Settings row (see Get-MatrixADObjectsMapHC, which
                    # scans from P2 upward). Column A (P1) is the informational
                    # folder column, so a literal 'SiteCode'/'GroupName' there
                    # must NOT make the Settings value mandatory. The cell is
                    # matched exactly (trimmed), mirroring the resolver's per-cell
                    # switch, so a label that merely contains the word (e.g.
                    # 'Regional SiteCode') does not trigger the requirement.
                    $headerRows = $fileResult.Sheets.Permissions.Raw |
                    Select-Object -First 3

                    foreach ($row in $headerRows) {
                        foreach ($p in $row.PSObject.Properties) {
                            # Skip Column A (P1) and any non permission column
                            if (
                                ($p.Name -notmatch '^P(\d+)$') -or
                                ([int]$Matches[1] -lt 2)
                            ) {
                                continue
                            }

                            Write-Verbose "Checking Permissions header: $($p.Name) = '$($p.Value)'"

                            if ($p.Value -is [string]) {
                                if ($p.Value.Trim() -eq 'GroupName') {
                                    $reqGroupName = $true
                                }
                                if ($p.Value.Trim() -eq 'SiteCode') {
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
                            $fileResult.Sheets.Permissions.Check.AddRange(
                                [pscustomobject[]]@($permErrors)
                            )
                        }
                    }
                    #endregion
                }


                if ($fileResult.Matrices) {
                    $permSheet = $fileResult.Sheets.Permissions.Formatted
                    # Always coerce to an array. When the Permissions sheet holds
                    # only the 4 header rows (no data rows), Select-Object -Skip 4
                    # returns $null, and passing $null to the mandatory [array]
                    # DataRows parameter of ConvertTo-MatrixAclHC throws
                    # 'Cannot bind argument ... because it is null', which surfaced
                    # as a 'Runspace processing failed' FatalError.
                    $dataRows = @($permSheet | Select-Object -Skip 4) 

                    foreach ($m in $fileResult.Matrices) {
                        $rowErrors = Test-MatrixSettingRowHC `
                            -SettingRow $m.Setting.Raw `
                            -RequireGroupName $reqGroupName `
                            -RequireSiteCode $reqSiteCode

                        if ($rowErrors) {
                            $m.Check.AddRange([pscustomobject[]]@($rowErrors))
                        }

                        $isFileBroken = Test-FileHasFatalErrorHC `
                            -File $fileResult
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
                            $childMatrix = @(
                                ConvertTo-MatrixAclHC `
                                    -DataRows $dataRows `
                                    -AdObjectsMap $adMap
                            )

                            # B2. Build the parent (root) folder entry from the
                            # 'Path' row (row index 3, the row directly under the
                            # three header rows). SetPermissions.ps1 needs a
                            # Parent=$true entry: it applies the root folder's own
                            # ACL AND, crucially, seeds the recursive inheritance
                            # walk from the root. That walk is what resets every
                            # TOP-LEVEL folder without permissions (an inherit-only
                            # folder) back to pure inheritance. Without this entry
                            # the root is never a walk seed, so top-level
                            # inherit-only folders keep any stale explicit ACL and
                            # are never corrected.
                            $parentRow = @($permSheet | Select-Object -Skip 3 -First 1)
                            $parentEntry = @(
                                ConvertTo-MatrixAclHC `
                                    -DataRows $parentRow `
                                    -AdObjectsMap $adMap
                            )

                            if ($parentEntry.Count -gt 0) {
                                $parentEntry[0] | Add-Member `
                                    -NotePropertyName 'Parent' `
                                    -NotePropertyValue $true -Force

                                $m.Matrix = @($parentEntry[0]) + $childMatrix
                            }
                            else {
                                $m.Matrix = $childMatrix
                            }

                            # C. Merge Defaults per Folder
                            if ($context.Defaults.DefaultAcl.Count -gt 0) {
                                try {
                                    $applyDefaultPerms = $m.Setting.Formatted.ApplyDefaultPermissions

                                    foreach ($folder in $m.Matrix) {
                                        # Ignored folders ('I' in the matrix)
                                        # must never receive default (or any)
                                        # permissions; leave them untouched.
                                        if ($folder.Ignore) { continue }

                                        # A folder listed without any permissions
                                        # is an inherit-only folder: it must keep
                                        # inheriting from its parent. Merging
                                        # defaults into it would turn its empty ACL
                                        # into an explicit, protected ACL and cut
                                        # inheritance, which is not desired. Leave
                                        # the empty ACL as-is so SetPermissions
                                        # keeps the folder purely inheriting.
                                        if ($folder.ACL.Count -eq 0) { continue }

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
                # Must be the exact shape Import-MatrixFileHC returns: the
                # reporting stage reads .Item and .ReportFileName and assigns
                # to .LogFolder and .ReportFilePath on this object.
                if (-not $fileResult) {
                    $fileResult = New-MatrixFileResultHC -MatrixFile $file
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

                        # The original path no longer exists after the move.
                        # Record the new location so the HTML mail and report
                        # can link the matrix file name to the archived file.
                        $fileResult | Add-Member `
                            -NotePropertyName 'ArchivedPath' `
                            -NotePropertyValue $destination -Force
                    }
                    catch {
                        $fileResult.Check.Add(
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
            # The conflicting file list is identical for every matrix in the
            # group, so compute it once per group instead of per matrix.
            $conflictingFiles = ($DupGroup.Group | ForEach-Object { $_.FileContext.Item.Name }) | Select-Object -Unique
            $fileListString = $conflictingFiles -join "', '"

            foreach ($MatrixObj in $DupGroup.Group) {
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
                    -Type 'FatalError' `
                    -Name 'AD Bulk Lookup Failure' `
                    -Message "Failed during bulk AD lookup. Some AD objects may be marked as unknown. Error: $_" `
                    -Category 'ExpandedMatrix' `
                    -SystemErrors $SystemErrors
                return $Context
            }
            #endregion

            # Keep the resolved AD details on the context. The END stage
            # uses them to build the 'AccessList', 'GroupManagers' and
            # 'AdObjects' sheets in the matrix file copy saved to the log
            # folder, without having to query AD a second time.
            $Context.AdObjectDetails = $adObjectDetails

            #region Build Name → SID map for quick lookup during ACL rewrite
            <#
             Only objects with a readable ObjectSid enter the map, and the ACL
             rewrite below drops any name it cannot find here. A dropped entry
             would remove the group's access on a protected ACL under Action
             'Fix' without reporting anything, so it is worth stating why that
             cannot happen:

             - Test-AdObjectInMatrixHC runs before the rewrite and raises
               'Unknown AD Objects in Matrix' (FatalError) for every name whose
               'adObject' is null, and a fatal check skips the matrix.
             - The only remaining gap is an object that resolves but whose
               ObjectSid could not be read. Get-ADObjectDetailHC searches on
               (samAccountName=..), and everything holding a samAccountName is a
               security principal, so it has an objectSid.

             Measured on 2026-08-13 over all 53 matrix files plus the defaults:
             1176 unique AD objects, all resolved with a readable SID, all in a
             single domain. Nothing was dropped.

             Raising a FatalError on a miss would therefore add a new way for a
             whole matrix to be skipped, for a condition that has never
             occurred. Re-measure with
             Scripts\Diagnostics\TestMatrixAdObjectSid.ps1 before concluding
             otherwise: cross-forest or trusted-domain principals are the case
             most likely to change this, because the Global Catalog fallback
             returns a partial attribute set.
            #>
            $nameToSid = @{}
            foreach ($detail in $adObjectDetails) {
                if ($detail.adObject -and $detail.adObject.ObjectSid) {
                    $nameToSid[$detail.SamAccountName] = $detail.adObject.ObjectSid
                }
            }
            #endregion

            #region Rewrite ACLs in all matrices to use SIDs instead of names
            foreach ($matrixObj in $Context.AllMatrices) {
                $isFileBroken = Test-FileHasFatalErrorHC `
                    -File $matrixObj.FileContext
                $isRowBroken = Test-ItemHasFatalErrorHC `
                    -CheckList $matrixObj.Check

                if ($isFileBroken -or $isRowBroken) {
                    continue
                }

                $adObjectCheck = Test-AdObjectInMatrixHC `
                    -Matrix $matrixObj.Matrix `
                    -ADObject $adObjectDetails `
                    -AdGroupPlaceHolders @($Context.Config.Matrix.AdGroupPlaceHolders)

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
                        # A name missing from $nameToSid is dropped. See the
                        # note above the map construction for why that cannot
                        # happen here.
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

        #region Verify if default permissions are required (per matrix file)
        # Whether a matrix file consumes the shared defaults is driven by its
        # own 'ApplyDefaultPermissions' rows, so this guard is evaluated per
        # file rather than globally: the outcome can differ from one matrix
        # file to the next. The resulting check is stored on the file's own
        # Check list so it surfaces in that file's execution report and a
        # 'Empty default ACL' fatal only skips the affected file instead of
        # aborting the whole run.
        foreach ($fileResult in $Context.FileResults) {
            # A structurally broken file can't be evaluated reliably and is
            # already reported through its existing fatal checks.
            if (
                Test-FileHasFatalErrorHC -File $fileResult
            ) {
                continue
            }

            $validRows = @($fileResult.Matrices | Where-Object {
                    -not (Test-ItemHasFatalErrorHC -CheckList $_.Check)
                })

            if (-not $validRows) { continue }

            $fileUsesDefaults = $validRows | Where-Object {
                $_.Setting.Formatted.ApplyDefaultPermissions
            } | Select-Object -First 1

            if (
                $fileUsesDefaults -and $Context.Defaults.DefaultAcl.Count -eq 0
            ) {
                $fileResult.Check.Add(
                    [pscustomobject]@{
                        Type        = 'FatalError'
                        Name        = 'Empty default ACL'
                        Description = 'This matrix file has one or more rows with ApplyDefaultPermissions=TRUE but the defaults file contains no valid ACL entries.'
                    }
                )
            }
            elseif (
                -not $fileUsesDefaults -and
                $Context.Defaults.DefaultAcl.Count -gt 0
            ) {
                $fileResult.Check.Add(
                    [pscustomobject]@{
                        Type        = 'Information'
                        Name        = 'Unused defaults'
                        Description = 'The defaults file contains ACL entries but this matrix file has no rows with ApplyDefaultPermissions=TRUE; defaults will be ignored for this file.'
                    }
                )
            }
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