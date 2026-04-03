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
            Add-ErrorHC -Type 'FatalError' -Name 'Configuration file not found' -Message "File '$ConfigurationJsonFile' does not exist." -Category 'RuntimeSettings' -SystemErrors $SystemErrors
            return $null
        }

        $json = Get-Content -LiteralPath $ConfigurationJsonFile -Raw -Encoding UTF8 | ConvertFrom-Json -Depth 50
        
        Validate-ConfigurationStructureHC -Json $json -SystemErrors $SystemErrors
        if ($SystemErrors.Value.Count -gt 0) { return $null }

        $Context = [pscustomobject]@{
            Settings      = $json.Settings
            Matrix        = $json.Matrix
            Export        = $json.Export
            ServiceNow    = $json.ServiceNow
            MaxConcurrent = $json.MaxConcurrent
            ScriptPath    = $ScriptPath
            StartTime     = Get-Date
            Counter       = New-CounterObjectHC
            ExportedFiles = @{}
            FoundMatrices = $false
            Matrices      = @()
            Defaults      = $null
        }

        # =====================================================================
        # 2. PARALLEL: Read, Validate, and Archive Matrix Files
        # =====================================================================
        try {
            $matrixFiles = Get-ChildItem -Path $Context.Matrix.FolderPath -Filter '*.xlsx' -File -ErrorAction Stop
        }
        catch {
            Add-ErrorHC -Type 'FatalError' -Name 'Matrix folder access failed' -Message "Cannot access '$($Context.Matrix.FolderPath)'." -Category 'Matrix' -SystemErrors $SystemErrors
            return $Context
        }

        # Exclude Defaults file from the processing list
        $matrixFiles = $matrixFiles | Where-Object { $_.FullName -ne $Context.Matrix.DefaultsFile }

        if (-not $matrixFiles -or $matrixFiles.Count -eq 0) {
            return $Context # No files found, exit BEGIN gracefully
        }
        
        $Context.FoundMatrices = $true

        # Setup Archive Folder
        $archivePath = $null
        if ($Context.Matrix.Archive) {
            $archivePath = Join-Path -Path $Context.Matrix.FolderPath -ChildPath 'Archive'
            if (-not (Test-Path -LiteralPath $archivePath -PathType Container)) {
                $null = New-Item -ItemType Directory -Path $archivePath -Force -ErrorAction SilentlyContinue
            }
        }

        $throttle = $Context.MaxConcurrent.FoldersPerMatrix ?? 4

        # Read and Archive in Parallel
        $parallelResults = Invoke-WithOptionalParallelismHC -InputObject $matrixFiles -ThrottleLimit $throttle -ArgumentList $Context, $archivePath -ScriptBlock {
            param($file, $context, $archiveFolder)

            Import-Module $context.ScriptPath.PermissionMatrixModule -Force -ErrorAction Stop
            
            # Read and validate the matrix
            $fileResult = Import-MatrixFileHC -MatrixFile $file -Context $context

            # Quarantine/Archive immediately to prevent error loops on next schedule
            if ($archiveFolder) {
                try {
                    $destination = Join-Path -Path $archiveFolder -ChildPath $file.Name
                    Move-Item -LiteralPath $file.FullName -Destination $destination -Force -ErrorAction Stop
                }
                catch {
                    $fileResult.File.Check.Add([pscustomobject]@{
                            Type = 'Warning'; Name = 'Archiving failed'; Description = 'File could not be moved to archive.'; Value = $_
                        })
                }
            }
            return $fileResult
        }

        $importedMatrices = [System.Collections.Generic.List[pscustomobject]]::new()
        foreach ($res in $parallelResults) {
            if ($res.Matrices) { $importedMatrices.AddRange($res.Matrices) }
        }
        $Context.Matrices = $importedMatrices

        # =====================================================================
        # 3. SEQUENTIAL: Cross-Matrix Checks & AD Lookups
        # =====================================================================
        
        # 3a. Duplicate ComputerName/Path Validation
        $duplicateSettings = $Context.Matrices.Settings | Group-Object -Property { $_.Import.ComputerName }, { $_.Import.Path } | Where-Object Count -GE 2
        foreach ($DupGroup in $duplicateSettings) {
            foreach ($Setting in $DupGroup.Group) {
                $Setting.Check += [PSCustomObject]@{
                    Type        = 'FatalError'
                    Name        = 'Duplicate ComputerName/Path combination'
                    Description = 'The combination must be unique across all active matrix files.'
                    Value       = "Computer: $($Setting.Import.ComputerName), Path: $($Setting.Import.Path)"
                }
            }
        }

        # 3b. Read Defaults Excel file and validate
        $defaults = Import-MatrixDefaultsHC -Matrix $Context.Matrix -SystemErrors $SystemErrors
        if (Test-HasFatalErrorsHC $SystemErrors) { return $Context }
        $Context.Defaults = $defaults

        # 3c. One AD query for all objects combined
        $allAdObjects = $Context.Matrices.Settings.Matrix.ACL.Keys | Sort-Object -Unique

        if ($allAdObjects.Count -gt 0) {
            $adObjectDetails = @(Get-ADObjectDetailHC -ADObjectName $allAdObjects -Type 'SamAccountName')
            
            # 3d. Combine AD info with matrix data (Expanded Matrix Validation)
            foreach ($matrixObj in $Context.Matrices) {
                foreach ($S in $matrixObj.Settings) {
                    if (-not $S.Matrix) { continue }

                    $expandedCheck = Test-ExpandedMatrixHC `
                        -Matrix $S.Matrix `
                        -ADObject $adObjectDetails `
                        -DefaultAcl $Context.Defaults.DefaultAcl `
                        -ExcludedSamAccountName $Context.Matrix.ExcludedSamAccountName

                    if ($expandedCheck) {
                        $S.Check += $expandedCheck | ConvertTo-StructuredObjectHC
                    }
                }
            }
        }

        return $Context
    }
    catch {
        Add-ErrorHC -Type 'FatalError' -Category 'Runtime' -Name 'BEGIN stage failure' -Message "Unhandled exception: $_" -SystemErrors $SystemErrors
        return $null
    }
}