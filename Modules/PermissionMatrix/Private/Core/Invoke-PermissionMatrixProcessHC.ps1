function Invoke-PermissionMatrixProcessHC {
    <#
    .SYNOPSIS
        PROCESS stage for the Permission Matrix pipeline.

    .DESCRIPTION
        - Discovers matrix Excel files
        - Loads defaults only when needed
        - Imports and validates matrix files in parallel
        - Returns structured execution units and checks

        This function performs NO:
        - AD queries
        - Permission execution
        - Export / logging / mail
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Context,

        [Parameter(Mandatory)]
        [ref]$SystemErrors
    )

    try {
        #region Discover matrix Excel files (serial, cheap)
        try {
            $matrixFiles = Get-ChildItem `
                -Path $Context.Matrix.FolderPath `
                -Filter '*.xlsx' `
                -File `
                -ErrorAction Stop
        }
        catch {
            Add-ErrorHC `
                -Type 'FatalError' `
                -Name 'Matrix folder access failed' `
                -Message "Failed to access matrix folder '$($Context.Matrix.FolderPath)'." `
                -Description $_ `
                -Category 'Matrix' `
                -SystemErrors $SystemErrors

            return [pscustomobject]@{
                FoundMatrices = $false
                Imported      = @()
            }
        }
        #endregion

        #region No files, exit early with success (but no matrices)
        if (-not $matrixFiles -or $matrixFiles.Count -eq 0) {
            return [pscustomobject]@{
                FoundMatrices = $false
                Imported      = @()
            }
        }
        #endregion

        #region Import defaults ONCE (only now)
        $defaults = Import-MatrixDefaultsHC `
            -Matrix $Context.Matrix `
            -SystemErrors $SystemErrors

        if (-not $defaults -or (Test-HasFatalErrorsHC $SystemErrors)) {
            return [pscustomobject]@{
                FoundMatrices = $true
                Imported      = @()
            }
        }
        #endregion

        #region Exclude defaults file itself from matrix list
        $matrixFiles = $matrixFiles |
        Where-Object { $_.FullName -ne $defaults.FilePath }

        if (-not $matrixFiles -or $matrixFiles.Count -eq 0) {
            return [pscustomobject]@{
                FoundMatrices = $true
                Imported      = @()
            }
        }
        #endregion

        #region Parallel import + validation per matrix file
        $throttle = $Context.MaxConcurrent.FoldersPerMatrix ?? 4

        $parallelResults = $matrixFiles | Sort-Object Name |
        ForEach-Object -Parallel {
            param($file, $context)

            #region Import module
            Import-Module `
                -FullyQualifiedName $context.ScriptPath.PermissionMatrixModule `
                -Force `
                -ErrorAction Stop
            #endregion

            Import-MatrixFileHC `
                -MatrixFile $file `
                -Context $context

        } -ThrottleLimit $throttle -ArgumentList $Context
        #endregion

        #region Merge results on main thread
        $allMatrices = @()

        foreach ($fileResult in $parallelResults) {
            # Promote file-level checks to SystemErrors
            # foreach ($check in $fileResult.File.Check) {
            #     Add-ErrorHC `
            #         -Type $check.Type `
            #         -Name $check.Name `
            #         -Message $check.Description `
            #         -Category 'Matrix' `
            #         -SystemErrors $SystemErrors
            # }

            # Collect execution matrices
            if ($fileResult.Matrices) {
                $allMatrices += $fileResult.Matrices
            }
        }
        #endregion

        #region Archive processed matrix files
        if ($Context.Matrix.Archive) {

            $archiveRootReady = $false
            $archivePath = Join-Path `
                -Path $Context.Matrix.FolderPath `
                -ChildPath 'Archive'

            try {
                if (-not (Test-Path -LiteralPath $archivePath -PathType Container)) {
                    $null = New-Item -ItemType Directory -Path $archivePath -ErrorAction Stop
                }
                $archiveRootReady = $true
            }
            catch {
                Add-ErrorHC `
                    -Type 'Warning' `
                    -Name 'Archive folder creation failed' `
                    -Message "Failed to create archive folder '$archivePath'." `
                    -Description $_ `
                    -Category 'Matrix' `
                    -SystemErrors $SystemErrors
            }

            if ($archiveRootReady) {
                foreach ($file in $matrixFiles) {
                    try {
                        $destination = Join-Path `
                            -Path $archivePath `
                            -ChildPath $file.Name

                        Move-Item `
                            -LiteralPath $file.FullName `
                            -Destination $destination `
                            -Force `
                            -ErrorAction Stop
                    }
                    catch {
                        Add-ErrorHC `
                            -Type 'Warning' `
                            -Name 'Archive failed' `
                            -Message "Failed to archive matrix file '$($file.Name)'." `
                            -Description $_ `
                            -Category 'Matrix' `
                            -SystemErrors $SystemErrors
                    }
                }
            }
        }
        #endregion

        return [pscustomobject]@{
            FoundMatrices = $true
            Imported      = $allMatrices
        }
    }
    catch {
        Add-ErrorHC `
            -Type 'FatalError' `
            -Category 'Runtime' `
            -Name 'PROCESS stage failure' `
            -Message "Unhandled exception occurred: $_" `
            -SystemErrors $SystemErrors

        return [pscustomobject]@{
            FoundMatrices = $true
            Imported      = @()
        }
    }
}