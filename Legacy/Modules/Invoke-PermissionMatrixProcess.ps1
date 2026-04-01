function Invoke-PermissionMatrixProcess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Context,
        [Parameter(Mandatory)][ref]$SystemErrors
    )

    try {
        $matrixFiles =
        Get-ChildItem -LiteralPath $Context.Matrix.FolderPath |
        Where-Object { -not $_.PSIsContainer -and $_.Extension -match '\.xlsx$' }

        if (-not $matrixFiles) { return $null }

        $importedMatrix = foreach ($file in $matrixFiles) {
            try {
                Import-MatrixFileHC `
                    -File $file `
                    -Settings $Context.Settings `
                    -MatrixConfig $Context.Matrix `
                    -SystemErrors $SystemErrors
            }
            catch {
                $SystemErrors.Value.Add([pscustomobject]@{
                        DateTime = Get-Date
                        Message  = "Matrix import failed for $($file.FullName): $_"
                    })
            }
        }

        return $importedMatrix
    }
    catch {
        $SystemErrors.Value.Add([pscustomobject]@{
                DateTime = Get-Date
                Message  = "PROCESS failed: $_"
            })
        return $null
    }
}