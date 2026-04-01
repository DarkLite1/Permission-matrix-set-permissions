function Remove-FileHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$FilePath,
        [ref]$SystemErrors
    )

    try {
        if (-not (Test-Path -LiteralPath $FilePath -PathType Leaf)) { return }
        Remove-Item -LiteralPath $FilePath -Force -ErrorAction Stop
    }
    catch {
        if ($SystemErrors) {
            Add-ErrorByCategoryHC `
                -Type 'Warning' `
                -Name 'RemoveFileFailed' `
                -Message "Failed to remove '$FilePath': $_" `
                -Category 'Logging' `
                -SystemErrors ([ref]$SystemErrors)
        }
        else {
            Write-Warning "Failed removing '$FilePath': $_"
        }
    }
}