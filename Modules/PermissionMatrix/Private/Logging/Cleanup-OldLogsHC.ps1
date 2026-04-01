function Cleanup-OldLogsHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$LogFolder,
        [int]$RetentionDays,
        [Parameter(Mandatory)] [ref]$SystemErrors
    )

    # Disabled or folder missing → nothing to do
    if ($RetentionDays -le 0 -or -not $LogFolder) { return }

    try {
        if (-not (Test-Path -LiteralPath $LogFolder -PathType Container)) { return }

        $cutoff = (Get-Date).AddDays(-$RetentionDays)

        # --- 1. Delete old files ---
        Get-ChildItem -LiteralPath $LogFolder -Recurse -File -ErrorAction Stop |
        Where-Object { $_.CreationTime -lt $cutoff } |
        ForEach-Object {
            try {
                Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
            }
            catch {
                Add-ErrorByCategoryHC `
                    -Type 'Warning' `
                    -Name 'LogCleanupFailedFile' `
                    -Message "Failed to delete log file '$($_.FullName)': $_" `
                    -Category 'Logging' `
                    -SystemErrors ([ref]$SystemErrors)
            }
        }

        # --- 2. Empty folder cleanup (bottom-up) ---
        Get-ChildItem -LiteralPath $LogFolder -Recurse -Directory -ErrorAction Stop |
        Sort-Object FullName -Descending |
        ForEach-Object {
            if (-not $_.GetFileSystemInfos().Count) {
                try {
                    Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
                }
                catch {
                    Add-ErrorByCategoryHC `
                        -Type 'Warning' `
                        -Name 'LogCleanupFailedDirectory' `
                        -Message "Failed to remove empty folder '$($_.FullName)': $_" `
                        -Category 'Logging' `
                        -SystemErrors ([ref]$SystemErrors)
                }
            }
        }
    }
    catch {
        Add-ErrorByCategoryHC `
            -Type 'Warning' `
            -Name 'LogCleanupFailedGeneral' `
            -Message "General log cleanup failure: $_" `
            -Category 'Logging' `
            -SystemErrors ([ref]$SystemErrors)
    }
}