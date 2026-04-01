function Invoke-PermissionMatrixEndHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Context,

        [Parameter(Mandatory)]
        [array]$ImportedMatrix,

        [Parameter(Mandatory)]
        [ref]$SystemErrors
    )

    try {
        # ------------------------------------------------------------
        # Export
        # ------------------------------------------------------------
        if ($Context.Export) {
            $Context.ExportedFiles = Export-FilesHC `
                -ImportedMatrix $ImportedMatrix `
                -Context $Context `
                -SystemErrors $SystemErrors
        }

        # ------------------------------------------------------------
        # HTML / Mail / Logging
        # ------------------------------------------------------------
        Write-SystemLogsHC `
            -Context $Context `
            -ImportedMatrix $ImportedMatrix `
            -SystemErrors $SystemErrors
    }
    catch {
        Add-ErrorHC `
            -Type 'FatalError' `
            -Category 'Runtime' `
            -Name 'END stage failure' `
            -Message "Unhandled exception occurred: $_" `
            -SystemErrors $SystemErrors
    }
}
