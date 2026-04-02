
function Invoke-PermissionMatrixBeginHC {
    <#
    .SYNOPSIS
        BEGIN stage for the Permission Matrix pipeline.

    .DESCRIPTION
        - Loads and validates the configuration JSON
        - Validates required top-level properties
        - Normalizes runtime settings
        - Builds and returns the execution context

        This function must NOT:
        - Read Excel files
        - Query AD
        - Scan matrix folders
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
        #region Configuration file existence
        if (-not (Test-Path -LiteralPath $ConfigurationJsonFile -PathType Leaf)) {
            Add-ErrorHC `
                -Type 'FatalError' `
                -Name 'Configuration file not found' `
                -Message "Configuration file '$ConfigurationJsonFile' does not exist." `
                -Category 'RuntimeSettings' `
                -SystemErrors $SystemErrors

            return $null
        }
        #endregion

        #region Load JSON
        try {
            $json = Get-Content `
                -LiteralPath $ConfigurationJsonFile `
                -Raw `
                -Encoding UTF8 |
            ConvertFrom-Json -Depth 50
        }
        catch {
            Add-ErrorHC `
                -Type 'FatalError' `
                -Name 'Invalid JSON' `
                -Message "Failed to parse configuration file '$ConfigurationJsonFile'." `
                -Description $_ `
                -Category 'RuntimeSettings' `
                -SystemErrors $SystemErrors

            return $null
        }
        #endregion

        #region Validate configuration
        Validate-ConfigurationStructureHC `
            -Json $json `
            -SystemErrors $SystemErrors

        if ($SystemErrors.Value.Count -gt 0) {
            return $null
        }
        #endregion

        return [pscustomobject]@{
            Settings      = $json.Settings
            Matrix        = $json.Matrix
            Export        = $json.Export
            ServiceNow    = $json.ServiceNow
            MaxConcurrent = $json.MaxConcurrent
            ScriptPath    = $ScriptPath
            StartTime     = Get-Date
            Counter       = New-CounterObjectHC
            ExportedFiles = @{}
        }
    }
    catch {
        Add-ErrorHC `
            -Type 'FatalError' `
            -Category 'Runtime' `
            -Name 'BEGIN stage failure' `
            -Message "Unexpected failure in BEGIN stage: $_" `
            -SystemErrors $SystemErrors

        return $null
    }
}