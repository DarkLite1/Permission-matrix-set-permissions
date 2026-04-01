
function Invoke-PermissionMatrixBeginHC {
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
        # ------------------------------------------------------------
        # Load JSON
        # ------------------------------------------------------------
        if (-not (Test-Path $ConfigurationJsonFile)) {
            Add-ErrorHC `
                -Type 'FatalError' `
                -Category 'Configuration' `
                -Name 'No configuration file' `
                -Message "Configuration file '$ConfigurationJsonFile' not found." `
                -SystemErrors $SystemErrors
            
            return $null
        }

        $json = Get-Content $ConfigurationJsonFile -Raw -Encoding UTF8 |
        ConvertFrom-Json -Depth 50

        # ------------------------------------------------------------
        # Validate configuration
        # ------------------------------------------------------------
        Validate-ConfigurationStructureHC `
            -Json $json `
            -SystemErrors $SystemErrors

        if (Test-HasFatalErrorsHC $SystemErrors) {
            return $null
        }

        # ------------------------------------------------------------
        # Build context object
        # ------------------------------------------------------------
        return [pscustomobject]@{
            Settings      = $json.Settings
            Matrix        = $json.Matrix
            Export        = $json.Export
            ServiceNow    = $json.ServiceNow
            MaxConcurrent = $json.MaxConcurrent
            ScriptPath    = $ScriptPath
            StartTime     = Get-Date
            Counter       = New-CounterObjectHC
        }
    }
    catch {
        Add-ErrorHC `
            -Type 'FatalError' `
            -Category 'Runtime' `
            -Name 'BEGIN stage failure' `
            -Message "Unhandled exception occurred: $_" `
            -SystemErrors $SystemErrors

        return $null
    }
}