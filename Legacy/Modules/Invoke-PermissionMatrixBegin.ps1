function Invoke-PermissionMatrixBegin {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ConfigurationJsonFile,
        [Parameter(Mandatory)][hashtable]$ScriptPath,
        [Parameter(Mandatory)][ref]$SystemErrors
    )

    $context = [ordered]@{}
    $context.ScriptStartTime = Get-Date

    try {
        # Load JSON
        $json = Get-Content -LiteralPath $ConfigurationJsonFile -ErrorAction Stop | ConvertFrom-Json -Depth 20

        $context.Settings = $json.Settings
        $context.Matrix = $json.Matrix
        $context.Export = $json.Export
        $context.ServiceNow = $json.ServiceNow
        $context.MaxConcurrent = $json.MaxConcurrent
        $context.ScriptPathItem = $ScriptPath
        $context.json = $json

        # Convert settings to safe defaults
        $context.Settings = Ensure-SafeSettingsHC $context.Settings

        # Log folder
        $context.LogFolder = Ensure-LogFolderHC `
            -RequestedFolder $context.Settings.SaveLogFiles.Where.Folder `
            -SystemErrors $SystemErrors

        # Validate JSON schema
        Validate-ConfigurationStructure `
            -Json $context.json `
            -SystemErrors $SystemErrors
    }
    catch {
        $SystemErrors.Value.Add([pscustomobject]@{
                DateTime = Get-Date
                Message  = "BEGIN failed: $_"
            })
    }

    return $context
}