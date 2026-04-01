function Invoke-PermissionMatrix {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ConfigurationJsonFile,

        [Parameter(Mandatory)]
        [hashtable]$ScriptPath
    )

    Invoke-PermissionMatrixInternalHC `
        -ConfigurationJsonFile $ConfigurationJsonFile `
        -ScriptPath $ScriptPath `
        -Mode 'Full'
}