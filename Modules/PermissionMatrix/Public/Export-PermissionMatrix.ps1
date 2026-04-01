function Export-PermissionMatrix {
    <#
    .SYNOPSIS
        Runs the Permission Matrix export pipeline only.
    #>

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
        -Mode 'ExportOnly'
}