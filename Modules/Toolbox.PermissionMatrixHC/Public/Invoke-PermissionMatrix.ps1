function Invoke-PermissionMatrix {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ConfigurationJsonFile,
        [Parameter(Mandatory)][hashtable]$ScriptPath
    )

    $systemErrors = [System.Collections.Generic.List[object]]::new()

    # BEGIN stage
    $context = Invoke-PermissionMatrixBegin `
        -ConfigurationJsonFile $ConfigurationJsonFile `
        -ScriptPath $ScriptPath `
        -SystemErrors ([ref]$systemErrors)

    # PROCESS stage
    $importedMatrix = Invoke-PermissionMatrixProcess `
        -Context $context `
        -SystemErrors ([ref]$systemErrors)

    # END stage
    Invoke-PermissionMatrixEnd `
        -Context $context `
        -ImportedMatrix $importedMatrix `
        -SystemErrors ([ref]$systemErrors)
}