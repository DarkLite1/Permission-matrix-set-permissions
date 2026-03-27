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
    if (Test-HasFatalErrorsHC ([ref]$systemErrors)) {
        Write-Warning 'Skipping PROCESS stage due to fatal BEGIN errors...'
        $importedMatrix = $null
    }
    else {
        $importedMatrix = Invoke-PermissionMatrixProcess `
            -Context $context `
            -SystemErrors ([ref]$systemErrors)
    }

    # END stage
    Invoke-PermissionMatrixEnd `
        -Context $context `
        -ImportedMatrix $importedMatrix `
        -SystemErrors ([ref]$systemErrors)

    
    # ------------------------------------------------------------
    # EXIT CODE HANDLING
    # ------------------------------------------------------------
    if ($systemErrors.Type -contains 'FatalError') {
        exit 1
    }
}