function Import-MatrixDefaultsFileHC {
    <#
    .SYNOPSIS
        Validates and imports the global matrix defaults Excel file.

    .DESCRIPTION
        Reads the specified Defaults Excel file to extract global fallback 
        permissions and notification email addresses. 
        
        The script strictly validates the file structure: it requires a 
        'Settings' worksheet containing 'MailTo', 'ADObjectName', and 
        'Permission' columns. If the file is missing, the worksheet is absent, 
        or the mandatory columns are missing, it safely catches the exception 
        and pushes a 'FatalError' to the global SystemErrors reference.

    .PARAMETER Matrix
        A custom object representing the 'Matrix' configuration node from the 
        JSON file. It must contain a 'DefaultsFile' property with the absolute 
        path to the defaults Excel file.

    .PARAMETER SystemErrors
        A reference variable ([ref]) containing a List[pscustomobject]. Used to 
        capture and bubble up terminating pipeline errors without crashing the 
        main orchestrator.

    .OUTPUTS
        System.Management.Automation.PSCustomObject. 
        Returns a custom object containing:
        - FilePath   : Absolute path to the loaded defaults file.
        - DefaultAcl : A hashtable mapping AD Objects to their default 
                       Permission characters.
        - MailTo     : A Generic List of email addresses extracted from the     
                       file.
        Returns $null if a FatalError occurs during import.

    .EXAMPLE
        $sysErrors = [System.Collections.Generic.List[pscustomobject]]::new()
        $matrixConfig = [pscustomobject]@{ 
            DefaultsFile = 'C:\Matrix\Defaults.xlsx' 
        }
        
        $defaults = Import-MatrixDefaultsFileHC `
            -Matrix $matrixConfig `
            -SystemErrors ([ref]$sysErrors)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Matrix,

        [Parameter(Mandatory)]
        [ref]$SystemErrors
    )

    try {
        #region Check if defaults file exists
        try {
            $defaultsItem = Get-Item -LiteralPath $Matrix.DefaultsFile -ErrorAction Stop
        }
        catch {
            Add-ErrorHC `
                -Type 'FatalError' `
                -Name 'Defaults file not found' `
                -Message "Defaults file '$($Matrix.DefaultsFile)' not found." `
                -Category 'Matrix' `
                -SystemErrors $SystemErrors
            return $null
        }
        #endregion

        #region Import defaults file and check for 'Settings' worksheet
        try {
            Write-Verbose "Import matrix defaults file '$($defaultsItem.FullName)'"

            $defaultsImport = Import-Excel `
                -Path $defaultsItem.FullName `
                -Sheet 'Settings' `
                -DataOnly `
                -ErrorAction Stop
        }
        catch {
            Add-ErrorHC `
                -Type 'FatalError' `
                -Name 'Defaults worksheet missing' `
                -Message "Worksheet 'Settings' not found in defaults file." `
                -Category 'Matrix' `
                -SystemErrors $SystemErrors
            return $null
        }
        #endregion

        #region Validate mandatory columns
        $columns = $defaultsImport[0].PSObject.Properties.Name
        foreach ($required in 'MailTo', 'ADObjectName', 'Permission') {
            if ($required -notin $columns) {
                Add-ErrorHC `
                    -Type 'FatalError' `
                    -Name 'Invalid defaults format' `
                    -Message "Mandatory column '$required' not found in defaults file." `
                    -Category 'Matrix' `
                    -SystemErrors $SystemErrors
                return $null
            }
        }
        #endregion

        #region Get default ACL
        $defaultAcl = Get-DefaultAclHC `
            -Sheet $defaultsImport `
            -SystemErrors $SystemErrors

        if (Test-ItemHasFatalErrorHC -CheckList $SystemErrors.Value) {
            return $null
        }
            
        $mailTo = [System.Collections.Generic.List[string]]::new()
        foreach ($row in $defaultsImport) {
            if (-not [string]::IsNullOrWhiteSpace($row.MailTo)) {
                $mailTo.Add($row.MailTo.ToString().Trim())
            }
        }

        if ($mailTo.Count -eq 0) {
            Add-ErrorHC `
                -Type 'FatalError' `
                -Name 'No MailTo addresses' `
                -Message 'No valid mail addresses found in defaults file.' `
                -Category 'Matrix' `
                -SystemErrors $SystemErrors
            return $null
        }

        # ------------------------------------------------------------
        # Return defaults object
        # ------------------------------------------------------------
        return [pscustomobject]@{
            FilePath   = $defaultsItem.FullName
            DefaultAcl = $defaultAcl
            MailTo     = $mailTo
        }
    }
    catch {
        Add-ErrorHC `
            -Type 'FatalError' `
            -Name 'Defaults import failed' `
            -Message $_ `
            -Category 'Matrix' `
            -SystemErrors $SystemErrors
        return $null
    }
}
