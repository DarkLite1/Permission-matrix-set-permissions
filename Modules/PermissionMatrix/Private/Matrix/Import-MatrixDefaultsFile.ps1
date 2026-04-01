function Import-MatrixDefaultsHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Matrix,

        [Parameter(Mandatory)]
        [ref]$SystemErrors
    )

    try {
        # ------------------------------------------------------------
        # Locate defaults file
        # ------------------------------------------------------------
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

        # ------------------------------------------------------------
        # Import Settings worksheet
        # ------------------------------------------------------------
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

        # ------------------------------------------------------------
        # Validate mandatory columns
        # ------------------------------------------------------------
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

        # ------------------------------------------------------------
        # Extract default ACL
        # ------------------------------------------------------------
        $defaultAcl = Get-DefaultAclHC -Sheet $defaultsImport

        # ------------------------------------------------------------
        # Extract MailTo
        # ------------------------------------------------------------
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
