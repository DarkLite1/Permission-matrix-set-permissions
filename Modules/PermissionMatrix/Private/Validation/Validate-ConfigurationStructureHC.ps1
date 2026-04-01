function Validate-ConfigurationStructureHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Json,
        [Parameter(Mandatory)][ref]   $SystemErrors
    )

    #
    # 1. Required Top-Level Properties
    #
    foreach ($prop in @(
            'Matrix', 'Export', 'ServiceNow', 'MaxConcurrent', 'PSSessionConfiguration', 'Settings'
        )) {
        if ($null -eq $Json.$prop) {
            Add-JsonSchemaErrorHC `
                -Type 'FatalError' `
                -Name "Missing '$prop'" `
                -Message "Property '$prop' not found in JSON." `
                -SystemErrors $SystemErrors
        }
    }

    #
    # 2. Settings block
    #
    if ($Json.Settings) {

        # SaveLogFiles
        if (-not $Json.Settings.SaveLogFiles.Where.Folder) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Missing 'Settings.SaveLogFiles.Where.Folder'" `
                -Message 'Folder is required.' `
                -SystemErrors $SystemErrors
        }

        if ($null -eq $Json.Settings.SaveLogFiles.Detailed) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Missing 'Settings.SaveLogFiles.Detailed'" `
                -Message 'Detailed is required.' `
                -SystemErrors $SystemErrors
        }
        elseif ($Json.Settings.SaveLogFiles.Detailed -isnot [bool]) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Incorrect 'Settings.SaveLogFiles.Detailed'" `
                -Message 'Must be boolean.' `
                -SystemErrors $SystemErrors
        }

        # SendMail
        if (-not $Json.Settings.SendMail) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Missing 'Settings.SendMail'" `
                -Message 'SendMail block is mandatory.' `
                -SystemErrors $SystemErrors
        }
        else {
            if (-not $Json.Settings.SendMail.From) {
                Add-JsonSchemaErrorHC -Type 'FatalError' `
                    -Name "Missing 'Settings.SendMail.From'" `
                    -Message 'From is required.' `
                    -SystemErrors $SystemErrors
            }
            if ($Json.Settings.SendMail.To -and
                ($Json.Settings.SendMail.To -isnot [string] -and
                $Json.Settings.SendMail.To -isnot [array])) {
                Add-JsonSchemaErrorHC -Type 'FatalError' `
                    -Name "Incorrect 'Settings.SendMail.To'" `
                    -Message 'Must be string or array.' `
                    -SystemErrors $SystemErrors
            }
            if ($null -eq $Json.Settings.SendMail.Body) {
                Add-JsonSchemaErrorHC -Type 'FatalError' `
                    -Name "Missing 'Settings.SendMail.Body'" `
                    -Message 'Body is required.' `
                    -SystemErrors $SystemErrors
            }
        }
    }


    #
    # 3. Matrix
    #
    if ($Json.Matrix) {

        if (-not $Json.Matrix.FolderPath) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Missing 'Matrix.FolderPath'" `
                -Message 'FolderPath is required.' `
                -SystemErrors $SystemErrors
        }

        if (-not $Json.Matrix.DefaultsFile) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Missing 'Matrix.DefaultsFile'" `
                -Message 'DefaultsFile is required.' `
                -SystemErrors $SystemErrors
        }

        if ($Json.Matrix.ExcludedSamAccountName -and
            $Json.Matrix.ExcludedSamAccountName -isnot [array]) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Incorrect 'Matrix.ExcludedSamAccountName'" `
                -Message 'Must be an array.' `
                -SystemErrors $SystemErrors
        }

        if ($null -eq $Json.Matrix.Archive -or $Json.Matrix.Archive -isnot [bool]) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Incorrect 'Matrix.Archive'" `
                -Message 'Must be boolean.' `
                -SystemErrors $SystemErrors
        }
    }


    #
    # 4. MaxConcurrent
    #
    if ($Json.MaxConcurrent) {
        foreach ($prop in 'Computers', 'FoldersPerMatrix', 'JobsPerRemoteComputer') {
            $val = $Json.MaxConcurrent.$prop
            if ($null -eq $val -or $val -notmatch '^\d+$') {
                Add-JsonSchemaErrorHC -Type 'FatalError' `
                    -Name "Incorrect 'MaxConcurrent.$prop'" `
                    -Message "$prop must be numeric." `
                    -SystemErrors $SystemErrors
            }
        }
    }


    #
    # 5. Export
    #
    if (-not $Json.Export) {
        Add-JsonSchemaErrorHC -Type 'FatalError' `
            -Name "Missing 'Export'" `
            -Message 'Export section missing.' `
            -SystemErrors $SystemErrors
    }
}

