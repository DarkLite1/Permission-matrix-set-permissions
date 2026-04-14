function Validate-ConfigurationStructureHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Json,
        [Parameter(Mandatory)][ref]   $SystemErrors
    )

    #region Top-Level properties
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
    #endregion

    #region Settings
    if ($Json.Settings) {
        #region SaveInEventLog
        if ($Json.Settings.SaveLogFiles.Detailed -isnot [bool]) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Incorrect 'Settings.SaveLogFiles.Detailed'" `
                -Message 'Must be boolean.' `
                -SystemErrors $SystemErrors
        }

        if ($Json.Settings.SaveInEventLog.Save -isnot [bool]) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Incorrect 'Settings.SaveInEventLog.Save'" `
                -Message 'Must be boolean.' `
                -SystemErrors $SystemErrors
        }
        #endregion

        #region SaveLogFiles
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
        #endregion

        #region SendMail
        if ( $Json.Settings.SendMail) {
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
            if (-not $Json.Settings.SendMail.Smtp.Port -or $Json.Settings.SendMail.Smtp.Port -notmatch '^\d+$') {
                Add-JsonSchemaErrorHC -Type 'FatalError' -Name "Incorrect 'SendMail.Smtp.Port'" `
                    -Message 'Port must be numeric.' `
                    -SystemErrors $SystemErrors
            }

            $validConn = @('None', 'Auto', 'SslOnConnect', 'StartTls', 'StartTlsWhenAvailable')
            if ($Json.Settings.SendMail.Smtp.ConnectionType -notin $validConn) {
                Add-JsonSchemaErrorHC -Type 'FatalError' -Name "Incorrect 'Settings.SendMail.Smtp.ConnectionType'" `
                    -Message 'Invalid connection type.' `
                    -SystemErrors $SystemErrors
            }
        }
        else {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Missing 'Settings.SendMail'" `
                -Message 'SendMail block is mandatory.' `
                -SystemErrors $SystemErrors
            
        }
        #endregion

        if (-not $json.Settings.ScriptName) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Missing 'Settings.ScriptName'" `
                -Message 'ScriptName is required.' `
                -SystemErrors $SystemErrors
        }
    }
    #endregion

    #region Matrix
    if ($Json.Matrix) {
        if (-not $Json.Matrix.FolderPath) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Missing 'Matrix.FolderPath'" `
                -Message 'FolderPath is required.' `
                -SystemErrors $SystemErrors
        }
        elseif (-not (Test-Path -LiteralPath $Json.Matrix.FolderPath -PathType Leaf)) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Incorrect 'Matrix.FolderPath'" `
                -Message "FolderPath '$($Matrix.FolderPath)' not found" `
                -SystemErrors $SystemErrors
        }

        if (-not $Json.Matrix.DefaultsFile) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Missing 'Matrix.DefaultsFile'" `
                -Message 'DefaultsFile is required.' `
                -SystemErrors $SystemErrors
        }
        elseif (-not (Test-Path -LiteralPath $Json.Matrix.DefaultsFile -PathType Leaf)) {
            Add-JsonSchemaErrorHC -Type 'FatalError' `
                -Name "Incorrect 'Matrix.DefaultsFile'" `
                -Message "DefaultsFile '$($Matrix.DefaultsFile)' not found" `
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
    #endregion

    #region MaxConcurrent
    if ($Json.MaxConcurrent) {
        foreach ($prop in 'Computers', 'FoldersPerMatrix', 'JobsPerRemoteComputer') {
            $val = $Json.MaxConcurrent.$prop
            if ($null -eq $val -or $val -notmatch '^\d+$') {
                Add-JsonSchemaErrorHC -Type 'FatalError' `
                    -Name "Incorrect 'MaxConcurrent.$prop'" `
                    -Message "Property 'MaxConcurrent.$prop' must be numeric." `
                    -SystemErrors $SystemErrors
            }
        }
    }
    #endregion

    #region Export
    if ($Json.Export) {
        if ($Json.Export.PermissionsExcelFile -and $Json.Export.PermissionsExcelFile -notmatch '\.xlsx$') {
            Add-JsonSchemaErrorHC -Type 'FatalError' -Name "Incorrect 'Export.PermissionsExcelFile'" `
                -Message 'Must end with .xlsx' `
                -SystemErrors $SystemErrors
        }

        if ($Json.Export.OverviewHtmlFile -and $Json.Export.OverviewHtmlFile -notmatch '\.html?$') {
            Add-JsonSchemaErrorHC -Type 'FatalError' -Name "Incorrect 'Export.OverviewHtmlFile'" `
                -Message 'Must end with .html' `
                -SystemErrors $SystemErrors
        }

        if ($Json.Export.ServiceNowFormDataExcelFile) {

            if ($Json.Export.ServiceNowFormDataExcelFile -notmatch '\.xlsx$') {
                Add-JsonSchemaErrorHC -Type 'FatalError' `
                    -Name "Incorrect 'Export.ServiceNowFormDataExcelFile'" `
                    -Message 'Must end with .xlsx' `
                    -SystemErrors $SystemErrors
            }

            if (-not $ServiceNow) {
                Add-JsonSchemaErrorHC -Type 'FatalError' `
                    -Name 'Incorrect configuration' `
                    -Message 'ServiceNow must be defined when using ServiceNowFormDataExcelFile.' `
                    -SystemErrors $SystemErrors
            }
            else {
                foreach ($p in 'CredentialsFilePath', 'TableName', 'Environment') {
                    if (-not $ServiceNow.$p) {
                        Add-JsonSchemaErrorHC -Type 'FatalError' `
                            -Name "Missing 'ServiceNow.$p'" `
                            -Message "$p is required." `
                            -SystemErrors $SystemErrors
                    }
                }
            }
        }
    }
    #endregion
}

