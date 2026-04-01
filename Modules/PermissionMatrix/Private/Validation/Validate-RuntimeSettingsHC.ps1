function Validate-RuntimeSettingsHC {
    [CmdletBinding()]
    param(
        [object]$Settings,
        [object]$Matrix,
        [object]$Export,
        [object]$ServiceNow,
        [object]$MaxConcurrent,
        [ref]$SystemErrors
    )

    #
    # 1. Settings
    #
    if (-not $Settings) {
        Add-RuntimeErrorHC -Type 'FatalError' -Name "Missing 'Settings'" `
            -Message 'Settings are required.' `
            -SystemErrors $SystemErrors
        return
    }

    if (-not $Settings.ScriptName) {
        Add-RuntimeErrorHC -Type 'Warning' -Name "Missing 'Settings.ScriptName'" `
            -Message 'Using default script name.' `
            -SystemErrors $SystemErrors

        $Settings |
        Add-Member -NotePropertyName ScriptName -NotePropertyValue 'Default script name' -Force
    }

    if ($Settings.SaveLogFiles.Detailed -isnot [bool]) {
        Add-RuntimeErrorHC -Type 'FatalError' -Name "Incorrect 'Settings.SaveLogFiles.Detailed'" `
            -Message 'Must be boolean.' `
            -SystemErrors $SystemErrors
    }

    if ($Settings.SaveInEventLog.Save -isnot [bool]) {
        Add-RuntimeErrorHC -Type 'FatalError' -Name "Incorrect 'Settings.SaveInEventLog.Save'" `
            -Message 'Must be boolean.' `
            -SystemErrors $SystemErrors
    }

    if (-not $Settings.SendMail.From) {
        Add-RuntimeErrorHC -Type 'FatalError' -Name "Missing 'Settings.SendMail.From'" `
            -Message 'From is required.' `
            -SystemErrors $SystemErrors
    }

    if (-not $Settings.SendMail.To) {
        Add-RuntimeErrorHC -Type 'FatalError' -Name "Missing 'Settings.SendMail.To'" `
            -Message 'To is required.' `
            -SystemErrors $SystemErrors
    }

    if (-not $Settings.SendMail.Body) {
        Add-RuntimeErrorHC -Type 'FatalError' -Name "Missing 'Settings.SendMail.Body'" `
            -Message 'Body is required.' `
            -SystemErrors $SystemErrors
    }

    if (-not $Settings.SendMail.Smtp.Port -or $Settings.SendMail.Smtp.Port -notmatch '^\d+$') {
        Add-RuntimeErrorHC -Type 'FatalError' -Name "Incorrect 'SendMail.Smtp.Port'" `
            -Message 'Port must be numeric.' `
            -SystemErrors $SystemErrors
    }

    $validConn = @('None', 'Auto', 'SslOnConnect', 'StartTls', 'StartTlsWhenAvailable')
    if ($Settings.SendMail.Smtp.ConnectionType -notin $validConn) {
        Add-RuntimeErrorHC -Type 'FatalError' -Name "Incorrect 'Settings.SendMail.Smtp.ConnectionType'" `
            -Message 'Invalid connection type.' `
            -SystemErrors $SystemErrors
    }


    #
    # 2. Matrix
    #
    if (-not $Matrix) {
        Add-RuntimeErrorHC -Type 'FatalError' -Name "Missing 'Matrix'" `
            -Message 'Matrix block is required.' `
            -SystemErrors $SystemErrors
    }
    else {
        if (-not (Test-Path -LiteralPath $Matrix.DefaultsFile -PathType Leaf)) {
            Add-RuntimeErrorHC -Type 'FatalError' -Name "Incorrect 'Matrix.DefaultsFile'" `
                -Message "DefaultsFile not found: $($Matrix.DefaultsFile)" `
                -SystemErrors $SystemErrors
        }
        if (-not $Matrix.FolderPath) {
            Add-RuntimeErrorHC -Type 'FatalError' -Name "Missing 'Matrix.FolderPath'" `
                -Message 'FolderPath missing.' `
                -SystemErrors $SystemErrors
        }
        if ($Matrix.ExcludedSamAccountName -isnot [array]) {
            Add-RuntimeErrorHC -Type 'FatalError' -Name "Incorrect 'Matrix.ExcludedSamAccountName'" `
                -Message 'Must be array.' `
                -SystemErrors $SystemErrors
        }
    }


    #
    # 3. MaxConcurrent
    #
    if (-not $MaxConcurrent) {
        Add-RuntimeErrorHC -Type 'FatalError' -Name "Missing 'MaxConcurrent'" `
            -Message 'MaxConcurrent block missing.' `
            -SystemErrors $SystemErrors
    }
    else {
        foreach ($p in 'Computers', 'FoldersPerMatrix', 'JobsPerRemoteComputer') {
            if (-not $MaxConcurrent.$p -or $MaxConcurrent.$p -notmatch '^\d+$') {
                Add-RuntimeErrorHC -Type 'FatalError' -Name "Incorrect 'MaxConcurrent.$p'" `
                    -Message "$p must be numeric." `
                    -SystemErrors $SystemErrors
            }
        }
    }


    #
    # 4. Export & ServiceNow validation
    #
    if ($Export) {

        if ($Export.PermissionsExcelFile -and $Export.PermissionsExcelFile -notmatch '\.xlsx$') {
            Add-RuntimeErrorHC -Type 'FatalError' -Name "Incorrect 'Export.PermissionsExcelFile'" `
                -Message 'Must end with .xlsx' `
                -SystemErrors $SystemErrors
        }

        if ($Export.OverviewHtmlFile -and $Export.OverviewHtmlFile -notmatch '\.html?$') {
            Add-RuntimeErrorHC -Type 'FatalError' -Name "Incorrect 'Export.OverviewHtmlFile'" `
                -Message 'Must end with .html' `
                -SystemErrors $SystemErrors
        }

        if ($Export.ServiceNowFormDataExcelFile) {

            #
            # validate extension
            #
            if ($Export.ServiceNowFormDataExcelFile -notmatch '\.xlsx$') {
                Add-RuntimeErrorHC -Type 'FatalError' `
                    -Name "Incorrect 'Export.ServiceNowFormDataExcelFile'" `
                    -Message 'Must end with .xlsx' `
                    -SystemErrors $SystemErrors
            }

            #
            # ServiceNow must exist
            #
            if (-not $ServiceNow) {
                Add-RuntimeErrorHC -Type 'FatalError' `
                    -Name 'Incorrect configuration' `
                    -Message 'ServiceNow must be defined when using ServiceNowFormDataExcelFile.' `
                    -SystemErrors $SystemErrors
            }
            else {
                foreach ($p in 'CredentialsFilePath', 'TableName', 'Environment') {
                    if (-not $ServiceNow.$p) {
                        Add-RuntimeErrorHC -Type 'FatalError' `
                            -Name "Missing 'ServiceNow.$p'" `
                            -Message "$p is required." `
                            -SystemErrors $SystemErrors
                    }
                }
            }
        }
    }
}