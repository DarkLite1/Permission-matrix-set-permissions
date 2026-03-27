function Ensure-SafeSettingsHC {
    param([object]$Settings)

    if (-not $Settings) {
        # Create a safe dummy object with defaults so END block cannot break
        return [PSCustomObject]@{
            ScriptName     = 'Default script name'
            SaveLogFiles   = @{
                Where               = @{
                    Folder = $null 
                } 
                Detailed            = $false 
                DeleteLogsAfterDays = 0 
            }
            SaveInEventLog = @{ 
                Save    = $false 
                LogName = $null 
            }
            SendMail       = @{ 
                From = $null 
                To   = @()
                Body = $null 
                Smtp = @{ Port     = 25
                    ConnectionType = 'None' 
                } 
            }
        }
    }

    # Ensure ScriptName exists and is safe for filenames
    if ([string]::IsNullOrWhiteSpace($Settings.ScriptName)) {
        $Settings.ScriptName = 'Default script name'
    }

    if (-not $Settings.SaveLogFiles) {
        $Settings | Add-Member -NotePropertyName SaveLogFiles -NotePropertyValue @{
            Where               = @{ 
                Folder = $null 
            }
            Detailed            = $false
            DeleteLogsAfterDays = 0
        }
    }

    if (-not $Settings.SaveInEventLog) {
        $Settings | Add-Member -NotePropertyName SaveInEventLog -NotePropertyValue @{
            Save    = $false
            LogName = $null
        }
    }

    if (-not $Settings.SendMail) {
        $Settings | Add-Member -NotePropertyName SendMail -NotePropertyValue @{
            From = $null
            To   = @()
            Body = $null
            Smtp = @{
                Port           = 25
                ConnectionType = 'None'
            }
        }
    }

    return $Settings
}
function Ensure-LogFolderHC {
    param(
        [Parameter()]
        [string]$RequestedFolder,

        [Parameter()]
        [ref]$SystemErrors
    )

    #
    # 1 - If requested folder is null or empty → immediate fallback
    #
    if ([string]::IsNullOrWhiteSpace($RequestedFolder)) {

        $fallback = Join-Path $env:TEMP 'PermissionMatrixLogs'

        try {
            if (-not (Test-Path -LiteralPath $fallback -PathType Container)) {
                New-Item -ItemType Directory -Path $fallback -ErrorAction Stop | Out-Null
            }
        }
        catch {
            # Last-resort fallback (extremely rare)
            $fallback = $env:TEMP
        }

        if ($SystemErrors) {
            $SystemErrors.Value.Add(
                [PSCustomObject]@{
                    DateTime = Get-Date
                    Message  = "LogFolder missing. Using fallback folder '$fallback'."
                }
            )
        }

        return $fallback
    }

    #
    # 2 - Try to create/verify the requested folder
    #
    try {
        if (-not (Test-Path -LiteralPath $RequestedFolder -PathType Container)) {
            New-Item -ItemType Directory -Path $RequestedFolder -ErrorAction Stop | Out-Null
        }

        return $RequestedFolder
    }
    catch {

        #
        # 3 - Requested folder invalid → use fallback
        #
        $fallback = Join-Path $env:TEMP 'PermissionMatrixLogs'

        try {
            if (-not (Test-Path -LiteralPath $fallback -PathType Container)) {
                New-Item -ItemType Directory -Path $fallback -ErrorAction Stop | Out-Null
            }
        }
        catch {
            # last resort (TEMP folder is always safe)
            $fallback = $env:TEMP
        }

        if ($SystemErrors) {
            $SystemErrors.Value.Add(
                [PSCustomObject]@{
                    DateTime = Get-Date
                    Message  = "LogFolder '$RequestedFolder' invalid or uncreatable. Using fallback '$fallback'. Error: $_"
                }
            )
        }

        return $fallback
    }
}
function Validate-RuntimeSettings {
    param(
        [object]$Settings,
        [object]$Matrix,
        [object]$Export,
        [object]$ServiceNow,
        [object]$MaxConcurrent,
        [ref]$SystemErrors
    )

    # ---------------------------
    # 1. Base Settings Validation
    # ---------------------------
    if (-not $Settings) { 
        Add-RuntimeErrorHC 'FatalError' 'Invalid configuration' "Property 'Settings' missing from JSON."
        return [PSCustomObject]@{ IsValid = $false; Errors = $errors; Settings = $null }
    }

    if ([string]::IsNullOrWhiteSpace($Settings.ScriptName)) {
        Add-RuntimeErrorHC 'Warning' 'Missing Script Name' "No 'Settings.ScriptName' found in JSON. A default name will be used."
        $Settings | Add-Member -NotePropertyName ScriptName -NotePropertyValue 'Default script name' -Force
    }

    if ($Settings.SaveLogFiles.Detailed -isnot [bool]) {
        Add-RuntimeErrorHC 'FatalError' 'Invalid type' 'Settings.SaveLogFiles.Detailed must be a boolean.'
    }

    if ($Settings.SaveInEventLog.Save -isnot [bool]) {
        Add-RuntimeErrorHC 'FatalError' 'Invalid type' 'Settings.SaveInEventLog.Save must be a boolean.'
    }

    if ([string]::IsNullOrWhiteSpace($Settings.SendMail.From)) {
        Add-RuntimeErrorHC 'FatalError' 'Invalid configuration' 'Settings.SendMail.From cannot be empty.'
    }

    if (-not $Settings.SendMail.To) {
        Add-RuntimeErrorHC 'FatalError' 'Invalid configuration' 'Settings.SendMail.To cannot be empty.'
    }
    elseif ($Settings.SendMail.To -isnot [array] -and $Settings.SendMail.To -isnot [string]) {
        Add-RuntimeErrorHC 'FatalError' 'Invalid type' 'Settings.SendMail.To must be an array or a string.'
    }
        
    if ([string]::IsNullOrWhiteSpace($Settings.SendMail.Body)) {
        Add-RuntimeErrorHC 'FatalError' 'Invalid configuration' 'Settings.SendMail.Body cannot be empty.'
    }
        
    if (-not $Settings.SendMail.Smtp.Port -or $Settings.SendMail.Smtp.Port -notmatch '^\d+$') {
        Add-RuntimeErrorHC 'FatalError' 'Invalid configuration' 'Settings.SendMail.Smtp.Port must be an integer.'
    }
        
    $validConnections = @('None', 'Auto', 'SslOnConnect', 'StartTls', 'StartTlsWhenAvailable')
    if ($Settings.SendMail.Smtp.ConnectionType -notin $validConnections) {
        Add-RuntimeErrorHC 'FatalError' 'Invalid configuration' "Settings.SendMail.Smtp.ConnectionType must be one of: $($validConnections -join ', ')."
    }

    # ---------------------------
    # 2. Matrix Validation
    # ---------------------------
    if ($Matrix) {
        if (-not (Test-Path -LiteralPath $Matrix.DefaultsFile -PathType Leaf)) {
            Add-RuntimeErrorHC 'FatalError' 'Invalid path' "Matrix.DefaultsFile '$($Matrix.DefaultsFile)' does not exist or is not a file."
        }
            
        # SCHEMA CHECK ONLY: Ensure the property is populated before the BEGIN block attempts to map it
        if ([string]::IsNullOrWhiteSpace($Matrix.FolderPath)) {
            Add-RuntimeErrorHC 'FatalError' 'Invalid configuration' 'Matrix.FolderPath cannot be empty.'
        }
            
        if ($Matrix.ExcludedSamAccountName -isnot [array]) {
            Add-RuntimeErrorHC 'FatalError' 'Invalid type' 'Matrix.ExcludedSamAccountName must be an array.'
        }
    }
    else {
        Add-RuntimeErrorHC 'FatalError' 'Invalid configuration' "Property 'Matrix' missing."
    }

    # ---------------------------
    # 3. MaxConcurrent Validation
    # ---------------------------
    if ($MaxConcurrent) {
        foreach ($prop in @('Computers', 'FoldersPerMatrix', 'JobsPerRemoteComputer')) {
            if (-not $MaxConcurrent.$prop -or $MaxConcurrent.$prop -notmatch '^\d+$') {
                Add-RuntimeErrorHC 'FatalError' 'Invalid type' "MaxConcurrent.$prop must be an integer."
            }
        }
    }
    else {
        Add-RuntimeErrorHC 'FatalError' 'Invalid configuration' "Property 'MaxConcurrent' missing."
    }

    # ---------------------------
    # 4. Export & ServiceNow
    # ---------------------------
    if ($Export) {
        if (-not [string]::IsNullOrWhiteSpace($Export.PermissionsExcelFile) -and $Export.PermissionsExcelFile -notmatch '\.xlsx$') {
            Add-RuntimeErrorHC 'FatalError' 'Invalid path' 'Export.PermissionsExcelFile must end in .xlsx.'
        }
        if (-not [string]::IsNullOrWhiteSpace($Export.OverviewHtmlFile) -and $Export.OverviewHtmlFile -notmatch '\.html?$') {
            Add-RuntimeErrorHC 'FatalError' 'Invalid path' 'Export.OverviewHtmlFile must end in .html.'
        }
        if (-not [string]::IsNullOrWhiteSpace($Export.ServiceNowFormDataExcelFile)) {
            if ($Export.ServiceNowFormDataExcelFile -notmatch '\.xlsx$') {
                Add-RuntimeErrorHC 'FatalError' 'Invalid path' 'Export.ServiceNowFormDataExcelFile must end in .xlsx.'
            }
            if (-not $ServiceNow) {
                Add-RuntimeErrorHC 'FatalError' 'Invalid configuration' 'ServiceNow configuration object is required when Export.ServiceNowFormDataExcelFile is populated.'
            }
            else {
                if ([string]::IsNullOrWhiteSpace($ServiceNow.CredentialsFilePath)) { 
                    Add-RuntimeErrorHC 'FatalError' 'Invalid configuration' 'ServiceNow.CredentialsFilePath is required.' 
                }
                if ([string]::IsNullOrWhiteSpace($ServiceNow.TableName)) { 
                    Add-RuntimeErrorHC 'FatalError' 'Invalid configuration' 'ServiceNow.TableName is required.' 
                }
                if ([string]::IsNullOrWhiteSpace($ServiceNow.Environment)) { 
                    Add-RuntimeErrorHC 'FatalError' 'Invalid configuration' 'ServiceNow.Environment is required.' 
                }
            }
        }
    }
}
function ConvertTo-StructuredObjectHC {
    <#
        .SYNOPSIS
            Normalizes various input types into a standard PSCustomObject for 
            HealthCheck reports.
        
        .DESCRIPTION
            This function takes strings, hashtables, or existing objects and 
            ensures they conform to a specific schema (Name, Description, Type, 
            Value). If properties are missing or  null, it injects "Missing 
            data" and sets the Type to "FatalError".
        
        .PARAMETER Objects
            The input data to be converted. Can be a single item or an array.
        
        .EXAMPLE
            "System Error" | ConvertTo-StructuredObjectHC

            Converts a simple string into a full object with Name: 
            'Error during execution'.
        
        .EXAMPLE
            ConvertTo-StructuredObjectHC -Objects @{ 
                Name = "Disk Check"
                Type = "Warning" 
            }
            Converts a hashtable into an object and fills in the missing 
            'Description' property.
        #>

    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [AllowEmptyCollection()]
        [array]$Objects
    )

    process {
        foreach ($checkObj in @($Objects)) {
            if ($null -eq $checkObj) { continue }

            # 1. Normalize into a PSCustomObject
            $current = if (
                $checkObj -is [string] -or 
                $checkObj -is [System.ValueType]
            ) {
                [PSCustomObject]@{
                    Name        = 'Error during execution'
                    Description = "Primitive value received: $checkObj"
                    Type        = 'FatalError'
                    Value       = $checkObj
                }
            }
            elseif ($checkObj -is [hashtable]) {
                [PSCustomObject]$checkObj
            }
            else {
                # Force a cast to ensure the object is extensible/malleable
                [PSCustomObject]$checkObj
            }

            # 2. Ensure the 'Value' property exists (avoiding null reference errors later)
            if (-not (Get-Member -InputObject $current -Name 'Value')) {
                $current | Add-Member -MemberType NoteProperty -Name 'Value' -Value $null
            }

            # 3. Validate Core Properties
            foreach ($prop in @('Name', 'Description', 'Type')) {
                if ([string]::IsNullOrWhiteSpace($current.$prop)) {
                    $current | Add-Member -MemberType NoteProperty -Name $prop -Value 'Missing data' -Force
                    $current | Add-Member -MemberType NoteProperty -Name 'Type' -Value 'FatalError' -Force
                }
            }

            # Output the object to the pipeline
            $current
        }
    }
}