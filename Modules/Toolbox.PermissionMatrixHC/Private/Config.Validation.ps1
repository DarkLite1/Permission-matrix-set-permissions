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
        $prop = 'Settings'
        Add-RuntimeErrorHC `
            -Type 'FatalError' `
            -Name "Missing '$prop'" `
            -Message "Property '$prop' not found in JSON." `
            -Description "The JSON input file must contain property '$prop'" `
            -SystemErrors ([ref]$SystemErrors)

        return
    }

    if ([string]::IsNullOrWhiteSpace($Settings.ScriptName)) {
        $prop = 'Settings.ScriptName'
        Add-RuntimeErrorHC `
            -Type 'Warning' `
            -Name "Missing '$prop'" `
            -Message "Property '$prop' not found in JSON." `
            -Description "The JSON input file must contain property '$prop'" `
            -SystemErrors ([ref]$SystemErrors)
        
        $Settings | Add-Member -NotePropertyName ScriptName -NotePropertyValue 'Default script name' -Force
    }

    if ($Settings.SaveLogFiles.Detailed -isnot [bool]) {
        $prop = 'Settings.SaveLogFiles.Detailed'
        Add-RuntimeErrorHC `
            -Type 'FatalError' `
            -Name "Incorrect '$prop'" `
            -Message "Property '$prop' needs to be a boolean." `
            -Description "The JSON input file property '$prop' needs to be a boolean" `
            -SystemErrors ([ref]$SystemErrors)
    }

    if ($Settings.SaveInEventLog.Save -isnot [bool]) {
        $prop = 'Settings.SaveInEventLog.Save'
        Add-RuntimeErrorHC `
            -Type 'FatalError' `
            -Name "Incorrect '$prop'" `
            -Message "Property '$prop' needs to be a boolean." `
            -Description "The JSON input file property '$prop' needs to be a boolean" `
            -SystemErrors ([ref]$SystemErrors)
    }

    if ([string]::IsNullOrWhiteSpace($Settings.SendMail.From)) {
        $prop = 'Settings.SendMail.From'
        Add-RuntimeErrorHC `
            -Type 'FatalError' `
            -Name "Missing '$prop'" `
            -Message "Property '$prop' not found in JSON." `
            -Description "The JSON input file must contain property '$prop'" `
            -SystemErrors ([ref]$SystemErrors)
    }

    if (-not $Settings.SendMail.To) {
        $prop = 'Settings.SendMail.To'
        Add-RuntimeErrorHC `
            -Type 'FatalError' `
            -Name "Missing '$prop'" `
            -Message "Property '$prop' not found in JSON." `
            -Description "The JSON input file must contain property '$prop'" `
            -SystemErrors ([ref]$SystemErrors)
    }
    elseif ($Settings.SendMail.To -isnot [array] -and $Settings.SendMail.To -isnot [string]) {
        $prop = 'Settings.SendMail.To'
        Add-RuntimeErrorHC `
            -Type 'FatalError' `
            -Name "Incorrect '$prop'" `
            -Message "Property '$prop' needs to be an array or a string." `
            -Description "The JSON input file property '$prop' needs to be ab array or a string" `
            -SystemErrors ([ref]$SystemErrors)
    }
        
    if ([string]::IsNullOrWhiteSpace($Settings.SendMail.Body)) {
        $prop = 'Settings.SendMail.Body'
        Add-RuntimeErrorHC `
            -Type 'FatalError' `
            -Name "Missing '$prop'" `
            -Message "Property '$prop' not found in JSON." `
            -Description "The JSON input file must contain property '$prop'" `
            -SystemErrors ([ref]$SystemErrors)
    }
        
    if (-not $Settings.SendMail.Smtp.Port -or $Settings.SendMail.Smtp.Port -notmatch '^\d+$') {
        $prop = 'SendMail.Smtp.Port'
        Add-RuntimeErrorHC `
            -Type 'FatalError' `
            -Name "Incorrect '$prop'" `
            -Message "Property '$prop' needs to be a number." `
            -Description "The JSON input file property '$prop' needs to be a number" `
            -SystemErrors ([ref]$SystemErrors)
    }
        
    $validConnections = @(
        'None', 'Auto', 'SslOnConnect', 'StartTls', 'StartTlsWhenAvailable'
    )

    if ($Settings.SendMail.Smtp.ConnectionType -notin $validConnections) {
        $prop = 'Settings.SendMail.Smtp.ConnectionType'
        Add-RuntimeErrorHC `
            -Type 'FatalError' `
            -Name "Incorrect '$prop'" `
            -Message "Property '$prop' has an incorrect value." `
            -Description "The JSON input file property '$prop' needs to have one of the following values: $($validConnections -join ', ')." `
            -SystemErrors ([ref]$SystemErrors)
    }

    # ---------------------------
    # 2. Matrix Validation
    # ---------------------------
    if ($Matrix) {
        if (-not (Test-Path -LiteralPath $Matrix.DefaultsFile -PathType Leaf)) {
            $prop = 'Matrix.DefaultsFile'
            Add-RuntimeErrorHC `
                -Type 'FatalError' `
                -Name "Incorrect '$prop'" `
                -Message "Path '$($Matrix.DefaultsFile)' in property '$prop' not found." `
                -Description "The JSON input file property '$prop' with path '$($Matrix.DefaultsFile)': path not found." `
                -SystemErrors ([ref]$SystemErrors)
        }
            
        if ([string]::IsNullOrWhiteSpace($Matrix.FolderPath)) {
            $prop = 'Matrix.FolderPath'
            Add-RuntimeErrorHC `
                -Type 'FatalError' `
                -Name "Missing '$prop'" `
                -Message "Property '$prop' not found in JSON." `
                -Description "The JSON input file must contain property '$prop'" `
                -SystemErrors ([ref]$SystemErrors)
        }
            
        if ($Matrix.ExcludedSamAccountName -isnot [array]) {
            $prop = 'Matrix.ExcludedSamAccountName'
            Add-RuntimeErrorHC `
                -Type 'FatalError' `
                -Name "Missing '$prop'" `
                -Message "Property '$prop' must be an array." `
                -Description "The JSON input file property '$prop' must be an array" `
                -SystemErrors ([ref]$SystemErrors)
        }
    }
    else {
        $prop = 'Matrix'
        Add-RuntimeErrorHC `
            -Type 'FatalError' `
            -Name "Missing '$prop'" `
            -Message "Property '$prop' not found in JSON." `
            -Description "The JSON input file must contain property '$prop'" `
            -SystemErrors ([ref]$SystemErrors)
    }

    # ---------------------------
    # 3. MaxConcurrent Validation
    # ---------------------------
    if ($MaxConcurrent) {
        foreach ($prop in @('Computers', 'FoldersPerMatrix', 'JobsPerRemoteComputer')) {
            if (-not $MaxConcurrent.$prop -or $MaxConcurrent.$prop -notmatch '^\d+$') {
                Add-RuntimeErrorHC `
                    -Type 'FatalError' `
                    -Name "Incorrect 'MaxConcurrent.$prop'" `
                    -Message "Property 'MaxConcurrent.$prop' needs to be a number." `
                    -Description "The JSON input file property '$prop' needs to be a number" `
                    -SystemErrors ([ref]$SystemErrors)
            }
        }
    }
    else {
        $prop = 'MaxConcurrent'
        Add-RuntimeErrorHC `
            -Type 'FatalError' `
            -Name "Missing '$prop'" `
            -Message "Property '$prop' not found in JSON." `
            -Description "The JSON input file must contain property '$prop'" `
            -SystemErrors ([ref]$SystemErrors)
    }

    # ---------------------------
    # 4. Export & ServiceNow
    # ---------------------------
    if ($Export) {
        if (-not [string]::IsNullOrWhiteSpace($Export.PermissionsExcelFile) -and $Export.PermissionsExcelFile -notmatch '\.xlsx$') {
            $prop = 'Export.PermissionsExcelFile'
            Add-RuntimeErrorHC `
                -Type 'FatalError' `
                -Name "Incorrect '$prop'" `
                -Message "Property '$prop' must end with '.xlsx'." `
                -Description "The JSON input file property '$prop' must end with '.xlsx'." `
                -SystemErrors ([ref]$SystemErrors)
        }
        if (-not [string]::IsNullOrWhiteSpace($Export.OverviewHtmlFile) -and $Export.OverviewHtmlFile -notmatch '\.html?$') {
            $prop = 'Export.OverviewHtmlFile'
            Add-RuntimeErrorHC `
                -Type 'FatalError' `
                -Name "Incorrect '$prop'" `
                -Message "Property '$prop' must end with '.html'." `
                -Description "The JSON input file property '$prop' must end with '.html'." `
                -SystemErrors ([ref]$SystemErrors)
        }
        if (-not [string]::IsNullOrWhiteSpace($Export.ServiceNowFormDataExcelFile)) {
            if ($Export.ServiceNowFormDataExcelFile -notmatch '\.xlsx$') {
                $prop = 'Export.ServiceNowFormDataExcelFile'
                Add-RuntimeErrorHC `
                    -Type 'FatalError' `
                    -Name "Incorrect '$prop'" `
                    -Message "Property '$prop' must end with '.xlsx'." `
                    -Description "The JSON input file property '$prop' must end with '.xlsx'." `
                    -SystemErrors ([ref]$SystemErrors)
            }
            if (-not $ServiceNow) {
                Add-RuntimeErrorHC `
                    -Type 'FatalError' `
                    -Name 'Incorrect configuration' `
                    -Message "Property 'ServiceNow' must be defined when 'ServiceNowFormDataExcelFile' is used." `
                    -Description 'The JSON input file property 'ServiceNow' must be defined when 'ServiceNowFormDataExcelFile' is used.' `
                    -SystemErrors ([ref]$SystemErrors)
            }
            else {
                if ([string]::IsNullOrWhiteSpace($ServiceNow.CredentialsFilePath)) { 
                    $prop = 'ServiceNow.CredentialsFilePath'
                    Add-RuntimeErrorHC `
                        -Type 'FatalError' `
                        -Name "Incorrect '$prop'" `
                        -Message "Property '$prop' not found in JSON." `
                        -Description "The JSON input file must contain property '$prop'" `
                        -SystemErrors ([ref]$SystemErrors)
                }
                if ([string]::IsNullOrWhiteSpace($ServiceNow.TableName)) { 
                    $prop = 'ServiceNow.TableName'
                    Add-RuntimeErrorHC `
                        -Type 'FatalError' `
                        -Name "Incorrect '$prop'" `
                        -Message "Property '$prop' not found in JSON." `
                        -Description "The JSON input file must contain property '$prop'" `
                        -SystemErrors ([ref]$SystemErrors)
                }
                if ([string]::IsNullOrWhiteSpace($ServiceNow.Environment)) { 
                    $prop = 'ServiceNow.Environment'
                    Add-RuntimeErrorHC `
                        -Type 'FatalError' `
                        -Name "Incorrect '$prop'" `
                        -Message "Property '$prop' not found in JSON." `
                        -Description "The JSON input file must contain property '$prop'" `
                        -SystemErrors ([ref]$SystemErrors)
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