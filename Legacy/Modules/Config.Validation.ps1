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