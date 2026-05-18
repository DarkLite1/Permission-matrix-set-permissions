
function Add-ErrorHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Type,       
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Message,
        [Parameter()][string]$Description = '',
        [Parameter(Mandatory)][string]$Category,
        [Parameter(Mandatory)][ref]$SystemErrors
    )

    $SystemErrors.Value.Add(
        [PSCustomObject]@{
            DateTime    = Get-Date
            Type        = $Type
            Name        = $Name
            Message     = $Message
            Description = $Description
            Category    = $Category
        }
    )
}

function Add-MatrixErrorHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Type,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Message,
        [string]$Description = '',
        [Parameter(Mandatory)][ref]$SystemErrors
    )

    Add-ErrorHC -Category 'Matrix' @PSBoundParameters
}

function Add-PermissionsErrorHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Type,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Message,
        [string]$Description = '',
        [Parameter(Mandatory)][ref]$SystemErrors
    )

    Add-ErrorHC -Category 'Permissions' @PSBoundParameters
}

function Add-RuntimeErrorHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Type,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Message,
        [string]$Description = '',
        [Parameter(Mandatory)][ref]$SystemErrors
    )

    Add-ErrorHC -Category 'RuntimeSettings' @PSBoundParameters
}

function Add-JsonSchemaErrorHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Type,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Message,
        [string]$Description = '',
        [Parameter(Mandatory)][ref]$SystemErrors
    )

    Add-ErrorHC -Category 'JsonSchema' @PSBoundParameters
}

function Get-StringValueHC {
    [CmdletBinding()]
    param([String]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return $null
    }
    elseif ($Name.StartsWith('ENV:', [System.StringComparison]::OrdinalIgnoreCase)) {
        $envVariableName = $Name.Substring(4).Trim()
        $envStringValue = Get-Item -Path "Env:\$envVariableName" -EA Ignore

        if ($envStringValue) {
            return $envStringValue.Value
        }
        else {
            throw "Environment variable '$envVariableName' not found."
        }
    }
    else {
        return $Name
    }
}

function Get-StringOrDefaultHC {
    <#
    .SYNOPSIS
        Returns $Default when $Value is $null, an empty string, or whitespace only.
        Otherwise returns $Value unchanged.

    .DESCRIPTION
        Display/fallback companion to Get-StringValueHC. Useful for rendering
        a placeholder when a string-shaped input is missing or blank
        (e.g. 'Unknown', 'N/A', or a sensible default like 'Permission Matrix'
        for a missing script name).

        Treats these as blank:
          - $null
          - ''
          - '   ' (whitespace only)
          - any object whose [string] conversion is null/whitespace

        Treats these as non-blank (passed through):
          - 0, $false, empty arrays, empty hashtables (they stringify to non-blank)

        If you need empty-array or empty-collection fallback behaviour,
        write a separate helper — don't extend this one.

        Note: this function does NOT resolve 'ENV:' prefixes. Use
        Get-StringValueHC for config strings that may reference environment
        variables.

    .PARAMETER Value
        The value to check. Any type; coerced via [string] for the blank check.

    .PARAMETER Default
        The fallback returned when Value is blank.

    .EXAMPLE
        Get-StringOrDefaultHC -Value $row.Name -Default 'Unknown'

    .EXAMPLE
        $row.Name | Get-StringOrDefaultHC 'Unknown'

    .EXAMPLE
        [System.Net.WebUtility]::HtmlEncode(
            (Get-StringOrDefaultHC $excel.LastModifiedBy 'Unknown')
        )
    #>
    [CmdletBinding()]
    [OutputType([object])]
    param(
        [Parameter(Position = 0, ValueFromPipeline)]
        [AllowNull()]
        $Value,

        [Parameter(Position = 1, Mandatory)]
        [AllowEmptyString()]
        [string]$Default
    )

    process {
        if ([string]::IsNullOrWhiteSpace([string]$Value)) {
            $Default
        }
        else {
            $Value
        }
    }
}

function Get-DatedLogFolderPathHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$LogFolder,
        [Parameter(Mandatory)][datetime]$ScriptStartTime,
        [Parameter(Mandatory)][object]$JsonFile
    )

    try {
        $datedLogFolder = Join-Path -Path $LogFolder -ChildPath (
            '{0:0000}_{1:00}_{2:00}_{3:00}{4:00}{5:00} ({6})' -f 
            $ScriptStartTime.Year,
            $ScriptStartTime.Month,
            $ScriptStartTime.Day,
            $ScriptStartTime.Hour,
            $ScriptStartTime.Minute,
            $ScriptStartTime.Second,
            $JsonFile
        )

        return (New-Item -ItemType 'Directory' -Path $datedLogFolder -Force -EA Stop).FullName
    }
    catch {
        return $LogFolder
    }
}

function Plural {
    [CmdletBinding()]
    param(
        [int]$Count,
        [string]$Word
    )

    if ($Count -eq 1) { return $Word }
    return "$Word`s"
}

function Test-ItemHasFatalErrorHC {
    <#
    .SYNOPSIS
        Checks if a localized validation list (like $MatrixObj.Check or 
        $Setting.Check) contains any terminating FatalErrors.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [array]$CheckList
    )

    if (-not $CheckList -or $CheckList.Count -eq 0) { 
        return $false 
    }

    return ($CheckList.Type -contains 'FatalError')
}

function New-CounterObjectHC {
    <#
    .SYNOPSIS
        Initializes an empty counter object for tracking errors and warnings.
    #>
    [CmdletBinding()]
    param()

    return [PSCustomObject]@{
        TotalErrors   = 0
        TotalWarnings = 0
        FormData      = [PSCustomObject]@{ Errors = 0; Warnings = 0 }
        Permissions   = [PSCustomObject]@{ Errors = 0; Warnings = 0 }
        Settings      = [PSCustomObject]@{ Errors = 0; Warnings = 0 }
        File          = [PSCustomObject]@{ Errors = 0; Warnings = 0 }
    }
}

function Update-MatrixCounterHC {
    <#
    .SYNOPSIS
        Calculates the total errors and warnings across all matrices and system errors.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Context,
        [Parameter(Mandatory)][ref]$SystemErrors
    )

    # Reset counter to ensure a clean tally
    $Context.Counter = New-CounterObjectHC

    if ($Context.Matrices) {
        foreach ($matrix in $Context.Matrices) {
            # File Checks
            $Context.Counter.File.Errors += @($matrix.File.Check | Where-Object { $_.Type -eq 'FatalError' }).Count
            $Context.Counter.File.Warnings += @($matrix.File.Check | Where-Object { $_.Type -eq 'Warning' }).Count

            # FormData Checks
            $Context.Counter.FormData.Errors += @($matrix.FormData.Check | Where-Object { $_.Type -eq 'FatalError' }).Count
            $Context.Counter.FormData.Warnings += @($matrix.FormData.Check | Where-Object { $_.Type -eq 'Warning' }).Count

            # Permissions Checks
            $Context.Counter.Permissions.Errors += @($matrix.Permissions.Check | Where-Object { $_.Type -eq 'FatalError' }).Count
            $Context.Counter.Permissions.Warnings += @($matrix.Permissions.Check | Where-Object { $_.Type -eq 'Warning' }).Count

            # Settings Checks
            foreach ($setting in $matrix.Settings) {
                $Context.Counter.Settings.Errors += @($setting.Check | Where-Object { $_.Type -eq 'FatalError' }).Count
                $Context.Counter.Settings.Warnings += @($setting.Check | Where-Object { $_.Type -eq 'Warning' }).Count
            }
        }
    }

    # Tally Totals (Including Orchestrator SystemErrors) [cite: 1612-1616]
    $Context.Counter.TotalErrors = $Context.Counter.File.Errors + 
    $Context.Counter.FormData.Errors + 
    $Context.Counter.Permissions.Errors + 
    $Context.Counter.Settings.Errors + 
    $SystemErrors.Value.Count

    $Context.Counter.TotalWarnings = $Context.Counter.File.Warnings + 
    $Context.Counter.FormData.Warnings + 
    $Context.Counter.Permissions.Warnings + 
    $Context.Counter.Settings.Warnings
                                     
    return $Context.Counter
}