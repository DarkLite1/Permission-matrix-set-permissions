
function Add-ErrorHC {
    <#
    .SYNOPSIS
        Append a structured error record to a system-error accumulator.

    .DESCRIPTION
        Builds a structured error object and adds it to the collection
        referenced by SystemErrors. The record captures the moment it was
        created (DateTime is stamped via Get-Date at call time) together with
        the supplied classification and message fields: Type, Name, Message,
        Description and Category.

        The function records rather than throws: it appends to the accumulator
        and returns, leaving the caller to inspect the collected errors and
        decide how to proceed.

        SystemErrors is passed by reference and must point at a collection that
        exposes an .Add() method (for example [System.Collections.ArrayList] or
        [System.Collections.Generic.List[object]]). A fixed-size array created
        with @() does not support .Add() and causes a terminating error.

    .PARAMETER Type
        The error severity or kind, for example 'FatalError' or 'Warning'. Free
        text; not validated against a fixed set.

    .PARAMETER Name
        A short title for the error, used to identify or group similar
        problems.

    .PARAMETER Message
        The human-readable description of what went wrong.

    .PARAMETER Description
        Optional additional detail or remediation guidance. Defaults to an
        empty string.

    .PARAMETER Category
        The error category, for example 'Matrix', 'Permissions',
        'RuntimeSettings' or 'JsonSchema'. The Add-*ErrorHC wrappers each supply
        a fixed value here.

    .PARAMETER SystemErrors
        A [ref] to the caller's error accumulator: a collection supporting
        .Add(). The new record is appended to SystemErrors.Value. This is an
        in/out parameter.

    .EXAMPLE
        $errors = [System.Collections.Generic.List[object]]::new()
        Add-ErrorHC -Type 'FatalError' -Name 'Bad row' -Message 'Missing path.' -Category 'Matrix' -SystemErrors ([ref]$errors)

        Appends one error record to $errors, with DateTime set to the current
        time and Description left empty.

    .OUTPUTS
        None. The function appends to the referenced collection and returns
        nothing.

    .NOTES
        - The function records errors; it does not throw. Callers inspect the
          accumulator afterwards.
        - DateTime is captured with Get-Date at the moment of the call (local
          time).
        - If SystemErrors.Value is a [System.Collections.ArrayList], its .Add()
          returns the insertion index, which would leak onto the pipeline. A
          generic List[T] returns void and avoids this.

    .LINK
        Add-MatrixErrorHC
    .LINK
        Add-PermissionsErrorHC
    .LINK
        Add-RuntimeErrorHC
    .LINK
        Add-JsonSchemaErrorHC
    #>

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
    <#
    .SYNOPSIS
        Add a 'Matrix'-category error to the system-error accumulator.

    .DESCRIPTION
        Thin wrapper around Add-ErrorHC that fixes Category to 'Matrix'. All
        other parameters are forwarded unchanged. See Add-ErrorHC for the full
        description of the record created and how SystemErrors is used.

    .PARAMETER Type
        The error severity or kind (for example 'FatalError'). Forwarded to
        Add-ErrorHC.

    .PARAMETER Name
        A short title for the error. Forwarded to Add-ErrorHC.

    .PARAMETER Message
        The human-readable description of the problem. Forwarded to
        Add-ErrorHC.

    .PARAMETER Description
        Optional additional detail. Defaults to an empty string. Forwarded to
        Add-ErrorHC.

    .PARAMETER SystemErrors
        A [ref] to the caller's error accumulator (a collection supporting
        .Add()). Forwarded to Add-ErrorHC.

    .EXAMPLE
        $errors = [System.Collections.Generic.List[object]]::new()
        Add-MatrixErrorHC -Type 'FatalError' -Name 'Duplicate entry' -Message "'GRP' defined twice." -SystemErrors ([ref]$errors)

        Appends a 'Matrix'-category error to $errors.

    .OUTPUTS
        None. Appends to the referenced collection and returns nothing.

    .NOTES
        Category is fixed to 'Matrix' and cannot be overridden through this
        function.

    .LINK
        Add-ErrorHC
    #>
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
    <#
    .SYNOPSIS
        Add a 'Permissions'-category error to the system-error accumulator.

    .DESCRIPTION
        Thin wrapper around Add-ErrorHC that fixes Category to 'Permissions'.
        All other parameters are forwarded unchanged. See Add-ErrorHC for the
        full description of the record created and how SystemErrors is used.

    .PARAMETER Type
        The error severity or kind (for example 'FatalError'). Forwarded to
        Add-ErrorHC.

    .PARAMETER Name
        A short title for the error. Forwarded to Add-ErrorHC.

    .PARAMETER Message
        The human-readable description of the problem. Forwarded to
        Add-ErrorHC.

    .PARAMETER Description
        Optional additional detail. Defaults to an empty string. Forwarded to
        Add-ErrorHC.

    .PARAMETER SystemErrors
        A [ref] to the caller's error accumulator (a collection supporting
        .Add()). Forwarded to Add-ErrorHC.

    .EXAMPLE
        $errors = [System.Collections.Generic.List[object]]::new()
        Add-PermissionsErrorHC -Type 'FatalError' -Name 'Invalid permission' -Message "Unknown permission 'X'." -SystemErrors ([ref]$errors)

        Appends a 'Permissions'-category error to $errors.

    .OUTPUTS
        None. Appends to the referenced collection and returns nothing.

    .NOTES
        Category is fixed to 'Permissions' and cannot be overridden through this
        function.

    .LINK
        Add-ErrorHC
    #>
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
    <#
    .SYNOPSIS
        Add a 'RuntimeSettings'-category error to the system-error accumulator.

    .DESCRIPTION
        Thin wrapper around Add-ErrorHC that fixes Category to 'RuntimeSettings'.
        All other parameters are forwarded unchanged. See Add-ErrorHC for the
        full description of the record created and how SystemErrors is used.

    .PARAMETER Type
        The error severity or kind (for example 'FatalError'). Forwarded to
        Add-ErrorHC.

    .PARAMETER Name
        A short title for the error. Forwarded to Add-ErrorHC.

    .PARAMETER Message
        The human-readable description of the problem. Forwarded to
        Add-ErrorHC.

    .PARAMETER Description
        Optional additional detail. Defaults to an empty string. Forwarded to
        Add-ErrorHC.

    .PARAMETER SystemErrors
        A [ref] to the caller's error accumulator (a collection supporting
        .Add()). Forwarded to Add-ErrorHC.

    .EXAMPLE
        $errors = [System.Collections.Generic.List[object]]::new()
        Add-RuntimeErrorHC -Type 'FatalError' -Name 'Missing setting' -Message 'LogFolder is not configured.' -SystemErrors ([ref]$errors)

        Appends a 'RuntimeSettings'-category error to $errors.

    .OUTPUTS
        None. Appends to the referenced collection and returns nothing.

    .NOTES
        Category is fixed to 'RuntimeSettings' and cannot be overridden through
        this function.

    .LINK
        Add-ErrorHC
    #>
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
    <#
    .SYNOPSIS
        Add a 'JsonSchema'-category error to the system-error accumulator.

    .DESCRIPTION
        Thin wrapper around Add-ErrorHC that fixes Category to 'JsonSchema'. All
        other parameters are forwarded unchanged. See Add-ErrorHC for the full
        description of the record created and how SystemErrors is used.

    .PARAMETER Type
        The error severity or kind (for example 'FatalError'). Forwarded to
        Add-ErrorHC.

    .PARAMETER Name
        A short title for the error. Forwarded to Add-ErrorHC.

    .PARAMETER Message
        The human-readable description of the problem. Forwarded to
        Add-ErrorHC.

    .PARAMETER Description
        Optional additional detail. Defaults to an empty string. Forwarded to
        Add-ErrorHC.

    .PARAMETER SystemErrors
        A [ref] to the caller's error accumulator (a collection supporting
        .Add()). Forwarded to Add-ErrorHC.

    .EXAMPLE
        $errors = [System.Collections.Generic.List[object]]::new()
        Add-JsonSchemaErrorHC -Type 'FatalError' -Name 'Schema violation' -Message "Property 'Path' is required." -SystemErrors ([ref]$errors)

        Appends a 'JsonSchema'-category error to $errors.

    .OUTPUTS
        None. Appends to the referenced collection and returns nothing.

    .NOTES
        Category is fixed to 'JsonSchema' and cannot be overridden through this
        function.

    .LINK
        Add-ErrorHC
    #>
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
    <#
    .SYNOPSIS
        Resolve a string value that may be a literal or an 'ENV:'-prefixed
        reference to an environment variable.

    .DESCRIPTION
        Returns a string based on the form of Name:

        - Null, empty or whitespace: returns $null.
        - Starts with 'ENV:' (case-insensitive): the text after the prefix is
          trimmed and treated as an environment variable name. The variable's
          value is returned. If the variable does not exist, the function
          throws.
        - Anything else: Name is returned unchanged.

        This lets a configuration field hold either a literal value or a
        pointer to an environment variable, so secrets and machine-specific
        paths can be kept out of the configuration itself.

    .PARAMETER Name
        The value to resolve. Either a literal string, or 'ENV:' followed by an
        environment variable name (for example 'ENV:SMTP_PASSWORD'). The prefix
        match is case-insensitive and the variable name is trimmed before
        lookup.

    .EXAMPLE
        Get-StringValueHC -Name 'smtp.contoso.com'

        Returns 'smtp.contoso.com'. With no 'ENV:' prefix the value is returned
        as-is.

    .EXAMPLE
        $env:SMTP_SERVER = 'smtp.contoso.com'
        Get-StringValueHC -Name 'ENV:SMTP_SERVER'

        Returns 'smtp.contoso.com', read from the SMTP_SERVER environment
        variable.

    .EXAMPLE
        Get-StringValueHC -Name 'ENV:DOES_NOT_EXIST'

        Throws "Environment variable 'DOES_NOT_EXIST' not found.", because the
        referenced variable is not set.

    .EXAMPLE
        Get-StringValueHC -Name '   '

        Returns $null, because the input is whitespace.

    .OUTPUTS
        System.String
        The resolved value, or $null when Name is null/empty/whitespace.

    .NOTES
        - The 'ENV:' prefix match is case-insensitive (ordinal), so 'env:',
          'Env:' and 'ENV:' all trigger environment-variable resolution.
        - A missing environment variable is a terminating error, whereas a
          missing/blank Name simply yields $null. The two "no value" situations
          are handled differently by design.
        - An environment variable that exists but is empty returns its empty
          value; it is not treated as "not found".
        - Only the leading 'ENV:' (4 characters) is stripped. A value like
          'ENV: NAME' resolves the variable ' NAME', which is then trimmed to
          'NAME'; but a literal value that genuinely begins with 'ENV:' cannot
          be returned as-is, since it will always be interpreted as a reference.
    #>
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

        [Parameter(Position = 1)]
        [AllowEmptyString()]
        [string]$Default = ''
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
        [Parameter(Mandatory)][string]$JsonFileName
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
            $JsonFileName
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

function Remove-BlankValueHC {
    <#
    .SYNOPSIS
        Returns a copy of a hashtable with entries whose value is $null or a
        blank/whitespace string removed.

    .DESCRIPTION
        Intended for cleaning a splatting hashtable so that missing optional
        values fall back to a command's parameter defaults instead of being
        passed as '' — which would be rejected by a [ValidateSet] at binding
        time (the default cannot apply once a value, even an empty one, is
        explicitly supplied).

        Only $null and blank/whitespace *strings* are removed. Other values —
        numbers, booleans, and arrays (including empty arrays) — are preserved,
        so collection parameters such as To or Attachments are never dropped.

        The input hashtable is not modified; a shallow clone is returned.

    .PARAMETER Hashtable
        The hashtable to clean.

    .EXAMPLE
        $mailParams = Remove-BlankValueHC -Hashtable $mailParams
        Send-MailKitMessageHC @mailParams
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Hashtable
    )

    $clean = $Hashtable.Clone()

    foreach ($key in @($clean.Keys)) {
        $value = $clean[$key]

        if (
            $null -eq $value -or
            ($value -is [string] -and [string]::IsNullOrWhiteSpace($value))
        ) {
            $clean.Remove($key)
        }
    }

    $clean
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
        Calculates the total errors and warnings across all matrix files and
        system-level errors.
    .DESCRIPTION
        Walks $Context.FileResults — the same data shape used by
        Build-MatrixFileCardHC — so the global "Detected issues" pills in the
        email always match the per-file cards.

        Buckets:
            File        — fileResult.Check                       (workbook-level)
            FormData    — fileResult.Sheets.FormData.Check       (FormData sheet)
            Permissions — fileResult.Sheets.Permissions.Check    (Permissions sheet)
            Settings    — fileResult.Matrices[].Check            (per-matrix rows)
            System      — $SystemErrors.Value                    (script-level)

        TotalErrors   = sum of all 'FatalError'-typed checks across every bucket,
                        including system errors.
        TotalWarnings = sum of all 'Warning'-typed checks across every bucket,
                        including system errors.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Context,
        [Parameter(Mandatory)][ref]$SystemErrors
    )

    $Context.Counter = New-CounterObjectHC

    $countByType = {
        param($Checks, [string]$Type)
        if (-not $Checks) { return 0 }
        return @($Checks | Where-Object { $_.Type -eq $Type }).Count
    }

    if ($Context.FileResults) {
        foreach ($fileResult in $Context.FileResults) {
            $Context.Counter.File.Errors += & $countByType $fileResult.Check 'FatalError'
            $Context.Counter.File.Warnings += & $countByType $fileResult.Check 'Warning'

            $Context.Counter.FormData.Errors += & $countByType $fileResult.Sheets.FormData.Check 'FatalError'
            $Context.Counter.FormData.Warnings += & $countByType $fileResult.Sheets.FormData.Check 'Warning'

            $Context.Counter.Permissions.Errors += & $countByType $fileResult.Sheets.Permissions.Check 'FatalError'
            $Context.Counter.Permissions.Warnings += & $countByType $fileResult.Sheets.Permissions.Check 'Warning'

            if ($fileResult.Matrices) {
                foreach ($m in $fileResult.Matrices) {
                    $Context.Counter.Settings.Errors += & $countByType $m.Check 'FatalError'
                    $Context.Counter.Settings.Warnings += & $countByType $m.Check 'Warning'
                }
            }
        }
    }

    $systemErrCount = & $countByType $SystemErrors.Value 'FatalError'
    $systemWarnCount = & $countByType $SystemErrors.Value 'Warning'

    $Context.Counter.TotalErrors =
    $Context.Counter.File.Errors +
    $Context.Counter.FormData.Errors +
    $Context.Counter.Permissions.Errors +
    $Context.Counter.Settings.Errors +
    $systemErrCount

    $Context.Counter.TotalWarnings =
    $Context.Counter.File.Warnings +
    $Context.Counter.FormData.Warnings +
    $Context.Counter.Permissions.Warnings +
    $Context.Counter.Settings.Warnings +
    $systemWarnCount

    return $Context.Counter
}