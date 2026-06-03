function Get-StringValueHC {
    <#
    .SYNOPSIS
        Resolve a string value that may be a literal or an 'ENV:'-prefixed
        reference to an environment variable.

    .DESCRIPTION
        Returns a string based on the form of Name:

        - Null, empty or whitespace: returns $null.
        - Starts with 'ENV:' (case-insensitive): the text after the prefix is
          trimmed and treated as an environment variable name. If nothing
          usable remains after the prefix, the function throws. Otherwise the
          variable's value is returned, or the function throws if the variable
          does not exist.
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
        Get-StringValueHC -Name 'ENV:'

        Throws "No environment variable name given after 'ENV:'.", because the
        prefix is present but no variable name follows it.

    .EXAMPLE
        Get-StringValueHC -Name '   '

        Returns $null, because the input is whitespace.

    .OUTPUTS
        System.String
        The resolved value, or $null when Name is null/empty/whitespace.

    .NOTES
        - The 'ENV:' prefix match is case-insensitive (ordinal), so 'env:',
          'Env:' and 'ENV:' all trigger environment-variable resolution.
        - 'ENV:' with nothing (or only whitespace) after it is a terminating
          error with a dedicated message, rather than an attempted lookup of an
          empty variable name.
        - A missing environment variable is a terminating error, whereas a
          missing/blank Name simply yields $null. The two "no value" situations
          are handled differently by design.
        - An environment variable that exists but is empty returns its empty
          value; it is not treated as "not found".
        - The variable lookup is a literal name match; characters such as '*'
          or '\' in the name are not interpreted as wildcards or provider
          paths.
        - Only the leading 'ENV:' (4 characters) is stripped. A literal value
          that genuinely begins with 'ENV:' cannot be returned as-is, since it
          will always be interpreted as a reference.
    #>

    [CmdletBinding()]
    param([String]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return $null
    }
    elseif ($Name.StartsWith('ENV:', [System.StringComparison]::OrdinalIgnoreCase)) {
        $envVariableName = $Name.Substring(4).Trim()

        # Guard against 'ENV:' with no usable variable name after the prefix,
        # so the error names the problem instead of reporting an empty variable.
        if ([string]::IsNullOrWhiteSpace($envVariableName)) {
            throw "No environment variable name given after 'ENV:'."
        }

        # Plain literal lookup: no Env-provider path parsing, so characters
        # like '*' or '\' in the name are matched as-is.
        $envStringValue = [System.Environment]::GetEnvironmentVariable($envVariableName)

        # Explicit $null check (not truthiness) so an existing-but-empty
        # variable returns '' rather than being reported as "not found".
        if ($null -ne $envStringValue) {
            return $envStringValue
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
        Return Default when Value is null, empty, or whitespace only; otherwise
        return Value unchanged.

    .DESCRIPTION
        Display/fallback companion to Get-StringValueHC. Useful for rendering a
        placeholder when a string input is missing or blank (for example
        'Unknown', 'N/A', or a sensible default like 'Permission Matrix' for a
        missing script name).

        Value is typed [string], so any argument is coerced to its string form
        at binding ($null becomes ''). The function returns Default when the
        resulting string is null, empty or whitespace only, and otherwise
        returns Value unchanged.

        Note: this function does NOT resolve 'ENV:' prefixes. Use
        Get-StringValueHC for config strings that may reference environment
        variables.

    .PARAMETER Value
        The string to check. A non-string argument is coerced to [string] at
        binding; $null becomes '' and is therefore treated as blank.

    .PARAMETER Default
        The fallback returned when Value is blank. Defaults to ''.

    .EXAMPLE
        Get-StringOrDefaultHC -Value $row.Name -Default 'Unknown'

        Returns $row.Name when it has content, or 'Unknown' when it is null,
        empty or whitespace.

    .EXAMPLE
        Get-StringOrDefaultHC $row.Name 'Unknown'

        Same result using positional arguments: Value first, Default second.

    .EXAMPLE
        [System.Net.WebUtility]::HtmlEncode(
            (Get-StringOrDefaultHC $excel.LastModifiedBy 'Unknown')
        )

        Guarantees a non-blank string before HTML-encoding, so a missing
        LastModifiedBy renders as 'Unknown' rather than an empty cell.

    .OUTPUTS
        System.String
        Value when it is non-blank, otherwise Default.

    .NOTES
        - "Blank" means [string]::IsNullOrWhiteSpace is true: $null, '' or
          whitespace only.
        - Value is coerced to [string] at binding, so the original type of a
          non-string argument is not preserved.
        - Does not resolve 'ENV:' references; use Get-StringValueHC for that.

    .LINK
        Get-StringValueHC
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Position = 0)]
        [AllowEmptyString()]
        [string]$Value,

        [Parameter(Position = 1)]
        [AllowEmptyString()]
        [string]$Default = ''
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $Default
    }

    return $Value
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

function New-ValidationCheckHC {
    <#
    .SYNOPSIS
        Create a structured validation-check record and return it.

    .DESCRIPTION
        Builds a single PSCustomObject describing a validation check or result
        and returns it. The DateTime field is stamped with Get-Date at creation
        time; the remaining fields are taken from the parameters.

        Unlike the Add-*ErrorHC family, which appends to a referenced
        accumulator, this function returns the record to the caller, who
        decides where to store or emit it. The record carries a free-form Value
        (any type) in place of the error family's string Message.

    .PARAMETER Type
        The kind or severity of the check (for example 'Info', 'Warning',
        'FatalError'). Free text; not validated against a fixed set. Mandatory.

    .PARAMETER Name
        A short title identifying the check. Mandatory.

    .PARAMETER Description
        Optional human-readable detail about the check. When omitted, the
        record's Description is $null.

    .PARAMETER Value
        Optional payload of any type — the data the check produced or examined
        (a count, a list, an object, and so on). When omitted, the record's
        Value is $null.

    .PARAMETER Category
        Optional grouping label (for example 'Matrix', 'Permissions'). When
        omitted, the record's Category is $null.

    .EXAMPLE
        New-ValidationCheckHC -Type 'Info' -Name 'Row count' -Value 42 -Category 'Matrix'

        Returns a record with Type 'Info', Name 'Row count', Value 42, Category
        'Matrix', a current DateTime, and a $null Description.

    .EXAMPLE
        $checks = [System.Collections.Generic.List[object]]::new()
        $checks.Add((New-ValidationCheckHC -Type 'Warning' -Name 'Empty sheet' -Description 'No data rows found.'))

        Creates a record and adds it to the caller's own collection. Because the
        function returns rather than accumulates, the caller controls storage.

    .OUTPUTS
        System.Management.Automation.PSCustomObject
        One record with the fields DateTime, Type, Name, Description, Value and
        Category.

    .NOTES
        - DateTime is captured with Get-Date at creation (local time).
        - Value accepts any type; the other fields are strings (or $null when
          their optional parameter is omitted).
        - This is the return-a-record counterpart to the Add-*ErrorHC functions,
          which instead append to a [ref] accumulator. The field shape is
          parallel except this record's Value replaces the error record's
          Message.

    .LINK
        Add-ErrorHC
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Type,
        [Parameter(Mandatory)][string]$Name,
        [Parameter()][string]$Description,
        [Parameter()][object]$Value,
        [Parameter()][string]$Category
    )

    return [pscustomobject]@{
        DateTime    = Get-Date
        Type        = $Type
        Name        = $Name
        Description = $Description
        Value       = $Value
        Category    = $Category
    }
}
