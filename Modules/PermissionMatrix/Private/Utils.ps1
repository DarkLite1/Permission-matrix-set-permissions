function Get-StringValueHC {
    <#
    .SYNOPSIS
        Resolve a string value that may be a literal or an 'ENV:'-prefixed
        reference to an environment variable.

    .DESCRIPTION
        Lets a configuration field hold either a literal value or a pointer to
        an environment variable, so secrets and machine-specific paths stay out
        of the configuration file itself.

        - Null, empty or whitespace: returns $null.
        - Starts with 'ENV:': the trimmed remainder is an environment variable
          name. Its value is returned, or the function throws if the variable
          does not exist. 'ENV:' with nothing usable after it throws its own
          dedicated error rather than looking up an empty name.
        - Anything else: returned unchanged.

    .NOTES
        - A blank Name yields $null, but a MISSING environment variable is a
          terminating error. The two "no value" situations differ by design.
        - A variable that exists but is empty returns its empty value; it is
          not treated as "not found".
        - The lookup is a literal name match, so '*' or '\' in the name are not
          interpreted as wildcards or provider paths.
        - The 'ENV:' prefix match is case-insensitive, and only the leading 4
          characters are stripped. A literal value that genuinely begins with
          'ENV:' can therefore never be returned as-is.

    .EXAMPLE
        $env:SMTP_SERVER = 'smtp.contoso.com'
        Get-StringValueHC -Name 'ENV:SMTP_SERVER'

        Returns 'smtp.contoso.com'.
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
        Display/fallback companion to Get-StringValueHC, for rendering a
        placeholder such as 'Unknown' or 'N/A' when a string is missing.

    .NOTES
        - Does NOT resolve 'ENV:' prefixes. Use Get-StringValueHC for config
          strings that may reference environment variables.
        - Value is coerced to [string] at binding, so $null becomes '' and the
          original type of a non-string argument is not preserved.

    .EXAMPLE
        [System.Net.WebUtility]::HtmlEncode(
            (Get-StringOrDefaultHC $excel.LastModifiedBy 'Unknown')
        )

        Guarantees a non-blank string before HTML-encoding.

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

function Remove-BlankValueHC {
    <#
    .SYNOPSIS
        Returns a copy of a hashtable with entries whose value is $null or a
        blank/whitespace string removed.

    .DESCRIPTION
        Cleans a splatting hashtable so missing optional values fall back to a
        command's parameter defaults instead of being passed as '' — which a
        [ValidateSet] rejects at binding time, because the default can no longer
        apply once a value, even an empty one, is explicitly supplied.

    .NOTES
        Only $null and blank/whitespace STRINGS are removed. Numbers, booleans
        and arrays (including empty arrays) are preserved, so collection
        parameters such as To or Attachments are never dropped. The input is not
        modified; a shallow clone is returned.

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
        Builds one PSCustomObject with the fields DateTime, Type, Name,
        Description, Value and Category. DateTime is stamped at creation.

        This is the return-a-record counterpart to Add-ErrorHC, which instead
        appends to a [ref] accumulator. The caller decides where to store the
        result. The field shape is parallel except this record's free-form Value
        (any type) replaces the error record's string Message.

    .NOTES
        Unlike Add-ErrorHC, Type is free text and is NOT validated against a
        fixed set.

    .EXAMPLE
        $checks = [System.Collections.Generic.List[object]]::new()
        $checks.Add(
            (New-ValidationCheckHC -Type 'Warning' -Name 'Empty sheet' -Description 'No data rows found.')
        )

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