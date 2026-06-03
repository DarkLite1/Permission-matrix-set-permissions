
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

function ConvertTo-StructuredObjectHC {
    <#
    .SYNOPSIS
        Normalize mixed pipeline input into structured records, wrapping strings
        and unknown objects and passing structured objects through.

    .DESCRIPTION
        Takes a stream of arbitrary objects and emits a structured record for
        each, so a mixed pipeline (strings, hashtables, custom objects, other
        types) becomes a uniform sequence of objects downstream code can handle
        consistently.

        Each input item is classified as follows:

        - $null: skipped, producing no output.
        - [string]: wrapped via New-ValidationCheckHC with Type 'Information'
          and Name 'Message', the string becoming the record's Description.
        - [hashtable] or [pscustomobject]: passed through unchanged, on the
          assumption it is already a structured record.
        - Anything else: stringified and wrapped via New-ValidationCheckHC with
          Type 'Information' and Name 'UnknownObject', the string form becoming
          the record's Description.

        The function processes pipeline input one item at a time and also
        iterates the items of any array passed as a single argument.

    .PARAMETER InputObject
        The object(s) to normalize. Accepts pipeline input. Each item is
        classified and emitted individually; $null items are dropped. Mandatory.

    .EXAMPLE
        'something happened' | ConvertTo-StructuredObjectHC

        Emits a validation-check record: Type 'Information', Name 'Message',
        Description 'something happened'.

    .EXAMPLE
        @(
            'a message',
            [pscustomobject]@{ Type = 'Warning'; Name = 'X' },
            42,
            $null
        ) | ConvertTo-StructuredObjectHC

        Emits three records: the string is wrapped as a 'Message', the
        PSCustomObject passes through unchanged, 42 is wrapped as an
        'UnknownObject' with Description '42', and the $null is skipped.

    .EXAMPLE
        Some-Step | ConvertTo-StructuredObjectHC | Where-Object Type -eq 'Information'

        Normalizes whatever Some-Step emits (free-form strings, ready-made
        records, or other values) so the downstream filter can rely on a
        consistent record shape.

    .OUTPUTS
        System.Management.Automation.PSCustomObject
        For strings and unrecognized types, a record from New-ValidationCheckHC.
        For hashtables and PSCustomObjects, the original object unchanged. No
        output is produced for $null items.

    .NOTES
        - $null items are silently dropped.
        - Strings and unknown types are wrapped with Type 'Information'; the
          difference is the Name ('Message' vs 'UnknownObject'). Note an unknown
          object is recorded as 'Information', not as a warning, even though it
          was an unexpected type.
        - Hashtables are passed through as-is and are not converted to
          PSCustomObjects or validated; downstream code receiving a [hashtable]
          alongside [pscustomobject] records should be ready for both shapes.
        - The wrapped records carry only Description (no Value); their Value and
          Category fields are $null.

    .LINK
        New-ValidationCheckHC
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $true)] 
        $InputObject
    )

    process {
        foreach ($obj in $InputObject) {
            
            if ($null -eq $obj) { continue }

            if ($obj -is [string]) {
                New-ValidationCheckHC `
                    -Type 'Information' `
                    -Name 'Message' `
                    -Description $obj
                continue
            }

            if ($obj -is [hashtable] -or $obj -is [pscustomobject]) {
                $obj
                continue
            }

            New-ValidationCheckHC `
                -Type 'Information' `
                -Name 'UnknownObject' `
                -Description "$obj"
        }
    }
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