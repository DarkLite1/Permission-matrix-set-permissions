function Add-ErrorHC {
    <#
    .SYNOPSIS
        Append a structured error record to a system-error accumulator.

    .DESCRIPTION
        Records rather than throws: the record is appended to SystemErrors and
        the caller decides how to proceed.

    .PARAMETER SystemErrors
        A [ref] to a collection exposing .Add(), for example
        [System.Collections.Generic.List[object]]. An array created with @() is
        fixed-size and causes a terminating error. Prefer a generic List over an
        ArrayList: ArrayList.Add() returns the insertion index, which would leak
        onto the pipeline.

    .NOTES
        Type is restricted to 'FatalError'/'Warning' because callers decide
        whether to halt by testing Type -eq 'FatalError'. A typo such as 'Fatal'
        would never match and would silently downgrade the error to advisory.

    .EXAMPLE
        $errors = [System.Collections.Generic.List[object]]::new()
        Add-ErrorHC -Type 'FatalError' -Name 'Bad row' -Message 'Missing path.' -Category 'Matrix' -SystemErrors ([ref]$errors)
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('FatalError', 'Warning')]
        [string]$Type,
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

function Add-JsonSchemaErrorHC {
    <#
    .SYNOPSIS
        Add a 'JsonSchema'-category error to the system-error accumulator.

    .DESCRIPTION
        Wrapper around Add-ErrorHC that fixes Category to 'JsonSchema'. All
        other parameters are forwarded unchanged.

    .LINK
        Add-ErrorHC
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('FatalError', 'Warning')]
        [string]$Type,
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
        Remote scripts return a mix of free-form strings and ready-made check
        records. This turns that stream into one downstream code can treat
        uniformly:

        - $null              skipped, no output
        - [string]           wrapped as Type 'Information', Name 'Message'
        - [hashtable] /
          [pscustomobject]   passed through unchanged
        - anything else      stringified, wrapped as Name 'UnknownObject'

    .NOTES
        - Hashtables pass through as-is and are NOT converted to PSCustomObject,
          so consumers can receive both shapes.
        - An unrecognized type is recorded as 'Information', not as a warning,
          even though its type was unexpected.

    .EXAMPLE
        Some-Step | ConvertTo-StructuredObjectHC | Where-Object Type -eq 'Information'

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

function Test-FileHasFatalErrorHC {
    <#
    .SYNOPSIS
        Checks whether a matrix file is fatally broken for permission handling.

    .DESCRIPTION
        A file blocks permission application when it has a fatal file-level
        error OR a fatal error on its Permissions sheet. FormData sheet errors
        are deliberately ignored here: they only affect the ServiceNow export,
        not the NTFS permissions.

    .PARAMETER File
        A file result object exposing a 'Check' list and a
        'Sheets.Permissions.Check' list.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$File
    )

    return Test-ItemHasFatalErrorHC -CheckList (
        @($File.Check) + @($File.Sheets.Permissions.Check)
    )
}

function New-CounterObjectHC {
    <#
    .SYNOPSIS
        Initializes an empty counter object for tracking errors and warnings.
    #>
    [CmdletBinding()]
    param()

    return [PSCustomObject]@{
        TotalErrors    = 0
        TotalIncorrect = 0
        TotalWarnings  = 0
        TotalFixed     = 0
        FormData       = [PSCustomObject]@{ Errors = 0; Incorrect = 0; Warnings = 0; Fixed = 0 }
        Permissions    = [PSCustomObject]@{ Errors = 0; Incorrect = 0; Warnings = 0; Fixed = 0 }
        Settings       = [PSCustomObject]@{ Errors = 0; Incorrect = 0; Warnings = 0; Fixed = 0 }
        File           = [PSCustomObject]@{ Errors = 0; Incorrect = 0; Warnings = 0; Fixed = 0 }
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

        System errors ($SystemErrors.Value) count towards the totals but have no
        bucket of their own: New-CounterObjectHC creates only the four above.
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
            $Context.Counter.File.Incorrect += & $countByType $fileResult.Check 'Incorrect'
            $Context.Counter.File.Warnings += & $countByType $fileResult.Check 'Warning'
            $Context.Counter.File.Fixed += & $countByType $fileResult.Check 'Fixed'

            $Context.Counter.FormData.Errors += & $countByType $fileResult.Sheets.FormData.Check 'FatalError'
            $Context.Counter.FormData.Incorrect += & $countByType $fileResult.Sheets.FormData.Check 'Incorrect'
            $Context.Counter.FormData.Warnings += & $countByType $fileResult.Sheets.FormData.Check 'Warning'
            $Context.Counter.FormData.Fixed += & $countByType $fileResult.Sheets.FormData.Check 'Fixed'

            $Context.Counter.Permissions.Errors += & $countByType $fileResult.Sheets.Permissions.Check 'FatalError'
            $Context.Counter.Permissions.Incorrect += & $countByType $fileResult.Sheets.Permissions.Check 'Incorrect'
            $Context.Counter.Permissions.Warnings += & $countByType $fileResult.Sheets.Permissions.Check 'Warning'
            $Context.Counter.Permissions.Fixed += & $countByType $fileResult.Sheets.Permissions.Check 'Fixed'

            if ($fileResult.Matrices) {
                foreach ($m in $fileResult.Matrices) {
                    $Context.Counter.Settings.Errors += & $countByType $m.Check 'FatalError'
                    $Context.Counter.Settings.Incorrect += & $countByType $m.Check 'Incorrect'
                    $Context.Counter.Settings.Warnings += & $countByType $m.Check 'Warning'
                    $Context.Counter.Settings.Fixed += & $countByType $m.Check 'Fixed'
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

    <# System errors have no 'Incorrect' or 'Fixed' equivalent: Add-ErrorHC
    only ever records 'FatalError' or 'Warning'. #>
    $Context.Counter.TotalIncorrect =
    $Context.Counter.File.Incorrect +
    $Context.Counter.FormData.Incorrect +
    $Context.Counter.Permissions.Incorrect +
    $Context.Counter.Settings.Incorrect

    $Context.Counter.TotalFixed =
    $Context.Counter.File.Fixed +
    $Context.Counter.FormData.Fixed +
    $Context.Counter.Permissions.Fixed +
    $Context.Counter.Settings.Fixed

    return $Context.Counter
}