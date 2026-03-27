# Permission-matrix-set-permissions\Modules\Toolbox.PermissionMatrixHC\Private\Utils.ps1
function Add-MatrixErrorHC {
    throw 'replace with Add-ErrorByCategoryHC'
}
function Add-ErrorByCategoryHC {
    <#
        .SYNOPSIS
            Add a structured error object into the shared SystemErrors collection.

        .DESCRIPTION
            This helper creates a canonical error object and injects it directly into 
            the SystemErrors reference. It enforces a single consistent error format 
            across the entire PermissionMatrix module.

        .PARAMETER Type
            Error type: 'FatalError' or 'Warning'.

        .PARAMETER Name
            Short label identifying the error.

        .PARAMETER Message
            A readable explanation summarizing the error.

        .PARAMETER Description
            Optional deeper technical context for HTML reports and logs.

        .PARAMETER Category
            Logical subsystem: 
            'JsonSchema', 'RuntimeSettings', 'Matrix', 'Permissions', 'FormData', 'File', etc.

        .PARAMETER SystemErrors
            Reference to the master SystemErrors list shared across all orchestrator components.

        .EXAMPLE
            Add-ErrorByCategoryHC -Type FatalError -Name 'Missing Settings' `
                -Message 'Settings object missing from JSON.' `
                -Category 'RuntimeSettings' `
                -SystemErrors ([ref]$SystemErrors)
    #>

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
function Add-RuntimeErrorHC {
    param(
        [string]$Type,
        [string]$Name,
        [string]$Message,
        [string]$Description = '',
        [ref]$SystemErrors
    )
    Add-ErrorByCategoryHC `
        -Type $Type `
        -Name $Name `
        -Message $Message `
        -Description $Description `
        -Category 'RuntimeSettings' `
        -SystemErrors $SystemErrors
}
function Add-JsonSchemaErrorHC {
    param(
        [string]$Type,
        [string]$Name,
        [string]$Message,
        [string]$Description = '',
        [ref]$SystemErrors
    )
    Add-ErrorByCategoryHC `
        -Type $Type `
        -Name $Name `
        -Message $Message `
        -Description $Description `
        -Category 'JsonSchema' `
        -SystemErrors $SystemErrors
}


function Get-StringValueHC {
    <#
        .SYNOPSIS
            Retrieve a string from the environment variables or a regular string.

        .DESCRIPTION
            This function checks the 'Name' property. If the value starts with
            'ENV:', it attempts to retrieve the string value from the specified
            environment variable. Otherwise, it returns the value directly.

        .PARAMETER Name
            Either a string starting with 'ENV:'; a plain text string or NULL.

        .EXAMPLE
            Get-StringValueHC -Name 'ENV:passwordVariable'

            # Output: the environment variable value of $ENV:passwordVariable
            # or an error when the variable does not exist

        .EXAMPLE
            Get-StringValueHC -Name 'mySecretPassword'

            # Output: mySecretPassword

        .EXAMPLE
            Get-StringValueHC -Name ''

            # Output: NULL
        #>
    param (
        [String]$Name
    )

    if (-not $Name) {
        return $null
    }
    elseif (
        $Name.StartsWith('ENV:', [System.StringComparison]::OrdinalIgnoreCase)
    ) {
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
function Get-DatedLogFolderPathHC {
    try {
        $datedLogFolder = Join-Path -Path $LogFolder -ChildPath (
            '{0:00}_{1:00}_{2:00}_{3:00}{4:00}{5:00} ({6})' -f $scriptStartTime.Year, $scriptStartTime.Month,
            $scriptStartTime.Day,
            $scriptStartTime.Hour, $scriptStartTime.Minute, $scriptStartTime.Second, $jsonFileItem.BaseName
        )

        return (New-Item -ItemType 'Directory' -Path $datedLogFolder -Force -EA Stop).FullName
    }
    catch {
        return $LogFolder
    }
}
function Plural {
    param(
        [int]$Count,
        [string]$Word
    )
    return ($Count -eq 1) ? $Word : ($Word + 's')
}
function Test-HasFatalErrorsHC {
    param([ref]$SystemErrors)

    return $SystemErrors.Value.Type -contains 'FatalError'
}
