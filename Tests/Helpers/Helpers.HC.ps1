function Assert-LogContainsSystemErrorHC {
    <#
        .SYNOPSIS
            Asserts that the system error log contains a message matching a pattern.

        .DESCRIPTION
            - Locates the system error log file in the given folder
            - Reads and converts JSON
            - Searches all error message entries for a wildcard pattern
            - Throws a Pester-friendly error if no match is found

        .PARAMETER LogFolderPath
            Folder where system error log files are generated.

        .PARAMETER Pattern
            Wildcard search pattern (e.g. "*Property 'Matrix.FolderPath' not found*").
    #>

    param(
        [Parameter(Mandatory)]
        [string]$LogFolderPath,

        [Parameter(Mandatory)]
        [string]$Pattern,

        [Parameter(Mandatory = $false)]
        [string]$FileName = 'SystemErrors.json'
    )

    #
    # 1. Locate system error log file
    #
    $logFiles = Get-ChildItem -Path $LogFolderPath -File -Filter $FileName -Recurse

    if ($logFiles.Count -eq 0) {
        throw "Assert-LogContainsSystemErrorHC: No system error log file found in '$LogFolderPath'."
    }
    elseif ($logFiles.Count -gt 1) {
        throw "Assert-LogContainsSystemErrorHC: Multiple system error log files found in '$LogFolderPath'."
    }

    #
    # 2. Parse the JSON log file
    #
    $json = Get-Content -LiteralPath $logFiles.FullName -Raw | ConvertFrom-Json

    if (-not $json) {
        throw "Assert-LogContainsSystemErrorHC: Log file '$($logFiles.FullName)' is empty or unreadable."
    }

    #
    # 3. Extract error messages from all objects
    #
    $messages = $json.Message

    if (-not $messages) {
        throw "Assert-LogContainsSystemErrorHC: No 'Message' fields found in the log file."
    }

    #
    # 4. Look for a wildcard match
    #
    $match = $messages | Where-Object { $_ -like $Pattern }

    if (-not $match) {
        throw "Assert-LogContainsSystemErrorHC: No log entry matching pattern '$Pattern' found.`nMessages:`n$($messages -join "`n")"
    }
}

function Assert-HtmlLogContainsPatternHC {
    <#
    .SYNOPSIS
        Scans all HTML files in the latest log run folder to verify an expected error or pattern was logged.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$LogFolderPath,
        [Parameter(Mandatory)]
        [string]$Pattern,
        [string]$FileMatch,
        [switch]$Not
    )

    $latestRunFolder = Get-LatestLogFolderHC -Root $LogFolderPath
    $latestRunFolder | Should -Not -BeNullOrEmpty -Because "A log folder should have been created in '$LogFolderPath'"

    $htmlFiles = Get-ChildItem -Path $latestRunFolder -Recurse -Filter '*.html'
    
    if (-not [string]::IsNullOrWhiteSpace($FileMatch)) {
        $htmlFiles = $htmlFiles | Where-Object { $_.FullName -match $FileMatch }
    }

    $htmlFiles.Count | Should -BeGreaterThan 0 -Because 'At least one HTML log file should have been generated'

    $foundMatch = $false

    foreach ($file in $htmlFiles) {
        $rawHtml = Get-Content -LiteralPath $file.FullName -Raw
        
        $decodedHtml = [System.Net.WebUtility]::HtmlDecode($rawHtml)

        if ($decodedHtml -like $Pattern) {
            $foundMatch = $true
            break
        }
    }

    if ($Not) {
        $foundMatch | Should -Be $false -Because "The HTML logs must NOT contain the pattern: $Pattern"
    }
    else {
        $foundMatch | Should -Be $true -Because "The HTML logs must contain the expected pattern: $Pattern"
    }
}

function Clear-TestLogFoldersHC {
    <#
        .SYNOPSIS
            Removes all log folders created by the script during prior test runs.

        .DESCRIPTION
            This helper deletes both:
              1. The configured log folder defined in test input JSON
              2. The TEMP fallback folder used when Settings is missing/invalid

            This ensures each test starts with a clean log environment and that
            Assert-LogContainsSystemErrorHC finds exactly one "System errors log" file.
    #>

    param(
        [Parameter(Mandatory)]
        [string]$ConfiguredLogFolder
    )

    # --- Normal configured log folder ---
    if ($ConfiguredLogFolder -and (Test-Path -LiteralPath $ConfiguredLogFolder)) {
        Get-ChildItem -LiteralPath $ConfiguredLogFolder -Recurse -Force -ErrorAction SilentlyContinue |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }

    # --- Fallback TEMP log folder ---
    $fallback = Join-Path $env:TEMP 'PermissionMatrixLogs'

    if (Test-Path -LiteralPath $fallback) {
        Remove-Item -LiteralPath $fallback -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Compare-HashTableHC {
    param (
        [Parameter(Mandatory)]
        [hashtable]$ReferenceObject,
        [Parameter(Mandatory)]
        [hashtable]$DifferenceObject
    )

    (
        $ReferenceObject.GetEnumerator() |
        Sort-Object Key |
        ConvertTo-Json
    ) |
    Should -BeExactly (
        $DifferenceObject.GetEnumerator() |
        Sort-Object Key |
        ConvertTo-Json
    )
}

function Copy-ObjectHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$InputObject
    )

    $json = $InputObject | ConvertTo-Json -Depth 100
    return ($json | ConvertFrom-Json -AsHashtable)
}

function Get-FallbackLogFolderHC {
    return Join-Path $env:TEMP 'PermissionMatrixLogs'
}

function Get-LatestLogFolderHC {
    param(
        [Parameter(Mandatory)]
        [string]$Root
    )

    if (-not (Test-Path $Root)) { return $null }

    return Get-ChildItem -LiteralPath $Root -Directory |
    Sort-Object Name -Descending |
    Select-Object -First 1
}

function Save-TestJson {
    param(
        [Parameter(Mandatory)]
        [object]$InputObject,
        [Parameter(Mandatory)]
        [string]$JsonFile
    )

    $InputObject |
    ConvertTo-Json -Depth 20 |
    Set-Content -LiteralPath $JsonFile -Encoding UTF8
}

function Test-GetLogFileDataHC {
    param(
        [Parameter(Mandatory)]
        [string]$LogFolderPath,
        [string]$FileNameRegex = 'System errors log.json'
    )

    $files = @(Get-ChildItem -Path $LogFolderPath -File -Filter $FileNameRegex -Recurse)

    if ($files.Count -eq 0) {
        throw "No log file found in '$LogFolderPath' matching '$FileNameRegex'"
    }
    elseif ($files.Count -gt 1) {
        throw "Multiple log files found in '$LogFolderPath' matching '$FileNameRegex'"
    }

    return (Get-Content -LiteralPath $files[0].FullName | ConvertFrom-Json)
}

function Test-GetDatedLogFolderPathHC {
    param(
        [Parameter(Mandatory)]
        [string]$LogFolderRoot
    )

    return Get-ChildItem -Path $LogFolderRoot -Directory -Filter (
        '{0:00}_{1:00}_{2:00}* (*)' -f
        (Get-Date).Year,
        (Get-Date).Month,
        (Get-Date).Day
    )
}

function Test-GetMatrixLogFolderPathHC {
    param(
        [Parameter(Mandatory)]
        [string]$DatedLogFolder
    )

    return Join-Path $DatedLogFolder 'Matrix'
}

function Send-MailKitMessageHC {
    param()
}

function Set-NestedPropertyHC {
    param(
        [Parameter(Mandatory)]
        [object]$Object,

        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        $Value
    )

    # Split "Matrix.FolderPath" into ["Matrix","FolderPath"]
    $parts = $Path -split '\.'

    # Navigate down until the second-to-last property
    $cursor = $Object
    for ($i = 0; $i -lt $parts.Count - 1; $i++) {
        $prop = $parts[$i]

        if ($null -eq $cursor.$prop) {
            throw "Property '$prop' in path '$Path' does not exist."
        }

        $cursor = $cursor.$prop
    }

    # Set final property
    $final = $parts[-1]
    $cursor.$final = $Value
}
