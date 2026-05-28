function Cleanup-OldLogsHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$LogFolder,
        [int]$RetentionDays,
        [Parameter(Mandatory)] [ref]$SystemErrors
    )

    # Disabled or folder missing → nothing to do
    if ($RetentionDays -le 0 -or -not $LogFolder) { return }

    try {
        if (-not (Test-Path -LiteralPath $LogFolder -PathType Container)) { return }

        $cutoff = (Get-Date).AddDays(-$RetentionDays)

        # --- 1. Delete old files ---
        Get-ChildItem -LiteralPath $LogFolder -Recurse -File -ErrorAction Stop |
        Where-Object { $_.CreationTime -lt $cutoff } |
        ForEach-Object {
            try {
                Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
            }
            catch {
                Add-ErrorHC `
                    -Type 'Warning' `
                    -Name 'Log cleanup failed' `
                    -Message "Failed to delete log file '$($_.FullName)': $_" `
                    -Category 'Logging' `
                    -SystemErrors $SystemErrors
            }
        }

        # --- 2. Empty folder cleanup (bottom-up) ---
        Get-ChildItem -LiteralPath $LogFolder -Recurse -Directory -ErrorAction Stop |
        Sort-Object FullName -Descending |
        ForEach-Object {
            if (-not $_.GetFileSystemInfos().Count) {
                try {
                    Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
                }
                catch {
                    Add-ErrorHC `
                        -Type 'Warning' `
                        -Name 'Log cleanup failed' `
                        -Message "Failed to remove empty folder '$($_.FullName)': $_" `
                        -Category 'Logging' `
                        -SystemErrors $SystemErrors
                }
            }
        }
    }
    catch {
        Add-ErrorHC `
            -Type 'Warning' `
            -Name 'Log cleanup failed' `
            -Message "General log cleanup failure: $_" `
            -Category 'Logging' `
            -SystemErrors $SystemErrors
    }
}

function Out-LogFileHC {
    <#
        Generic exporter for CSV / JSON / TXT / XLSX log files.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [PSCustomObject[]]$DataToExport,
        [Parameter(Mandatory)] [String]$PartialPath,
        [Parameter(Mandatory)] [String[]]$FileExtensions,
        [hashtable]$ExcelFile = @{
            SheetName = 'Overview'
            TableName = 'Overview'
            CellStyle = $null
        },
        [Switch]$Append
    )

    $allPaths = @()

    foreach ($ext in ($FileExtensions | Sort-Object -Unique)) {

        $logFilePath = "$PartialPath$ext"

        try {
            switch ($ext) {

                '.csv' {
                    $DataToExport |
                    Export-Csv -LiteralPath $logFilePath -Delimiter ';' `
                        -Append:$Append -NoTypeInformation
                    break
                }

                '.json' {
                    $converted = foreach ($item in $DataToExport) {
                        foreach ($p in $item.PSObject.Properties) {
                            if ($p.Value -is [System.Management.Automation.ErrorRecord]) {
                                $item.$($p.Name) = $p.Value.Exception.Message
                            }
                        }
                        $item
                    }

                    if ($Append -and (Test-Path $logFilePath)) {
                        $existing = Get-Content -LiteralPath $logFilePath -Raw | ConvertFrom-Json
                        $converted = @($converted) + @($existing)
                    }

                    $converted |
                    ConvertTo-Json -Depth 7 |
                    Out-File -LiteralPath $logFilePath -Encoding utf8 -Force
                    break
                }

                '.txt' {
                    $DataToExport |
                    Format-List * |
                    Out-File -LiteralPath $logFilePath -Append:$Append
                    break
                }

                '.xlsx' {
                    if (-not $Append -and (Test-Path $logFilePath)) {
                        Remove-Item -LiteralPath $logFilePath -Force
                    }

                    $params = @{
                        Path          = $logFilePath
                        Append        = $true
                        AutoNameRange = $true
                        AutoSize      = $true
                        FreezeTopRow  = $true
                        WorksheetName = $ExcelFile.SheetName
                        TableName     = $ExcelFile.TableName
                    }

                    if ($ExcelFile.CellStyle) {
                        $params.CellStyleSB = $ExcelFile.CellStyle
                    }

                    $DataToExport | Export-Excel @params
                    break
                }

                default {
                    throw "Unsupported file extension '$ext'."
                }
            }

            $allPaths += $logFilePath
        }
        catch {
            Write-Warning "Failed to export log '$logFilePath': $_"
        }
    }

    return $allPaths
}

function Remove-FileHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$FilePath,
        [ref]$SystemErrors
    )

    try {
        if (-not (Test-Path -LiteralPath $FilePath -PathType Leaf)) { return }
        Remove-Item -LiteralPath $FilePath -Force -ErrorAction Stop
    }
    catch {
        if ($SystemErrors) {
            Add-ErrorHC `
                -Type 'Warning' `
                -Name 'Failed to remove file' `
                -Message "Failed to remove '$FilePath': $_" `
                -Category 'Logging' `
                -SystemErrors $SystemErrors
        }
        else {
            Write-Warning "Failed removing '$FilePath': $_"
        }
    }
}

function Write-EventLogSafe {
    <#
        Safe wrapper to write data into Windows EventLog.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$EventLogData,
        [Parameter(Mandatory)][string]$ScriptName,
        [Parameter(Mandatory)][object]$Settings,
        [Parameter(Mandatory)][ref]$SystemErrors
    )

    $maxLen = 31000

    try {
        $logName = Get-StringValueHC $Settings.SaveInEventLog.LogName
        if (-not ($Settings.SaveInEventLog.Save -and $logName)) { return }

        # Append SystemErrors as individual error entries
        foreach ($err in $SystemErrors.Value) {
            $EventLogData.Add(
                [PSCustomObject]@{
                    Message   = $err.Message
                    DateTime  = $err.DateTime
                    EntryType = 'Error'
                    EventID   = '2'
                }
            )
        }

        # Add “script ended”
        $EventLogData.Add(
            [PSCustomObject]@{
                Message   = 'Script ended'
                DateTime  = Get-Date
                EntryType = 'Information'
                EventID   = '199'
            }
        )

        # Truncate too-long messages
        foreach ($item in $EventLogData) {
            if ($item.Message.Length -gt $maxLen) {
                $item.Message =
                $item.Message.Substring(0, $maxLen) +
                '... [TRUNCATED DUE TO EVENT LOG SIZE LIMITS]'
            }
        }

        Write-EventsToEventLogHC `
            -Source $ScriptName `
            -LogName $logName `
            -Events $EventLogData
    }
    catch {
        Add-ErrorHC `
            -Type 'Warning' `
            -Name 'Failed to write to event log' `
            -Message "Failed writing to event log: $_" `
            -Category 'Logging' `
            -SystemErrors $SystemErrors
    }
}

function Write-EventsToEventLogHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][String]$Source,
        [Parameter(Mandatory)][String]$LogName,
        [PSCustomObject[]]$Events
    )

    try {
        if (-not [System.Diagnostics.EventLog]::SourceExists($Source)) {
            New-EventLog -LogName $LogName -Source $Source -EA Stop
        }

        foreach ($eventItem in $Events) {

            $params = @{
                LogName     = $LogName
                Source      = $Source
                EntryType   = $eventItem.EntryType
                EventID     = $eventItem.EventID
                Message     = ''
                ErrorAction = 'Stop'
            }

            if (-not $params.EntryType) { $params.EntryType = 'Information' }
            if (-not $params.EventID) { $params.EventID = 4 }

            foreach ($prop in $eventItem.PSObject.Properties |
                Where-Object { $_.Name -notin 'EntryType', 'EventID' }) {

                $params.Message += "`n- $($prop.Name): $($prop.Value)"
            }

            Write-EventLog @params
        }
    }
    catch {
        throw "Failed writing events to Windows Event Log: $_"
    }
}

function Write-SystemErrorLogHC {
    <#
        Creates JSON log file of system errors and attaches to email params.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$SystemErrors,
        [Parameter(Mandatory)][string]$LogFolder,
        [Parameter(Mandatory)][ref]$MailParams,
        [datetime]$ScriptStartTime = (Get-Date),
        [string]$JsonFileName = 'MatrixConfig' 
    )

    if ($SystemErrors.Count -eq 0) { return }
    if (-not (Test-Path -LiteralPath $LogFolder -PathType Container)) { return }

    $datedFolder = Get-DatedLogFolderPathHC `
        -LogFolder $LogFolder `
        -ScriptStartTime $ScriptStartTime `
        -JsonFileName $JsonFileName

    $partial = Join-Path $datedFolder 'SystemErrors'

    $attachments = Out-LogFileHC `
        -DataToExport $SystemErrors `
        -PartialPath $partial `
        -FileExtensions '.json' `
        -ErrorAction Ignore

    if ($attachments) {
        if (-not $MailParams.Value.ContainsKey('Attachments')) {
            $MailParams.Value['Attachments'] = @()
        }
        $MailParams.Value['Attachments'] += $attachments
    }
}