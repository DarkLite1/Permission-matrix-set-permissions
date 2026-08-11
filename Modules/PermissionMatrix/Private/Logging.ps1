function Remove-OldLogsHC {
    <#
    .SYNOPSIS
        Purges old log files and the directories left empty behind them.

    .DESCRIPTION
        Deletes files older than the retention threshold, then walks the
        directory tree bottom-up (sorted descending) removing any subdirectory
        that is now empty.

    .NOTES
        - Age is measured on CreationTime, NOT LastWriteTime. A file that is
          still being appended to is deleted once it is old enough.
        - RetentionDays of 0 or less disables the cleanup entirely.
        - Deletion failures (a log open in another process, access denied) are
          appended to SystemErrors as WARNINGS and never throw, so cleanup
          problems cannot crash the orchestrator.

    .EXAMPLE
        $sysErrors = [System.Collections.Generic.List[pscustomobject]]::new()

        Remove-OldLogsHC `
            -LogFolder 'C:\MatrixLogs' `
            -RetentionDays 30 `
            -SystemErrors ([ref]$sysErrors)
    #>
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

function Write-CheckDetailJsonHC {
    <#
    .SYNOPSIS
        Writes a single check object to its detailed JSON log file.

    .DESCRIPTION
        Stamps the check with 'JsonFileName' and 'JsonFilePath', then, when the
        check carries a 'Value', serializes a copy (excluding those two link
        properties) to disk as JSON.

    .NOTES
        - The passed-in check object is MUTATED in place.
        - ErrorRecord/Exception values are rendered to their string form first,
          to avoid serialization depth failures.
        - When the check has no 'Value', or serialization fails, both link
          properties are reset to $null so downstream reporting never links to a
          file that was not written. A failure is also appended to the check's
          'Description'.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Check,
        [Parameter(Mandatory)] [string]$JsonFileName,
        [Parameter(Mandatory)] [string]$LogFolder
    )

    $Check | Add-Member -NotePropertyMembers @{
        JsonFileName = $JsonFileName
        JsonFilePath = Join-Path -Path $LogFolder -ChildPath $JsonFileName
    } -Force

    if (-not $Check.Value) {
        $Check.JsonFileName = $null
        $Check.JsonFilePath = $null
        return
    }

    try {
        $cForJson = $Check | Select-Object -ExcludeProperty JsonFilePath, JsonFileName

        if (
            $cForJson.Value -is [System.Management.Automation.ErrorRecord] -or
            $cForJson.Value -is [Exception]
        ) {
            $cForJson.Value = ($cForJson.Value | Out-String).Trim()
        }

        $cForJson | ConvertTo-Json -Depth 10 |
        Out-File -FilePath $Check.JsonFilePath -Encoding UTF8 -Force
    }
    catch {
        $Check.Description += "[Detailed JSON log failed to generate: $($_)]"
        $Check.JsonFileName = $null
        $Check.JsonFilePath = $null
    }
}

function Out-LogFileHC {
    <#
    .SYNOPSIS
        Writes PowerShell objects to several file formats at once, from one
        shared base path.

    .DESCRIPTION
        Exports DataToExport to each requested extension ('.csv', '.json',
        '.txt', '.xlsx'), appending the extension to PartialPath. Extensions are
        de-duplicated and sorted before use. Returns the paths actually written.

        Format-specific handling:
        - JSON: ErrorRecord values are flattened to their message text so
          serialization cannot fail on depth. The caller's DataToExport is never
          mutated: a fresh object is built per item. With -Append the existing
          file is read and merged, rather than concatenating raw text.
        - XLSX: built through ImportExcel with frozen headers and auto-sized
          columns. Without -Append an existing file is deleted first, because
          Export-Excel is always called in append mode.

    .PARAMETER ExcelFile
        Formatting rules for '.xlsx'. Keys: 'SheetName', 'TableName',
        'CellStyle'.

    .NOTES
        A format that fails is reported with Write-Warning and its path is
        omitted from the return value; the remaining formats still export. A
        caller that ignores the returned paths will not notice a partial
        failure. An unsupported extension throws inside the loop and is caught
        by that same handler.

    .OUTPUTS
        System.String[]
        Paths of the log files successfully generated or updated.

    .EXAMPLE
        $exportedPaths = Out-LogFileHC `
            -DataToExport $data `
            -PartialPath 'C:\Logs\DailyReport' `
            -FileExtensions @('.csv', '.json', '.xlsx')
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
                    # Build a new object per item so the caller's $DataToExport
                    # is never mutated. ErrorRecord values are rendered to their
                    # message text; all other values are copied as-is.
                    $converted = foreach ($item in $DataToExport) {
                        $clone = [ordered]@{}
                        foreach ($p in $item.PSObject.Properties) {
                            $clone[$p.Name] = if ($p.Value -is [System.Management.Automation.ErrorRecord]) {
                                $p.Value.Exception.Message
                            }
                            else {
                                $p.Value
                            }
                        }
                        [PSCustomObject]$clone
                    }

                    if ($Append -and (Test-Path $logFilePath)) {
                        $existing = Get-Content -LiteralPath $logFilePath -Raw | ConvertFrom-Json
                        $converted = @($existing) + @($converted)
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
    <#
    .SYNOPSIS
        Deletes a file, downgrading locking/permission failures to a warning.

    .DESCRIPTION
        Removes FilePath when it exists. A missing file is a silent no-op.

    .NOTES
        Never throws. On failure the warning is routed to SystemErrors when that
        reference is supplied, and to the PowerShell warning stream when it is
        not.

    .EXAMPLE
        $sysErrors = [System.Collections.Generic.List[pscustomobject]]::new()
        Remove-FileHC `
            -FilePath 'C:\Temp\OldLog.txt' `
            -SystemErrors ([ref]$sysErrors)
    #>
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

function Write-EventLogSafeHC {
    <#
    .SYNOPSIS
        Formats and writes aggregated execution data and system errors to the
        Windows Event Log, without ever throwing.

    .DESCRIPTION
        Returns immediately unless 'SaveInEventLog.Save' is set and a LogName is
        configured. Otherwise it adds every entry in SystemErrors as its own
        Error event (EventID 2), appends a 'Script ended' Information event
        (EventID 199), then hands the batch to Write-EventsToEventLogHC.

    .NOTES
        - Messages longer than 31,000 characters are truncated with a marker.
          The Event Log API throws on oversized entries, which a large stack
          trace or data dump would otherwise trigger.
        - Failures (no permission to create the source or write the log) are
          appended to SystemErrors as a warning. Contrast
          Write-EventsToEventLogHC, which throws.

    .EXAMPLE
        Write-EventLogSafeHC `
            -EventLogData $eventData `
            -ScriptName 'Permission Matrix' `
            -Settings $Context.Config.Settings `
            -SystemErrors ([ref]$sysErrors)
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
    <#
    .SYNOPSIS
        Writes an array of custom objects to the Windows Event Log, flattening
        their properties into the event message.

    .DESCRIPTION
        Registers the Event Source when missing, then writes one event per
        object. 'EntryType' and 'EventID' map to the matching Event Log fields;
        every other property is flattened into a bulleted message body. Missing
        values default to 'Information' and EventID 4.

    .NOTES
        - Creating a new Event Source requires Administrator privileges.
        - Unlike its wrapper Write-EventLogSafeHC, this function THROWS on
          failure and passes the exception back up the chain.

    .EXAMPLE
        $events = @(
            [pscustomobject]@{
                EntryType = 'Warning'
                EventID   = 99
                Action    = 'Cleanup'
                Status    = 'Folder locked by another process'
            }
        )

        Write-EventsToEventLogHC `
            -Source 'MyScript' `
            -LogName 'Application' `
            -Events $events
    #>
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
    .SYNOPSIS
        Exports system errors to a JSON log file and attaches it to the outgoing
        email parameters.

    .DESCRIPTION
        Serializes SystemErrors to 'SystemErrors.json' in the dated log folder,
        then adds that path to the 'Attachments' key of the referenced
        MailParams hashtable, creating the key when absent. Administrators then
        receive the raw error data alongside the HTML summary.

    .PARAMETER MailParams
        A [ref] to the SMTP splatting hashtable. MUTATED in place.

    .NOTES
        Returns without doing anything when there are no system errors or when
        the log folder does not exist.

    .EXAMPLE
        $mailSplat = @{ To = 'admin@domain.com'; Subject = 'Execution Report' }

        Write-SystemErrorLogHC `
            -SystemErrors $sysErrors `
            -LogFolder 'C:\MatrixLogs' `
            -MailParams ([ref]$mailSplat)
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