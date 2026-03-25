function Cleanup-OldLogs {
    param(
        [Parameter(Mandatory)]
        [string]$LogFolder,
        [int]$RetentionDays,
        [Parameter(Mandatory)]
        [ref]$SystemErrors
    )

    #
    # No cleanup if retention is disabled or folder missing
    #
    if ($RetentionDays -le 0 -or -not $LogFolder) {
        return
    }

    try {
        if (-not (Test-Path -LiteralPath $LogFolder -PathType Container)) {
            return
        }

        $cutoff = (Get-Date).AddDays(-$RetentionDays)

        #
        # 1. Delete old files
        #
        Get-ChildItem -LiteralPath $LogFolder -Recurse -File -ErrorAction Stop |
        Where-Object { $_.CreationTime -lt $cutoff } |
        ForEach-Object {
            try {
                Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
            }
            catch {
                $SystemErrors.Value.Add(
                    [PSCustomObject]@{
                        DateTime = Get-Date
                        Message  = "Log cleanup failed to delete file '$($_.FullName)': $_"
                    }
                )
            }
        }

        #
        # 2. Remove empty folders (bottom‑up)
        #
        Get-ChildItem -LiteralPath $LogFolder -Recurse -Directory -ErrorAction Stop |
        Sort-Object FullName -Descending | # ensures children deleted before parents
        ForEach-Object {
            if (-not $_.GetFileSystemInfos().Count) {
                try {
                    Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
                }
                catch {
                    $SystemErrors.Value.Add(
                        [PSCustomObject]@{
                            DateTime = Get-Date
                            Message  = "Log cleanup failed to remove empty folder '$($_.FullName)': $_"
                        }
                    )
                }
            }
        }
    }
    catch {
        $SystemErrors.Value.Add(
            [PSCustomObject]@{
                DateTime = Get-Date
                Message  = "Log cleanup failed: $_"
            }
        )
    }
}
function Remove-FileHC {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath,
        [Parameter()]
        [ref]$SystemErrors
    )

    try {
        # If file doesn't exist, nothing to do
        if (-not (Test-Path -LiteralPath $FilePath -PathType Leaf)) {
            return
        }

        # Remove file safely
        Remove-Item -LiteralPath $FilePath -Force -ErrorAction Stop
    }
    catch {
        # Only record errors if SystemErrors was passed in
        if ($SystemErrors) {
            $SystemErrors.Value.Add(
                [PSCustomObject]@{
                    DateTime = Get-Date
                    Message  = "Remove-FileHC failed for '$FilePath': $_"
                }
            )
        }
        else {
            Write-Warning "Remove-FileHC failed for '$FilePath': $_"
        }
    }
}
function Write-EventLogSafe {
    param(
        [Parameter(Mandatory)][object]$EventLogData,
        [Parameter(Mandatory)][string]$ScriptName,
        [Parameter(Mandatory)][object]$Settings,
        [Parameter(Mandatory)][ref]$SystemErrors
    )

    # Maximum message length allowed by Windows Event Log
    $maxMessageLength = 31000

    try {
        # Is event log writing enabled?
        $logName = Get-StringValueHC $Settings.SaveInEventLog.LogName

        if (-not ($Settings.SaveInEventLog.Save -and $logName)) {
            return
        }

        #
        # 1. Convert system errors into event entries
        #
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

        #
        # 2. Add "Script ended" marker
        #
        $EventLogData.Add(
            [PSCustomObject]@{
                Message   = 'Script ended'
                DateTime  = Get-Date
                EntryType = 'Information'
                EventID   = '199'
            }
        )

        #
        # 3. Truncate messages exceeding allowed length
        #
        foreach ($item in $EventLogData) {
            if ($item.Message -and $item.Message.Length -gt $maxMessageLength) {
                $item.Message = $item.Message.Substring(0, $maxMessageLength) +
                '... [TRUNCATED DUE TO EVENT LOG SIZE LIMITS]'
            }
        }

        #
        # 4. Write the events
        #
        Write-EventsToEventLogHC `
            -Source $ScriptName `
            -LogName $logName `
            -Events $EventLogData
    }
    catch {
        $SystemErrors.Value.Add(
            [PSCustomObject]@{
                DateTime = Get-Date
                Message  = "Event log write failed: $_"
            }
        )
    }
}
function Write-SystemErrorLog {
    param(
        [Parameter(Mandatory)][object]$SystemErrors,
        [Parameter(Mandatory)][string]$LogFolder,
        [Parameter(Mandatory)][ref]$MailParams
    )
    if ($SystemErrors.Count -gt 0 -and (Test-Path -LiteralPath $LogFolder -PathType Container)) {
        $path = Join-Path (Get-DatedLogFolderPathHC) 'System errors log'
        $attachments = Out-LogFileHC -DataToExport $SystemErrors -PartialPath $path -FileExtensions '.json' -ErrorAction Ignore
        if ($attachments) {
            if (-not $MailParams.Value.Contains('Attachments')) { $MailParams.Value['Attachments'] = @() }
            $MailParams.Value['Attachments'] += $attachments
        }
    }
}
