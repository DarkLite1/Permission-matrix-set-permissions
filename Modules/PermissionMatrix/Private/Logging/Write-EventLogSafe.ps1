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
        Add-ErrorByCategoryHC `
            -Type 'Warning' `
            -Name 'EventLogWriteFailed' `
            -Message "Failed writing to event log: $_" `
            -Category 'Logging' `
            -SystemErrors ([ref]$SystemErrors)
    }
}