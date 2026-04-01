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