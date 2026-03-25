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
function Out-LogFileHC {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [PSCustomObject[]]$DataToExport,
        [Parameter(Mandatory)]
        [String]$PartialPath,
        [Parameter(Mandatory)]
        [String[]]$FileExtensions,
        [hashtable]$ExcelFile = @{
            SheetName = 'Overview'
            TableName = 'Overview'
            CellStyle = $null
        },
        [Switch]$Append
    )

    $allLogFilePaths = @()

    foreach (
        $fileExtension in
        $FileExtensions | Sort-Object -Unique
    ) {
        try {
            $logFilePath = "$PartialPath$fileExtension"

            Write-Verbose "Export $($DataToExport.Count) object(s) to '$logFilePath'"

            switch ($fileExtension) {
                '.csv' {
                    $params = @{
                        LiteralPath       = $logFilePath
                        Append            = $Append
                        Delimiter         = ';'
                        NoTypeInformation = $true
                    }
                    $DataToExport | Export-Csv @params

                    break
                }
                '.json' {
                    #region Convert error object to error message string
                    $convertedDataToExport = foreach (
                        $exportObject in
                        $DataToExport
                    ) {
                        foreach ($property in $exportObject.PSObject.Properties) {
                            $name = $property.Name
                            $value = $property.Value
                            if (
                                $value -is [System.Management.Automation.ErrorRecord]
                            ) {
                                if (
                                    $value.Exception -and $value.Exception.Message
                                ) {
                                    $exportObject.$name = $value.Exception.Message
                                }
                                else {
                                    $exportObject.$name = $value.ToString()
                                }
                            }
                        }
                        $exportObject
                    }
                    #endregion

                    if (
                        $Append -and
                        (Test-Path -LiteralPath $logFilePath -PathType Leaf)
                    ) {
                        $params = @{
                            LiteralPath = $logFilePath
                            Raw         = $true
                            Encoding    = 'UTF8'
                        }
                        $jsonFileContent = Get-Content @params | ConvertFrom-Json

                        $convertedDataToExport = [array]$convertedDataToExport + [array]$jsonFileContent
                    }

                    $convertedDataToExport |
                    ConvertTo-Json -Depth 7 |
                    Out-File -LiteralPath $logFilePath

                    break
                }
                '.txt' {
                    $params = @{
                        LiteralPath = $logFilePath
                        Append      = $Append
                    }

                    $DataToExport | Format-List -Property * -Force |
                    Out-File @params

                    break
                }
                '.xlsx' {
                    if (
                        (-not $Append) -and
                        (Test-Path -LiteralPath $logFilePath -PathType Leaf)
                    ) {
                        $logFilePath | Remove-Item
                    }

                    $excelParams = @{
                        Path          = $logFilePath
                        Append        = $true
                        AutoNameRange = $true
                        AutoSize      = $true
                        FreezeTopRow  = $true
                        WorksheetName = $ExcelFile.SheetName
                        TableName     = $ExcelFile.TableName
                        Verbose       = $false
                    }
                    if ($ExcelFile.CellStyle) {
                        $excelParams.CellStyleSB = $ExcelFile.CellStyle
                    }
                    $DataToExport | Export-Excel @excelParams

                    break
                }
                default {
                    throw "Log file extension '$_' not supported. Supported values are '.csv', '.json', '.txt' or '.xlsx'."
                }
            }

            $allLogFilePaths += $logFilePath
        }
        catch {
            Write-Warning "Failed creating log file '$logFilePath': $_"
        }
    }

    $allLogFilePaths
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
function Write-EventsToEventLogHC {
    <#
        .SYNOPSIS
            Write events to the event log.

        .DESCRIPTION
            The use of this function will allow standardization in the Windows
            Event Log by using the same EventID's and other properties across
            different scripts.

            Custom Windows EventID's based on the PowerShell standard streams:

            PowerShell Stream     EventIcon    EventID   EventDescription
            -----------------     ---------    -------   ----------------
            [i] Info              [i] Info     100       Script started
            [4] Verbose           [i] Info     4         Verbose message
            [1] Output/Success    [i] Info     1         Output on success
            [3] Warning           [w] Warning  3         Warning message
            [2] Error             [e] Error    2         Fatal error message
            [i] Info              [i] Info     199       Script ended successfully

        .PARAMETER Source
            Specifies the script name under which the events will be logged.

        .PARAMETER LogName
            Specifies the name of the event log to which the events will be
            written. If the log does not exist, it will be created.

        .PARAMETER Events
            Specifies the events to be written to the event log. This should be
            an array of PSCustomObject with properties: Message, EntryType, and
            EventID.

        .PARAMETER Events.xxx
            All properties that are not 'EntryType' or 'EventID' will be used to
            create a formatted message.

        .PARAMETER Events.EntryType
            The type of the event.

            The following values are supported:
            - Information
            - Warning
            - Error
            - SuccessAudit
            - FailureAudit

            The default value is Information.

        .PARAMETER Events.EventID
            The ID of the event. This should be a number.
            The default value is 4.

        .EXAMPLE
            $eventLogData = [System.Collections.Generic.List[PSObject]]::new()

            $eventLogData.Add(
                [PSCustomObject]@{
                    Message   = 'Script started'
                    EntryType = 'Information'
                    EventID   = '100'
                }
            )
            $eventLogData.Add(
                [PSCustomObject]@{
                    Message  = 'Failed to read the file'
                    FileName = 'C:\Temp\test.txt'
                    DateTime = Get-Date
                    EntryType = 'Error'
                    EventID   = '2'
                }
            )
            $eventLogData.Add(
                [PSCustomObject]@{
                    Message  = 'Created file'
                    FileName = 'C:\Report.xlsx'
                    FileSize = 123456
                    DateTime = Get-Date
                    EntryType = 'Information'
                    EventID   = '1'
                }
            )
            $eventLogData.Add(
                [PSCustomObject]@{
                    Message   = 'Script finished'
                    EntryType = 'Information'
                    EventID   = '199'
                }
            )

            $params = @{
                Source  = 'Test (Brecht)'
                LogName = 'HCScripts'
                Events  = $eventLogData
            }
            Write-EventsToEventLogHC @params
        #>

    [CmdLetBinding()]
    param (
        [Parameter(Mandatory)]
        [String]$Source,
        [Parameter(Mandatory)]
        [String]$LogName,
        [PSCustomObject[]]$Events
    )

    try {
        if ([System.Diagnostics.EventLog]::SourceExists($Source)) {
            $existingLogName = [System.Diagnostics.EventLog]::LogNameFromSourceName($Source, '.')

            if ($existingLogName -ne $LogName) {
                throw "The event log source '$Source' is already registered with event log name '$existingLogName', it cannot be used with log name '$LogName'."
            }
        }
        else {
            Write-Verbose "Create event log source '$Source' with log name '$LogName'"

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

            if (-not $params.EntryType) {
                $params.EntryType = 'Information'
            }
            if (-not $params.EventID) {
                $params.EventID = 4
            }

            foreach (
                $property in
                $eventItem.PSObject.Properties | Where-Object {
                    ($_.Name -ne 'EntryType') -and ($_.Name -ne 'EventID')
                }
            ) {
                $params.Message += "`n- $($property.Name) '$($property.Value)'"
            }

            Write-Verbose "Write event to log '$LogName' source '$Source' message '$($params.Message)'"

            Write-EventLog @params
        }
    }
    catch {
        throw "Failed to write to event log '$LogName' source '$Source': $_"
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
