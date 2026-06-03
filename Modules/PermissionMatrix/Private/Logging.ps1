function Remove-OldLogsHC {
    <#
    .SYNOPSIS
        Purges old log files and orphaned directories based on a retention 
        policy.

    .DESCRIPTION
        Evaluates files within the specified log directory against a given 
        retention threshold (in days). Files with a 'CreationTime' older than 
        the threshold are permanently deleted. 
        
        Following the file cleanup, the function performs a highly efficient 
        bottom-up (descending sort) evaluation of the directory tree, removing 
        any subdirectories that are now empty. 
        
        Architectural Note: Deletion operations frequently encounter locked 
        files (e.g., if a log is currently open in another process). This 
        function safely catches those exceptions and appends them as 
        non-terminating 'Warning' records to the SystemErrors reference, 
        ensuring that cleanup failures never crash the main orchestrator.

    .PARAMETER LogFolder
        The absolute path to the root logging directory to be evaluated.

    .PARAMETER RetentionDays
        The number of days to retain logs. Files older than this threshold will 
        be deleted. A value of 0 or less will instantly bypass the cleanup 
        process.

    .PARAMETER SystemErrors
        A reference variable ([ref]) containing a List[pscustomobject]. Used to 
        capture and bubble up file-lock exceptions or permission errors as 
        warnings.

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

function Out-LogFileHC {
    <#
    .SYNOPSIS
        A versatile data export engine that writes PowerShell objects to 
        multiple file formats simultaneously.

    .DESCRIPTION
        Takes an array of custom objects and exports them to one or more 
        requested file formats (CSV, JSON, TXT, XLSX) using a shared base path. 
        
        It includes intelligent data handling for specific formats:
        - JSON: 
            Automatically intercepts and unwraps [System.Management. Automation.
            ErrorRecord] objects into flat string messages to prevent 
            serialization depth failures. Custom logic is also used to safely 
            append to existing JSON arrays.
        - XLSX: 
            Leverages the ImportExcel module to dynamically build formatted 
            Excel tables with frozen headers and auto-sized columns.

    .PARAMETER DataToExport
        An array of PSCustomObject items containing the data rows to be written 
        to disk.

    .PARAMETER PartialPath
        The absolute file path minus the extension 
        (e.g., 'C:\Logs\ExecutionReport'). The script will append the requested 
        extensions to this base path.

    .PARAMETER FileExtensions
        An array of string extensions dictating the desired output formats. 
        Valid values: '.csv', '.json', '.txt', '.xlsx'.

    .PARAMETER ExcelFile
        A hashtable defining the structural formatting rules for '.xlsx' 
        exports. 
        Expected keys: 'SheetName', 'TableName', and 'CellStyle'.

    .PARAMETER Append
        If specified, the function will attempt to append the new data to 
        existing files rather than overwriting them. 

    .OUTPUTS
        System.String[]
        Returns an array of strings representing the absolute paths of all 
        successfully generated or updated log files.

    .EXAMPLE
        $data = @(
            [pscustomobject]@{ Status = 'Success'; Server = 'SRV-01' }
            [pscustomobject]@{ Status = 'Failed'; Server = 'SRV-02' }
        )
        
        $extensions = @('.csv', '.json', '.xlsx')
        
        $exportedPaths = Out-LogFileHC `
            -DataToExport $data `
            -PartialPath 'C:\Logs\DailyReport' `
            -FileExtensions $extensions
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
    <#
    .SYNOPSIS
        Safely deletes a specified file and handles locking/permission errors 
        non-destructively.

    .DESCRIPTION
        Attempts to forcefully remove a target file if it currently exists on 
        the disk. 
        
        To ensure the stability of the broader orchestrator, this function will 
        never throw a terminating error. If the file is locked by another 
        process or access is denied, it safely catches the exception and logs 
        it as a 'Warning'. 
        
        It intelligently routes this warning: if the `$SystemErrors` reference 
        variable is provided, the error is added to the centralized collection. 
        If it is omitted, it falls back to the standard PowerShell warning 
        stream.

    .PARAMETER FilePath
        The absolute path to the target file that should be deleted.

    .PARAMETER SystemErrors
        An optional reference variable ([ref]) containing a List
        [pscustomobject]. Used to capture and bubble up file deletion failures 
        as structured warnings rather than crashing the script.

    .EXAMPLE
        # Standard deletion with console warnings on failure
        Remove-FileHC -FilePath 'C:\Temp\OldLog.txt'

    .EXAMPLE
        # Silent deletion routing failures to the global error tracker
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
        Safely formats and writes aggregated execution data and system errors 
        to the Windows Event Log.

    .DESCRIPTION
        Acts as a robust, non-terminating wrapper for system-level event 
        logging. When Event Logging is enabled in the configuration, this 
        function performs three critical steps:
        
        1. Error Consolidation: 
            Extracts all accumulated pipeline failures from the '$SystemErrors' 
            collection and translates them into standalone 'Error' events 
            (EventID 2).
        2. Execution Closure: 
            Appends a standardized "Script ended" 'Information' event (EventID 
            199) to formally mark the end of the run.
        3. Safety Truncation: 
            Automatically scans all outgoing messages and truncates anything 
            exceeding 31,000 characters. This prevents the underlying Windows 
            Event Log API from throwing fatal serialization errors when 
            processing massive stack traces or data dumps.

        If the function lacks permissions to create the Event Source or write 
        to the log, it safely catches the exception and appends a 'Warning' 
        back to the `$SystemErrors` reference.

    .PARAMETER EventLogData
        A generic list of PSCustomObjects containing the baseline event data 
        (like execution statistics and timestamps) to be written to the log.

    .PARAMETER ScriptName
        The string name to be used as the Event Log 'Source' 
        (e.g., 'Permission Matrix').

    .PARAMETER Settings
        The parsed JSON configuration settings object containing the 
        `SaveInEventLog` rules (LogName, Save boolean).

    .PARAMETER SystemErrors
        A reference variable ([ref]) containing a List[pscustomobject]. These 
        captured errors are unwrapped and directly injected into the Event Log 
        stream.

    .EXAMPLE
        $sysErrors = [System.Collections.Generic.List[pscustomobject]]::new()
        $eventData = [System.Collections.Generic.List[pscustomobject]]::new()
        
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
        Dynamically formats and writes an array of custom objects to the 
        Windows Event Log.

    .DESCRIPTION
        This function handles the physical writing of data to the Windows Event 
        Log. It evaluates the provided array of event objects and processes 
        them using the following logic:

        1. Source Registration: 
            Checks if the specified Event Source exists in the target Log. If 
            it is missing, it automatically creates it (Note: creating a new 
            Event Source requires Administrator privileges).
        2. Dynamic Message Construction: 
            It extracts the 'EntryType' and 'EventID' properties from the 
            object. All remaining properties are dynamically iterated over and 
            flattened into a bulleted string to construct the final Event Log 
            'Message'.
        3. Fallbacks: 
            If an object is missing an 'EntryType', it defaults to 
            'Information'. If it is missing an 'EventID', it defaults to '4'.

        Unlike its parent wrapper (Write-EventLogSafeHC), this function will 
        throw a terminating error if it fails to write, passing the exception 
        back up the chain.

    .PARAMETER Source
        The name of the application or script generating the event 
        (e.g., 'Permission Matrix'). This becomes the 'Source' column in the 
        Event Viewer.

    .PARAMETER LogName
        The name of the target Windows Event Log 
        (e.g., 'Application' or 'System').

    .PARAMETER Events
        An array of PSCustomObjects containing the data to log. Properties 
        named 'EntryType' and 'EventID' map directly to Event Log fields, while 
        all other properties are concatenated into the message body.

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
        Exports system errors to a JSON log file and automatically attaches it 
        to the outgoing email parameters.

    .DESCRIPTION
        This function processes the global collection of system errors 
        encountered during the pipeline's execution. 
        
        If errors exist, it resolves the appropriately dated log directory and 
        serializes the error records into a 'SystemErrors.json' file using the 
        Out-LogFileHC engine. Finally, it safely modifies the referenced 
        `$MailParams` hashtable to append the newly generated JSON file to its 
        'Attachments' array. This ensures that administrators receive the full, 
        raw error data alongside the HTML summary email.

    .PARAMETER SystemErrors
        An array or collection of PSCustomObjects representing the captured 
        pipeline errors and warnings.

    .PARAMETER LogFolder
        The absolute path to the root logging directory where the file will be 
        saved.

    .PARAMETER MailParams
        A reference variable ([ref]) containing the hashtable of SMTP 
        parameters destined for the email sending function. The function will 
        dynamically create or update the 'Attachments' key within this 
        hashtable.

    .PARAMETER ScriptStartTime
        The exact DateTime the script started executing. Used to ensure the log 
        file is placed in the correct timestamped subfolder.

    .PARAMETER JsonFileName
        The base name of the configuration file, utilized by the folder 
        creation logic to maintain consistent directory naming conventions.

    .EXAMPLE
        $mailSplat = @{ 
            To      = 'admin@domain.com'
            Subject = 'Execution Report' 
        }
        
        Write-SystemErrorLogHC `
            -SystemErrors $sysErrors `
            -LogFolder 'C:\MatrixLogs' `
            -MailParams ([ref]$mailSplat)
            
        # $mailSplat now contains an 'Attachments' array with the path to 
        SystemErrors.json
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