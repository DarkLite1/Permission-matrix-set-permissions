#Requires -Version 7.0
<#
.SYNOPSIS
    Builds an Excel overview of all errors and warnings found in the
    "Permission matrix set permissions" log folders.

.DESCRIPTION
    Walks every run folder (e.g. "2026_07_26_220004 (BNL Nightly)"), reads the
    per-matrix-file logs and produces a workbook with four sheets:

        Errors & Warnings    MatrixFileName / DateTime / Error / Warning / ComputerName
        Summary              totals and breakdowns (COUNTIFS formulas)
        All issues (detail)  same rows + Information level + full context
        Notes & method       how the data was collected, with caveats

    Two log formats are supported and detected automatically:

        old (up to 2026-06-29) : "ID <n> - Settings.html" per setting
                                 (+ "00 - Troubleshooting Log.html")
        new (from 2026-06-30)  : "00 - Execution Report.html" per matrix file

    Run level system issues are read from "SystemErrors.json" or
    "System errors log.json".

    Run folders are read one at a time by default. Add -Parallel when the log
    share is remote: reading a few thousand small files over SMB is latency
    bound, and that is the case where concurrency pays for itself.

.PARAMETER LogRoot
    Either the "Permission matrix set permissions (BNL)" folder itself or its
    parent. Also accepts a folder that was unzipped from the log archive.

.PARAMETER OutputFile
    Path of the .xlsx to create.

.PARAMETER CsvFolder
    Optional. Also writes Issues.csv and Runs.csv there (handy for a quick diff
    or for feeding another tool). No Excel module needed for this part.

.PARAMETER Parallel
    Read several run folders at the same time. Worth it when the log share is
    remote, because then the time goes into per-file network latency rather than
    into CPU. On a local disk it is slower than the default: every runspace has
    to be created and has to re-parse the parser definitions, and that overhead
    is not repaid unless the reads are actually waiting on something.

.PARAMETER ThrottleLimit
    How many run folders -Parallel reads at the same time. Default 8. Ignored
    without -Parallel.

.EXAMPLE
    .\PermissionMatrixIssueReport.ps1 `
        -LogRoot '\\BELSGFRANIT07\Log\File or folder\Permission matrix\Permission matrix set permissions (BNL)' `
        -OutputFile 'C:\Temp\Matrix issues.xlsx'

.EXAMPLE
    # counts only, no workbook, per folder progress messages
    .\PermissionMatrixIssueReport.ps1 -LogRoot D:\Logs -Verbose

.EXAMPLE
    # log share over the network: read 12 run folders at a time
    .\PermissionMatrixIssueReport.ps1 -LogRoot \\BELSGFRANIT07\Log\... -Parallel -ThrottleLimit 12

.NOTES
    PowerShell 7.0 or later.
    Requires the ImportExcel module for the workbook (Excel itself is NOT
    needed):  Install-Module ImportExcel -Scope CurrentUser
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $LogRoot,

    [string] $OutputFile,

    [string] $CsvFolder,

    [switch] $Parallel,

    [ValidateRange(1, 25)]
    [int] $ThrottleLimit = 8

)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region ------------------------------------------------------------- parser ---
# Everything that reads the logs lives in this one script block. It is dot
# sourced into the current scope for the sequential path, and its source text is
# dot sourced again inside every parallel runspace - ForEach-Object -Parallel
# starts with an empty function table, so the definitions have to travel as text.

$parser = {

    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    $script:RxOptions = [System.Text.RegularExpressions.RegexOptions]::Singleline -bor `
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase

    # Severity words used in the report pills; the icon glyphs are filtered by shape.
    $script:SeverityWords = @('ERROR', 'WARNING', 'INFORMATION', 'INFO')

    function Read-TextFile {
        param([string] $Path)
        # ReadAllText is noticeably faster than Get-Content for ~16.000 small files.
        return [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
    }

    function Get-CleanText {
        param([string] $Text)
        if ([string]::IsNullOrEmpty($Text)) { return '' }
        $t = [System.Net.WebUtility]::HtmlDecode($Text)
        return ($t -replace '\s+', ' ').Trim()
    }

    function Remove-HtmlNoise {
        <# Drops comments (the Outlook/mso duplicates live in there), style and
           script blocks, so every severity pill appears exactly once. #>
        param([string] $Html)
        $h = [regex]::Replace($Html, '<!--.*?-->', '', $script:RxOptions)
        $h = [regex]::Replace($h, '<style.*?</style>', '', $script:RxOptions)
        return [regex]::Replace($h, '<script.*?</script>', '', $script:RxOptions)
    }

    function Get-HtmlTextBlocks {
        <# Turns an HTML fragment into its visible text blocks, in order,
           dropping icon glyphs and severity words. #>
        param([string] $Fragment)

        $stripped = [regex]::Replace($Fragment, '<[^>]+>', "`n")
        $blocks = [System.Collections.Generic.List[string]]::new()

        foreach ($line in ($stripped -split "`n")) {
            $t = Get-CleanText $line
            if ($t -eq '') { continue }
            if ($script:SeverityWords -contains $t.ToUpperInvariant()) { continue }
            # icon glyphs such as the cross, warning triangle, bullet, check mark
            if ($t.Length -le 2 -and $t -notmatch '[0-9A-Za-z]') { continue }
            if (-not $blocks.Contains($t)) { $blocks.Add($t) }
        }
        return $blocks
    }

    function Get-SeverityName {
        param([string] $Raw)
        $r = (Get-CleanText $Raw).ToUpperInvariant()
        if ($r -like 'ERR*') { return 'Error' }
        if ($r -like 'WARN*') { return 'Warning' }
        if ($r -like 'INFO*') { return 'Information' }
        return 'Unknown'
    }

    function Convert-DmyDate {
        <# "22/07/2026 22:00:04" -> [datetime]. Returns $null when absent. #>
        param([string] $Text)
        if ([string]::IsNullOrEmpty($Text)) { return $null }
        $m = [regex]::Match($Text, '(\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2})')
        if (-not $m.Success) { return $null }
        $result = [datetime]::MinValue
        $ok = [datetime]::TryParseExact($m.Groups[1].Value, 'dd/MM/yyyy HH:mm:ss',
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None, [ref] $result)
        if ($ok) { return $result }
        return $null
    }

    function Convert-IsoDate {
        <# "2026-07-13T12:11:27.6697062+02:00" -> [datetime] 13/07/2026 12:11:27.
           The offset is deliberately ignored: the logs are written in server
           local time and that is the value an administrator wants to see back. #>
        param([string] $Text)
        if ([string]::IsNullOrEmpty($Text)) { return $null }
        $m = [regex]::Match($Text, '^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2}):(\d{2})')
        if ($m.Success) {
            return [datetime]::new(
                [int]$m.Groups[1].Value, [int]$m.Groups[2].Value, [int]$m.Groups[3].Value,
                [int]$m.Groups[4].Value, [int]$m.Groups[5].Value, [int]$m.Groups[6].Value)
        }
        $dto = [System.DateTimeOffset]::MinValue
        $ok = [System.DateTimeOffset]::TryParse($Text,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None, [ref] $dto)
        if ($ok) { return $dto.DateTime }
        return $null
    }

    function Get-RunFolderDate {
        <# "2026_07_26_220004 (BNL Nightly)" -> [datetime] 2026-07-26 22:00:04 #>
        param([string] $RunFolder)
        $m = [regex]::Match($RunFolder, '^(\d{4})_(\d{2})_(\d{2})_(\d{2})(\d{2})(\d{2})')
        if (-not $m.Success) { return $null }
        return [datetime]::new(
            [int]$m.Groups[1].Value, [int]$m.Groups[2].Value, [int]$m.Groups[3].Value,
            [int]$m.Groups[4].Value, [int]$m.Groups[5].Value, [int]$m.Groups[6].Value)
    }

    function Get-RunType {
        param([string] $RunFolder)
        $m = [regex]::Match($RunFolder, '\((.*?)\)')
        if ($m.Success) { return $m.Groups[1].Value }
        return ''
    }

    function Get-DetailFileDateTime {
        <# Exact timestamp of a single problem, from the detail JSON it links to. #>
        param([string] $Folder, [string] $Href)
        if ([string]::IsNullOrEmpty($Href)) { return $null }
        $name = ($Href -split '[\\/]')[-1]
        $path = Join-Path $Folder $name
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $null }
        try {
            $head = Read-TextFile $path
            if ($head.Length -gt 600) { $head = $head.Substring(0, 600) }
            $m = [regex]::Match($head, '"DateTime"\s*:\s*"([^"]+)"')
            if ($m.Success) { return Convert-IsoDate $m.Groups[1].Value }
        }
        catch {
            Write-Verbose "Could not read detail file '$path': $($_.Exception.Message)"
        }
        return $null
    }

    function New-Issue {
        param(
            [string] $MatrixFileName, $DateTime, [string] $Severity, [string] $Problem,
            [string] $Description, [string] $ComputerName, [string] $Path, [string] $Site,
            [string] $SettingId, [string] $RunFolder, [string] $LogFormat,
            [string] $DetailFile, [string] $TimestampSource
        )
        return [pscustomobject] @{
            MatrixFileName  = $MatrixFileName
            DateTime        = $DateTime
            Severity        = $Severity
            Problem         = $Problem
            Description     = $Description
            ComputerName    = $ComputerName
            Path            = $Path
            Site            = $Site
            SettingId       = $SettingId
            RunFolder       = $RunFolder
            RunType         = (Get-RunType $RunFolder)
            LogFormat       = $LogFormat
            DetailFile      = $DetailFile
            TimestampSource = $TimestampSource
        }
    }

    # ---------------------------------------------------- new format parser ---

    function Read-ExecutionReport {
        <# Parses "00 - Execution Report.html" (runs from 2026-06-30 onwards).

           The report is one flat stream of cards; problems belong to the settings
           card that precedes them, so a single pass in document order is enough.
           Three markup variants exist for a problem row (table with
           rr-content-cell, table with spans, and a flex div) - all three put the
           severity pill last, which is what bounds the row. #>
        param([string] $ReportPath, [string] $RunFolder, [string] $MatrixFolder,
            [System.Collections.Generic.List[object]] $Issues)

        $folder = [System.IO.Path]::GetDirectoryName($ReportPath)
        $html = Remove-HtmlNoise (Read-TextFile $ReportPath)

        $matrixName = "$MatrixFolder.xlsx"
        $m = [regex]::Match($html, '<title>\s*Execution Report\s*-\s*(.*?)</title>', $script:RxOptions)
        if ($m.Success) {
            $t = Get-CleanText $m.Groups[1].Value
            if ($t -ne '') { $matrixName = $t }
        }

        $bodyIndex = $html.IndexOf('<body')
        if ($bodyIndex -lt 0) { $bodyIndex = 0 }
        $body = $html.Substring($bodyIndex)

        # run start / end time from the footer
        $startTime = $null
        $endTime = $null
        $m = [regex]::Match($body, 'Start time.{0,400}?(\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2})', $script:RxOptions)
        if ($m.Success) { $startTime = Convert-DmyDate $m.Groups[1].Value }
        $m = [regex]::Match($body, 'End time.{0,400}?(\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2})', $script:RxOptions)
        if ($m.Success) { $endTime = Convert-DmyDate $m.Groups[1].Value }
        if ($null -eq $startTime) { $startTime = Get-RunFolderDate $RunFolder }

        # overall status pill in the page header ("Success", "1 Error, 2 Warnings", ...)
        $status = ''
        $m = [regex]::Match($body, 'rr-status-cell.*?>([^<>]{1,60})</span>', $script:RxOptions)
        if ($m.Success) { $status = Get-CleanText $m.Groups[1].Value }

        $computer = ''; $path = ''; $site = ''; $settingId = ''
        $settingCount = 0
        $counts = @{ Error = 0; Warning = 0; Information = 0; Unknown = 0 }

        $markers = [regex]::Matches($body, 'class="rr-(settings-head|check-row)"')
        for ($i = 0; $i -lt $markers.Count; $i++) {
            $from = $markers[$i].Index
            $to = if ($i + 1 -lt $markers.Count) { $markers[$i + 1].Index } else { $body.Length }
            $chunk = $body.Substring($from, $to - $from)
            # the match starts inside the opening tag, so skip the rest of that
            # tag - otherwise its attributes would be read as visible text
            $gt = $chunk.IndexOf('>')
            if ($gt -ge 0) { $chunk = $chunk.Substring($gt + 1) }

            if ($markers[$i].Groups[1].Value -eq 'settings-head') {
                $settingCount++
                $computer = ''; $path = ''; $site = ''; $settingId = ''
                $m = [regex]::Match($chunk, '<div[^>]*font-size:14px[^>]*>([^<]*)</div>')
                if ($m.Success) { $computer = Get-CleanText $m.Groups[1].Value }
                $m = [regex]::Match($chunk, 'class="rr-path"[^>]*>([^<]*)<')
                if ($m.Success) { $path = Get-CleanText $m.Groups[1].Value }
                $m = [regex]::Match($chunk, '>Site:</span>\s*<span[^>]*>([^<]*)</span>')
                if ($m.Success) { $site = Get-CleanText $m.Groups[1].Value }
                $m = [regex]::Match($chunk, 'title="([0-9a-fA-F\-]{36})"')
                if ($m.Success) { $settingId = $m.Groups[1].Value }
                continue
            }

            # ---- problem row ---------------------------------------------
            $pill = [regex]::Match($chunk, '>\s*(ERROR|WARNING|INFORMATION|INFO)\s*<')
            if ($pill.Success) {
                $severity = Get-SeverityName $pill.Groups[1].Value
                # cut at the '<' that opens the pill so its own tag cannot leak text
                $cut = $chunk.LastIndexOf('<', $pill.Index)
                if ($cut -lt 0) { $cut = $pill.Index }
                $content = $chunk.Substring(0, $cut)
            }
            else {
                $severity = 'Unknown'
                $content = $chunk
            }

            $href = ''
            $m = [regex]::Match($content, "href='([^']*)'")
            if (-not $m.Success) { $m = [regex]::Match($content, 'href="([^"]*)"') }
            if ($m.Success) { $href = $m.Groups[1].Value }

            # @() keeps this an array even when there is a single text block
            $blocks = @(Get-HtmlTextBlocks $content)
            $category = ''; $problem = ''; $description = ''
            if ($blocks.Count -ge 3) {
                $category = $blocks[0]
                $problem = $blocks[1]
                $description = ($blocks[2..($blocks.Count - 1)] -join ' ')
            }
            elseif ($blocks.Count -eq 2) {
                $problem = $blocks[0]; $description = $blocks[1]
            }
            elseif ($blocks.Count -eq 1) {
                $problem = $blocks[0]
            }
            if ($category -ne '') {
                $problem = if ($problem -ne '') { "$category`: $problem" } else { $category }
            }

            $stamp = Get-DetailFileDateTime -Folder $folder -Href $href
            $stampSource = 'problem detail file'
            if ($null -eq $stamp) { $stamp = $startTime; $stampSource = 'run start time' }

            $counts[$severity] = $counts[$severity] + 1
            $Issues.Add((New-Issue -MatrixFileName $matrixName -DateTime $stamp `
                        -Severity $severity -Problem $problem -Description $description `
                        -ComputerName $computer -Path $path -Site $site -SettingId $settingId `
                        -RunFolder $RunFolder -LogFormat 'new' -DetailFile $href `
                        -TimestampSource $stampSource))
        }

        return [pscustomobject] @{
            RunFolder      = $RunFolder
            RunType        = (Get-RunType $RunFolder)
            MatrixFileName = $matrixName
            Status         = $status
            Settings       = $settingCount
            Errors         = $counts['Error']
            Warnings       = $counts['Warning']
            Information    = $counts['Information']
            StartTime      = $startTime
            EndTime        = $endTime
            LogFormat      = 'new'
        }
    }

    # ---------------------------------------------------- old format parser ---

    function Read-SettingsReports {
        <# Parses the "ID <n> - Settings.html" files of one matrix folder (runs up
           to 2026-06-29). Each file is one setting = one computer, and carries its
           own start/end time, which is the best timestamp available in this
           format - the detail .txt files hold no DateTime. #>
        param([string] $MatrixDir, [string] $RunFolder, [string] $MatrixFolder,
            [System.Collections.Generic.List[object]] $Issues)

        $matrixName = "$MatrixFolder.xlsx"
        $settingFiles = @(Get-ChildItem -LiteralPath $MatrixDir -Filter 'ID * - Settings.html' -File |
            Sort-Object Name)

        $errors = 0; $warnings = 0; $information = 0
        $starts = [System.Collections.Generic.List[datetime]]::new()
        $ends = [System.Collections.Generic.List[datetime]]::new()

        foreach ($file in $settingFiles) {
            $html = Remove-HtmlNoise (Read-TextFile $file.FullName)

            $m = [regex]::Match($html, 'id="matrixTitle".*?>([^<>]+\.xlsx)</a>', $script:RxOptions)
            if ($m.Success) { $matrixName = Get-CleanText $m.Groups[1].Value }

            $startTime = $null
            $m = [regex]::Match($html, '<th>\s*Start time\s*</th>\s*<td>(.*?)</td>', $script:RxOptions)
            if ($m.Success) { $startTime = Convert-DmyDate $m.Groups[1].Value }
            $m = [regex]::Match($html, '<th>\s*End time\s*</th>\s*<td>(.*?)</td>', $script:RxOptions)
            if ($m.Success) {
                $e = Convert-DmyDate $m.Groups[1].Value
                if ($null -ne $e) { $ends.Add($e) }
            }
            if ($null -ne $startTime) { $starts.Add($startTime) }
            $stampSource = 'setting start time'
            if ($null -eq $startTime) {
                $startTime = Get-RunFolderDate $RunFolder
                $stampSource = 'run start time'
            }

            # settings summary row: <probType cell><ID><ComputerName><Path><Action><Duration>
            $computer = ''; $path = ''; $settingId = ''
            $m = [regex]::Match($html,
                '<tr>\s*<td id="probType\w+"[^>]*>\s*</td>\s*<td>\s*(\d+)\s*</td>\s*<td>([^<]*)</td>\s*<td>([^<]*)</td>',
                $script:RxOptions)
            if ($m.Success) {
                $settingId = Get-CleanText $m.Groups[1].Value
                $computer = Get-CleanText $m.Groups[2].Value
                $path = Get-CleanText $m.Groups[3].Value
            }
            if ($settingId -eq '') {
                $m = [regex]::Match($file.Name, 'ID (\d+) - Settings')
                if ($m.Success) { $settingId = $m.Groups[1].Value }
            }

            # problem rows
            $rx = '<tr>\s*<td id="probType(Error|Warning|Info)"[^>]*>\s*</td>\s*<td colspan="7">(.*?)</td>\s*</tr>'
            foreach ($pm in [regex]::Matches($html, $rx, $script:RxOptions)) {
                $severity = Get-SeverityName $pm.Groups[1].Value
                $block = $pm.Groups[2].Value

                $problem = ''
                $m = [regex]::Match($block, '<p id="probTitle">(.*?)</p>', $script:RxOptions)
                if ($m.Success) { $problem = Get-CleanText $m.Groups[1].Value }

                $description = ''
                foreach ($p in [regex]::Matches($block, '<p[^>]*>(.*?)</p>', $script:RxOptions)) {
                    $t = Get-CleanText ([regex]::Replace($p.Groups[1].Value, '<[^>]+>', ' '))
                    if ($t -ne '' -and $t -ne $problem) { $description = $t; break }
                }

                $detail = ''
                $m = [regex]::Match($block, 'href="([^"]*)"', $script:RxOptions)
                if ($m.Success) { $detail = ($m.Groups[1].Value -split '[\\/]')[-1] }

                switch ($severity) {
                    'Error' { $errors++ }
                    'Warning' { $warnings++ }
                    'Information' { $information++ }
                }

                $Issues.Add((New-Issue -MatrixFileName $matrixName -DateTime $startTime `
                            -Severity $severity -Problem $problem -Description $description `
                            -ComputerName $computer -Path $path -Site '' -SettingId $settingId `
                            -RunFolder $RunFolder -LogFormat 'old' -DetailFile $detail `
                            -TimestampSource $stampSource))
            }
        }

        $status = if ($errors -gt 0) { 'Error' } elseif ($warnings -gt 0) { 'Warning' } else { 'Success' }
        $runStart = if ($starts.Count -gt 0) { ($starts | Measure-Object -Minimum).Minimum } else { Get-RunFolderDate $RunFolder }
        $runEnd = if ($ends.Count -gt 0) { ($ends | Measure-Object -Maximum).Maximum } else { $null }

        return [pscustomobject] @{
            RunFolder      = $RunFolder
            RunType        = (Get-RunType $RunFolder)
            MatrixFileName = $matrixName
            Status         = $status
            Settings       = $settingFiles.Count
            Errors         = $errors
            Warnings       = $warnings
            Information    = $information
            StartTime      = $runStart
            EndTime        = $runEnd
            LogFormat      = 'old'
        }
    }

    # ------------------------------------------------ run level system logs ---

    function Read-SystemErrorLog {
        <# "SystemErrors.json" (new) / "System errors log.json" (old). These are
           run level: no matrix file and no computer, so they get MatrixFileName
           "(system)" and a "[System]" prefix on the main sheet. #>
        param([string] $RunDir, [string] $RunFolder,
            [System.Collections.Generic.List[object]] $Issues)

        foreach ($name in @('SystemErrors.json', 'System errors log.json')) {
            $path = Join-Path $RunDir $name
            if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { continue }

            $raw = Read-TextFile $path
            $entries = @()
            try {
                $entries = @(ConvertFrom-Json $raw)
            }
            catch {
                # tolerate several objects concatenated in one file
                foreach ($m in [regex]::Matches($raw, '\{.*?\r?\n\}', $script:RxOptions)) {
                    try { $entries += (ConvertFrom-Json $m.Value) } catch { }
                }
            }

            # ConvertFrom-Json turns an ISO timestamp into a [datetime] and shifts
            # it to the local time zone of the machine running this script, which
            # would move the value away from what the log actually says. So take
            # the timestamps from the raw text instead, in document order.
            $stamps = [System.Collections.Generic.List[object]]::new()
            foreach ($m in [regex]::Matches($raw, '"DateTime"\s*:\s*"([^"]+)"')) {
                $stamps.Add((Convert-IsoDate $m.Groups[1].Value))
            }

            $index = -1
            foreach ($e in $entries) {
                $index++
                if ($null -eq $e) { continue }
                $props = $e.PSObject.Properties.Name

                $type = if ($props -contains 'Type' -and $e.Type) { $e.Type } else { 'Error' }
                $problem = if ($props -contains 'Name' -and $e.Name) { Get-CleanText $e.Name } else { 'System issue' }
                $message = ''
                if ($props -contains 'Message' -and $e.Message) { $message = Get-CleanText $e.Message }
                elseif ($props -contains 'Description' -and $e.Description) { $message = Get-CleanText $e.Description }

                $stamp = if ($index -lt $stamps.Count) { $stamps[$index] } else { $null }
                $stampSource = 'system error log'
                if ($null -eq $stamp) { $stamp = Get-RunFolderDate $RunFolder; $stampSource = 'run start time' }

                $Issues.Add((New-Issue -MatrixFileName '(system)' -DateTime $stamp `
                            -Severity ('System ' + (Get-SeverityName $type)) -Problem $problem `
                            -Description $message -ComputerName '' -Path '' -Site '' -SettingId '' `
                            -RunFolder $RunFolder -LogFormat 'system' -DetailFile $name `
                            -TimestampSource $stampSource))
            }
        }
    }

    # --------------------------------------------------------- one run folder ---

    function Read-RunFolder {
        <# Reads one run folder and returns its issues plus one summary row per
           matrix file. This is the unit of parallelism. #>
        param([string] $RunDir, [string] $RunFolder)

        $issues = [System.Collections.Generic.List[object]]::new()
        $runs = [System.Collections.Generic.List[object]]::new()
        $withoutLog = 0

        Read-SystemErrorLog -RunDir $RunDir -RunFolder $RunFolder -Issues $issues

        foreach ($matrixDir in (Get-ChildItem -LiteralPath $RunDir -Directory | Sort-Object Name)) {
            if ($matrixDir.Name -eq 'Export') { continue }

            $newReport = Join-Path $matrixDir.FullName '00 - Execution Report.html'
            $oldSettings = @(Get-ChildItem -LiteralPath $matrixDir.FullName -Filter 'ID * - Settings.html' -File)

            if (Test-Path -LiteralPath $newReport -PathType Leaf) {
                $runs.Add((Read-ExecutionReport -ReportPath $newReport -RunFolder $RunFolder `
                            -MatrixFolder $matrixDir.Name -Issues $issues))
            }
            elseif ($oldSettings.Count -gt 0) {
                $runs.Add((Read-SettingsReports -MatrixDir $matrixDir.FullName -RunFolder $RunFolder `
                            -MatrixFolder $matrixDir.Name -Issues $issues))
            }
            else {
                $withoutLog++
                $runs.Add([pscustomobject] @{
                        RunFolder = $RunFolder; RunType = (Get-RunType $RunFolder)
                        MatrixFileName = "$($matrixDir.Name).xlsx"; Status = 'no log'
                        Settings = 0; Errors = 0; Warnings = 0; Information = 0
                        StartTime = (Get-RunFolderDate $RunFolder); EndTime = $null
                        LogFormat = 'none'
                    })
            }
        }

        Write-Verbose "$RunFolder - $($runs.Count) matrix files, $($issues.Count) issues"
        return [pscustomobject] @{
            Issues            = $issues.ToArray()
            Runs              = $runs.ToArray()
            FoldersWithoutLog = $withoutLog
        }
    }
}

. $parser
$parserSource = $parser.ToString()

#endregion

#region ------------------------------------------------------------ collect ---

if (-not (Test-Path -LiteralPath $LogRoot -PathType Container)) {
    throw "LogRoot '$LogRoot' does not exist or is not a folder."
}

# Accept either the folder holding the run folders, or one level above it.
$root = (Resolve-Path -LiteralPath $LogRoot).Path
$runFolders = @(Get-ChildItem -LiteralPath $root -Directory |
    Where-Object { $_.Name -match '^\d{4}_\d{2}_\d{2}_\d{6}' } |
    Sort-Object Name)
if ($runFolders.Count -eq 0) {
    foreach ($child in (Get-ChildItem -LiteralPath $root -Directory)) {
        $probe = @(Get-ChildItem -LiteralPath $child.FullName -Directory |
            Where-Object { $_.Name -match '^\d{4}_\d{2}_\d{2}_\d{6}' })
        if ($probe.Count -gt 0) { $root = $child.FullName; $runFolders = @($probe | Sort-Object Name); break }
    }
}
if ($runFolders.Count -eq 0) {
    throw "No run folders (yyyy_MM_dd_HHmmss ...) found under '$LogRoot'."
}

$mode = if ($Parallel) { "with $ThrottleLimit threads" } else { 'sequentially' }
Write-Host "Reading $($runFolders.Count) run folders $mode from $root" -ForegroundColor Cyan

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

if ($Parallel) {
    $verbose = $VerbosePreference
    $results = @(
        $runFolders | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
            . ([scriptblock]::Create($using:parserSource))
            $VerbosePreference = $using:verbose
            Read-RunFolder -RunDir $_.FullName -RunFolder $_.Name
        }
    )
}
else {
    $results = @(
        $runFolders | ForEach-Object {
            Read-RunFolder -RunDir $_.FullName -RunFolder $_.Name
        }
    )
}

$issues = [System.Collections.Generic.List[object]]::new()
$runs = [System.Collections.Generic.List[object]]::new()
$foldersWithoutLog = 0
foreach ($r in $results) {
    if ($r.Issues) { $issues.AddRange([object[]] $r.Issues) }
    if ($r.Runs) { $runs.AddRange([object[]] $r.Runs) }
    $foldersWithoutLog += $r.FoldersWithoutLog
}

# -Stable keeps rows with an identical timestamp in a predictable order, so two
# runs of this script over the same logs produce byte-identical sheets.
$allIssues = @(
    $issues | Sort-Object `
        -Stable `
        -Property @{ Expression = 'DateTime'; Descending = $true }, 
    MatrixFileName, ComputerName, Problem
)
$runs = @($runs | Sort-Object -Stable RunFolder, MatrixFileName)

$errorsAndWarnings = @(
    $allIssues | Where-Object { 
        $_.Severity -like '*Error' -or $_.Severity -like '*Warning' 
    }
)

$countError = @($allIssues | Where-Object Severity -EQ 'Error').Count
$countWarning = @($allIssues | Where-Object Severity -EQ 'Warning').Count
$countSysError = @($allIssues | Where-Object Severity -EQ 'System Error').Count
$countSysWarning = @($allIssues | Where-Object Severity -EQ 'System Warning').Count
$countInformation = @($allIssues | Where-Object Severity -Like '*Information').Count

Write-Host ('Matrix file runs : {0}   (new format {1}, old format {2}, no log {3})   in {4:n1}s' -f
    $runs.Count,
    @($runs | Where-Object LogFormat -EQ 'new').Count,
    @($runs | Where-Object LogFormat -EQ 'old').Count,
    $foldersWithoutLog,
    $stopwatch.Elapsed.TotalSeconds)
Write-Host ('Errors {0} (+{1} system)   Warnings {2} (+{3} system)   Information {4}' -f
    $countError, $countSysError, $countWarning, $countSysWarning, $countInformation)

if ($CsvFolder) {
    if (-not (Test-Path -LiteralPath $CsvFolder -PathType Container)) {
        New-Item -ItemType Directory -Path $CsvFolder -Force | Out-Null
    }

    $allIssues | Export-Csv `
        -LiteralPath (Join-Path $CsvFolder 'Issues.csv') `
        -NoTypeInformation -Encoding UTF8

    $runs | Export-Csv `
        -LiteralPath (Join-Path $CsvFolder 'Runs.csv') `
        -NoTypeInformation -Encoding UTF8

    Write-Host "CSV written to $CsvFolder"
}

if (-not $OutputFile) { return }

#endregion

#region ------------------------------------------------------------- Excel ---

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    throw 'The ImportExcel module is required. Install it with: Install-Module ImportExcel -Scope CurrentUser'
}
Import-Module ImportExcel -WarningAction SilentlyContinue

if (Test-Path -LiteralPath $OutputFile -PathType Leaf) { 
    Remove-Item -LiteralPath $OutputFile -Force 
}

$navy = [System.Drawing.Color]::FromArgb(31, 56, 100)
$redBg = [System.Drawing.Color]::FromArgb(252, 228, 228)
$amber = [System.Drawing.Color]::FromArgb(255, 243, 208)
$grey = [System.Drawing.Color]::FromArgb(242, 242, 242)
$white = [System.Drawing.Color]::White
$solid = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$dateFormat = 'dd/mm/yyyy hh:mm:ss'

function Set-HeaderStyle {
    param($Worksheet, [string] $Range)
    $h = $Worksheet.Cells[$Range]
    $h.Style.Font.Bold = $true
    $h.Style.Font.Color.SetColor($script:white)
    $h.Style.Fill.PatternType = $script:solid
    $h.Style.Fill.BackgroundColor.SetColor($script:navy)
}

function Set-ColumnWidths {
    param($Worksheet, [int[]] $Widths)
    for ($c = 1; $c -le $Widths.Count; $c++) { 
        $Worksheet.Column($c).Width = $Widths[$c - 1] 
    }
}

function Set-RowFills {
    <# One fill per row: red for errors, amber for warnings, grey for the rest. #>
    param($Worksheet, [object[]] $Rows, [string] $LastColumn, [switch] $GreyRemainder)
    for ($i = 0; $i -lt $Rows.Count; $i++) {
        $row = $i + 2
        $cells = $Worksheet.Cells["A$row" + ":$LastColumn$row"]
        $cells.Style.Fill.PatternType = $script:solid
        $color = if ($Rows[$i].Severity -like '*Error') { $script:redBg }
        elseif ($Rows[$i].Severity -like '*Warning') { $script:amber }
        else { $script:grey }
        $cells.Style.Fill.BackgroundColor.SetColor($color)
    }
}

# ---- sheet 1: the five requested columns ---------------------------------
$mainRows = foreach ($x in $errorsAndWarnings) {
    $message = if ($x.Severity -like 'System*') { "[System] $($x.Problem)" } else { $x.Problem }
    $isError = $x.Severity -like '*Error'
    [pscustomobject] @{
        MatrixFileName = $x.MatrixFileName
        DateTime       = $x.DateTime
        Error          = if ($isError) { $message } else { $null }
        Warning        = if ($isError) { $null } else { $message }
        ComputerName   = $x.ComputerName
    }
}

$pkg = $mainRows | 
Export-Excel -Path $OutputFile `
    -WorksheetName 'Errors & Warnings' `
    -AutoFilter -FreezeTopRow -PassThru

# ---- sheet 2: full detail ------------------------------------------------
$pkg = $allIssues |
Select-Object MatrixFileName, DateTime, Severity, Problem, Description,
ComputerName, Path, Site, SettingId, RunFolder, RunType,
LogFormat, DetailFile, TimestampSource |
Export-Excel -ExcelPackage $pkg -WorksheetName 'All issues (detail)' `
    -AutoFilter -FreezeTopRow -PassThru

# ---- styling -------------------------------------------------------------
$wsMain = $pkg.Workbook.Worksheets['Errors & Warnings']
$wsMain.Cells.Style.Font.Name = 'Arial'
$wsMain.Cells.Style.Font.Size = 10
Set-HeaderStyle -Worksheet $wsMain -Range 'A1:E1'
Set-ColumnWidths -Worksheet $wsMain -Widths @(42, 19, 40, 40, 18)
$wsMain.Column(2).Style.Numberformat.Format = $dateFormat
$wsMain.Cells['C:D'].Style.WrapText = $true
Set-RowFills -Worksheet $wsMain -Rows $errorsAndWarnings -LastColumn 'E'

$wsDetail = $pkg.Workbook.Worksheets['All issues (detail)']
$wsDetail.Cells.Style.Font.Name = 'Arial'
$wsDetail.Cells.Style.Font.Size = 10
Set-HeaderStyle -Worksheet $wsDetail -Range 'A1:N1'
Set-ColumnWidths -Worksheet $wsDetail -Widths @(42, 19, 14, 38, 70, 18, 46, 16, 38, 34, 13, 11, 46, 22)
$wsDetail.Column(2).Style.Numberformat.Format = $dateFormat
Set-RowFills -Worksheet $wsDetail -Rows $allIssues -LastColumn 'N'

# ---- sheet 3: summary, everything by formula ------------------------------
$detailLast = $allIssues.Count + 1
$det = "'All issues (detail)'"
$sevCol = "$det!`$C`$2:`$C`$$detailLast"
$matCol = "$det!`$A`$2:`$A`$$detailLast"
$probCol = "$det!`$D`$2:`$D`$$detailLast"
$compCol = "$det!`$F`$2:`$F`$$detailLast"
$dtCol = "$det!`$B`$2:`$B`$$detailLast"

$ws = $pkg.Workbook.Worksheets.Add('Summary')
$ws.Cells.Style.Font.Name = 'Arial'
$ws.Cells.Style.Font.Size = 10
$ws.Cells['A1'].Value = 'Permission matrix - errors & warnings overview'
$ws.Cells['A1'].Style.Font.Bold = $true
$ws.Cells['A1'].Style.Font.Size = 13
$ws.Cells['A1'].Style.Font.Color.SetColor($navy)
$ws.Cells['A2'].Value = "All counts are COUNTIFS formulas over the 'All issues (detail)' sheet, so they follow that sheet if it is filtered or edited."
$ws.Cells['A2'].Style.Font.Size = 9
$ws.Cells['A2'].Style.Font.Color.SetColor([System.Drawing.Color]::Gray)

function Add-BlockTitle {
    param($Worksheet, [int] $Row, [string] $Text)
    $cell = $Worksheet.Cells["A$Row"]
    $cell.Value = $Text
    $cell.Style.Font.Bold = $true
    $cell.Style.Font.Size = 11
    $cell.Style.Font.Color.SetColor($script:navy)
}

function Add-BlockHeader {
    param($Worksheet, [int] $Row, [string[]] $Headers)
    for ($c = 0; $c -lt $Headers.Count; $c++) {
        $col = $c + 1
        $Worksheet.Cells[$Row, $col].Value = $Headers[$c]
    }
    $endCol = [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($Headers.Count)
    Set-HeaderStyle -Worksheet $Worksheet -Range ("A$Row" + ":$endCol$Row")
}

$row = 4
Add-BlockTitle -Worksheet $ws -Row $row -Text 'Totals'
$row++
Add-BlockHeader -Worksheet $ws -Row $row -Headers @('Category', 'Count')
$row++
foreach ($pair in @(
        @('Errors (matrix files)', 'Error'),
        @('Warnings (matrix files)', 'Warning'),
        @('System errors', 'System Error'),
        @('System warnings', 'System Warning'),
        @('Information (not an issue)', 'Information'))) {
    $ws.Cells[$row, 1].Value = $pair[0]
    $ws.Cells[$row, 2].Formula = "COUNTIF($sevCol,""$($pair[1])"")"
    $ws.Cells[$row, 2].Style.Font.Bold = $true
    $row++
}
$ws.Cells[$row, 1].Value = 'Errors + warnings (rows on sheet 1)'
$ws.Cells[$row, 2].Formula = "COUNTIF($sevCol,""*Error"")+COUNTIF($sevCol,""*Warning"")"
$ws.Cells[$row, 2].Style.Font.Bold = $true
$row++
foreach ($pair in @(@('First issue timestamp', 'MIN'), @('Last issue timestamp', 'MAX'))) {
    $ws.Cells[$row, 1].Value = $pair[0]
    $ws.Cells[$row, 2].Formula = "$($pair[1])($dtCol)"
    $ws.Cells[$row, 2].Style.Numberformat.Format = $dateFormat
    $ws.Cells[$row, 2].Style.Font.Bold = $true
    $row++
}
$row++

# per matrix file
$byMatrix = @($errorsAndWarnings | Group-Object MatrixFileName | Sort-Object Count -Descending)
Add-BlockTitle -Worksheet $ws -Row $row `
    -Text "Per matrix file ($($byMatrix.Count) with at least one error or warning)"
$row++
Add-BlockHeader -Worksheet $ws -Row $row -Headers @('MatrixFileName', 'Errors', 'Warnings', 'Total', 'Last occurrence')
$row++
foreach ($g in $byMatrix) {
    $ws.Cells[$row, 1].Value = $g.Name
    $ws.Cells[$row, 2].Formula = "COUNTIFS($matCol,`$A$row,$sevCol,""*Error"")"
    $ws.Cells[$row, 3].Formula = "COUNTIFS($matCol,`$A$row,$sevCol,""*Warning"")"
    $ws.Cells[$row, 4].Formula = "B$row+C$row"
    $ws.Cells[$row, 4].Style.Font.Bold = $true
    # SUMPRODUCT keeps this a normal formula - no Ctrl+Shift+Enter needed
    $ws.Cells[$row, 5].Formula = "SUMPRODUCT(MAX(($matCol=`$A$row)*$dtCol))"
    $ws.Cells[$row, 5].Style.Numberformat.Format = $dateFormat
    $row++
}
$row++

# per problem type
Add-BlockTitle -Worksheet $ws -Row $row -Text 'Per problem type'
$row++
Add-BlockHeader -Worksheet $ws -Row $row -Headers @('Problem', 'Errors', 'Warnings', 'Total')
$row++
foreach ($g in ($errorsAndWarnings | Group-Object Problem | Sort-Object Count -Descending)) {
    $ws.Cells[$row, 1].Value = $g.Name
    $ws.Cells[$row, 2].Formula = "COUNTIFS($probCol,`$A$row,$sevCol,""*Error"")"
    $ws.Cells[$row, 3].Formula = "COUNTIFS($probCol,`$A$row,$sevCol,""*Warning"")"
    $ws.Cells[$row, 4].Formula = "B$row+C$row"
    $ws.Cells[$row, 4].Style.Font.Bold = $true
    $row++
}
$row++

# per computer
Add-BlockTitle -Worksheet $ws -Row $row -Text 'Per computer'
$row++
Add-BlockHeader -Worksheet $ws -Row $row -Headers @('ComputerName', 'Errors', 'Warnings', 'Total')
$row++
foreach ($g in ($errorsAndWarnings | Group-Object ComputerName | Sort-Object Count -Descending)) {
    $label = if ([string]::IsNullOrEmpty($g.Name)) { '(no computer / file level issue)' } else { $g.Name }
    $criteria = """$($g.Name)"""
    $ws.Cells[$row, 1].Value = $label
    $ws.Cells[$row, 2].Formula = "COUNTIFS($compCol,$criteria,$sevCol,""*Error"")"
    $ws.Cells[$row, 3].Formula = "COUNTIFS($compCol,$criteria,$sevCol,""*Warning"")"
    $ws.Cells[$row, 4].Formula = "B$row+C$row"
    $ws.Cells[$row, 4].Style.Font.Bold = $true
    $row++
}
Set-ColumnWidths -Worksheet $ws -Widths @(56, 12, 12, 12, 20)

# ---- sheet 4: notes ------------------------------------------------------
$wsNotes = $pkg.Workbook.Worksheets.Add('Notes & method')
$wsNotes.Cells.Style.Font.Name = 'Arial'
$wsNotes.Cells.Style.Font.Size = 10
$wsNotes.Cells['A1'].Value = 'How this workbook was built'
$wsNotes.Cells['A1'].Style.Font.Bold = $true
$wsNotes.Cells['A1'].Style.Font.Size = 13
$wsNotes.Cells['A1'].Style.Font.Color.SetColor($navy)

$notes = @(
    @('Generated', "$(Get-Date -Format 'dd/MM/yyyy HH:mm:ss') by $($MyInvocation.MyCommand.Name) on $env:COMPUTERNAME"),
    @('Source', "$root - $($runs.Count) matrix file runs in $($runFolders.Count) run folders"),
    @('Sheet 1', 'One row per logged error or warning, newest first. Information level entries are not on this sheet.'),
    @('Sheet 3', 'Same rows plus Information level entries and the full context (description, path, site, run folder, setting id, detail file).'),
    @('Two log formats', 'Old runs use "ID <n> - Settings.html", newer runs use "00 - Execution Report.html". The LogFormat column says which one a row came from.'),
    @('DateTime', 'New format rows use the exact timestamp from the problem detail JSON when the report links to one; otherwise the start time of the setting (old format) or of the run is used. See the TimestampSource column.'),
    @('ComputerName', 'Taken from the setting a problem belongs to. File level issues (for example "Excel File: Runspace processing failed") are not tied to a computer, so the cell is empty.'),
    @('System issues', 'SystemErrors.json / "System errors log.json" are run level, so they show as MatrixFileName "(system)" and are prefixed "[System]" on sheet 1.'),
    @('Folders without a log', "$foldersWithoutLog matrix folders contained no log file and are counted as 'no log'."),
    @('Coverage caveat', 'Only the matrix folders present under the log root can be reported on. If the run mails mention more matrix files than there are folders, the missing ones are not represented here.')
)
$r = 3
foreach ($n in $notes) {
    $wsNotes.Cells[$r, 1].Value = $n[0]
    $wsNotes.Cells[$r, 1].Style.Font.Bold = $true
    $wsNotes.Cells[$r, 2].Value = $n[1]
    $wsNotes.Cells[$r, 2].Style.WrapText = $true
    $wsNotes.Row($r).Height = 46
    $r++
}
$wsNotes.Column(1).Width = 26
$wsNotes.Column(2).Width = 110

# ---- order the sheets ----------------------------------------------------
$pkg.Workbook.Worksheets.MoveAfter('Summary', 'Errors & Warnings')
$pkg.Workbook.Worksheets.MoveToEnd('Notes & method')
$pkg.Workbook.Worksheets['Errors & Warnings'].View.TabSelected = $true

# EPPlus has no full calculation engine for every function; Excel recalculates
# the summary formulas when the file is opened. Try anyway so tools that read
# cached values see numbers.
try { 
    $pkg.Workbook.Calculate() 
}
catch { 
    Write-Verbose "Workbook.Calculate skipped: $($_.Exception.Message)" 
}

Close-ExcelPackage $pkg
Write-Host "Workbook written to $OutputFile" -ForegroundColor Green

#endregion
