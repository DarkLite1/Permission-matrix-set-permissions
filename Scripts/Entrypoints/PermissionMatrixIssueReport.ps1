#Requires -Version 7.0
<#
.SYNOPSIS
    Builds an Excel overview of all errors and warnings found in the
    "Permission matrix set permissions" log folders.

.DESCRIPTION
    Walks every run folder (e.g. "2026_07_26_220004 (BNL Nightly)"), reads the
    per-matrix-file logs and produces a workbook with four sheets:

        Errors & Warnings      MatrixFileName / DateTime / Type / Name / Description /
                               ComputerName
        Summary                totals and breakdowns (COUNTIFS formulas)
        Performance            processing time per matrix file and per run
        Performance by target  processing time per computer and per path
        All issues (detail)    same columns + Information level + full context
        Runs (per matrix file) one row per matrix file per run, with its duration
        Settings (per run)     one row per setting per run - computer, path, duration
        Notes & method         how the data was collected, with caveats

    The matrix file name on sheet 1, and the LogFile and DetailFile columns on
    the detail sheet, are hyperlinks: clicking one opens the log page that the
    row came from.

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

.PARAMETER LinkRoot
    Optional. Base path the hyperlinks should point at, when that is not the
    same as -LogRoot. Use it when the logs are read from a local copy but the
    workbook is shared with people who reach the logs over the network:

        -LogRoot D:\LogCopy -LinkRoot \\BELSGFRANIT07\Log\File or folder\...

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

    [string] $LinkRoot,

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

    # The words used in the report pills; the icon glyphs are filtered by shape.
    $script:TypeWords = @('ERROR', 'WARNING', 'INFORMATION', 'INFO')

    # Placeholder for issues that belong to the run rather than to a matrix file.
    # Scope lives here so that Type can stay a clean Error / Warning / Information:
    # the run level log is called SystemErrors.json but records all three. Defined
    # in the parser so the parallel runspaces get it too.
    $script:RunLevelName = '(run level)'

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
            if ($script:TypeWords -contains $t.ToUpperInvariant()) { continue }
            # icon glyphs such as the cross, warning triangle, bullet, check mark
            if ($t.Length -le 2 -and $t -notmatch '[0-9A-Za-z]') { continue }
            if (-not $blocks.Contains($t)) { $blocks.Add($t) }
        }
        return $blocks
    }

    function Get-IssueType {
        <# The pill text of a problem row -> Error / Warning / Information. #>
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

    function Convert-Duration {
        <# "00:04:16" -> [timespan]. Returns zero when it cannot be read. #>
        param([string] $Text)
        $result = [timespan]::Zero
        if ([timespan]::TryParse($Text, [System.Globalization.CultureInfo]::InvariantCulture,
                [ref] $result)) {
            return $result
        }
        return [timespan]::Zero
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

    function Get-DetailFilePath {
        <# The report links to a detail file with a UNC or relative href; this
           resolves it to a real path in the matrix folder, or '' when it is
           not there. #>
        param([string] $Folder, [string] $Href)
        if ([string]::IsNullOrEmpty($Href)) { return '' }
        $path = Join-Path $Folder (($Href -split '[\\/]')[-1])
        if (Test-Path -LiteralPath $path -PathType Leaf) { return $path }
        return ''
    }

    function Get-DetailFileDateTime {
        <# Exact timestamp of a single problem, from its detail JSON. #>
        param([string] $Path)
        if ([string]::IsNullOrEmpty($Path)) { return $null }
        try {
            $head = Read-TextFile $Path
            if ($head.Length -gt 600) { $head = $head.Substring(0, 600) }
            $m = [regex]::Match($head, '"DateTime"\s*:\s*"([^"]+)"')
            if ($m.Success) { return Convert-IsoDate $m.Groups[1].Value }
        }
        catch {
            Write-Verbose "Could not read detail file '$Path': $($_.Exception.Message)"
        }
        return $null
    }

    function New-Issue {
        param(
            [string] $MatrixFileName, $DateTime, [string] $Type, [string] $Name,
            [string] $Description, [string] $ComputerName, [string] $Path, [string] $Site,
            [string] $SettingId, [string] $RunFolder, [string] $LogFormat,
            [string] $DetailFile, [string] $TimestampSource,
            [string] $LogFile, [string] $DetailPath
        )
        return [pscustomobject] @{
            MatrixFileName  = $MatrixFileName
            DateTime        = $DateTime
            Type            = $Type
            Name            = $Name
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
            LogFile         = $LogFile
            DetailPath      = $DetailPath
        }
    }

    function New-Setting {
        <# One row per setting per run: this is the grain where ComputerName, Path
           and the duration live, so it is what the per target performance is
           built on. #>
        param(
            [string] $RunFolder, [string] $MatrixFileName, [string] $SettingId,
            [string] $ComputerName, [string] $Path, [string] $Site, [string] $Action,
            [timespan] $Duration, $StartTime, [string] $LogFormat, [string] $LogFile
        )
        return [pscustomobject] @{
            RunFolder      = $RunFolder
            RunType        = (Get-RunType $RunFolder)
            MatrixFileName = $MatrixFileName
            SettingId      = $SettingId
            ComputerName   = $ComputerName
            Path           = $Path
            Site           = $Site
            Action         = $Action
            Duration       = $Duration
            StartTime      = $StartTime
            Errors         = 0
            Warnings       = 0
            Information    = 0
            LogFormat      = $LogFormat
            LogFile        = $LogFile
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
            [System.Collections.Generic.List[object]] $Issues,
            [System.Collections.Generic.List[object]] $Settings)

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
        $current = $null
        $settingCount = 0
        # the footer Start/End time is the window of the whole run, identical in every
        # report of that run, so the only per matrix file timing is the clock value in
        # each settings card - summed here into the processing time of this matrix file
        $work = [timespan]::Zero
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
                $action = ''
                $m = [regex]::Match($chunk, '>Action:</span>\s*<span[^>]*>([^<]*)</span>')
                if ($m.Success) { $action = Get-CleanText $m.Groups[1].Value }
                $duration = [timespan]::Zero
                $m = [regex]::Match($chunk, '>(\d{2}:\d{2}:\d{2})<')
                if ($m.Success) { $duration = Convert-Duration $m.Groups[1].Value }
                $work += $duration

                # one row per setting; the problem rows that follow belong to it and
                # bump its counters, so the object is kept as $current
                $current = New-Setting -RunFolder $RunFolder -MatrixFileName $matrixName `
                    -SettingId $settingId -ComputerName $computer -Path $path -Site $site `
                    -Action $action -Duration $duration -StartTime $startTime `
                    -LogFormat 'new' -LogFile $ReportPath
                $Settings.Add($current)
                continue
            }

            # ---- problem row ---------------------------------------------
            $pill = [regex]::Match($chunk, '>\s*(ERROR|WARNING|INFORMATION|INFO)\s*<')
            if ($pill.Success) {
                $type = Get-IssueType $pill.Groups[1].Value
                # cut at the '<' that opens the pill so its own tag cannot leak text
                $cut = $chunk.LastIndexOf('<', $pill.Index)
                if ($cut -lt 0) { $cut = $pill.Index }
                $content = $chunk.Substring(0, $cut)
            }
            else {
                $type = 'Unknown'
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

            $detailPath = Get-DetailFilePath -Folder $folder -Href $href
            $stamp = Get-DetailFileDateTime -Path $detailPath
            $stampSource = 'problem detail file'
            if ($null -eq $stamp) { $stamp = $startTime; $stampSource = 'run start time' }

            $counts[$type] = $counts[$type] + 1
            if ($null -ne $current) {
                switch ($type) {
                    'Error' { $current.Errors++ }
                    'Warning' { $current.Warnings++ }
                    'Information' { $current.Information++ }
                }
            }
            $Issues.Add((New-Issue -MatrixFileName $matrixName -DateTime $stamp `
                        -Type $type -Name $problem -Description $description `
                        -ComputerName $computer -Path $path -Site $site -SettingId $settingId `
                        -RunFolder $RunFolder -LogFormat 'new' -DetailFile $href `
                        -TimestampSource $stampSource `
                        -LogFile $ReportPath -DetailPath $detailPath))
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
            Duration       = $work
            StartTime      = $startTime
            EndTime        = $endTime
            LogFormat      = 'new'
            LogFile        = $ReportPath
        }
    }

    # ---------------------------------------------------- old format parser ---

    function Read-SettingsReports {
        <# Parses the "ID <n> - Settings.html" files of one matrix folder (runs up
           to 2026-06-29). Each file is one setting = one computer, and carries its
           own start/end time, which is the best timestamp available in this
           format - the detail .txt files hold no DateTime. #>
        param([string] $MatrixDir, [string] $RunFolder, [string] $MatrixFolder,
            [System.Collections.Generic.List[object]] $Issues,
            [System.Collections.Generic.List[object]] $Settings)

        $matrixName = "$MatrixFolder.xlsx"
        $settingFiles = @(Get-ChildItem -LiteralPath $MatrixDir -Filter 'ID * - Settings.html' -File |
            Sort-Object Name)

        $errors = 0; $warnings = 0; $information = 0
        $work = [timespan]::Zero
        $starts = [System.Collections.Generic.List[datetime]]::new()
        $ends = [System.Collections.Generic.List[datetime]]::new()

        # the troubleshooting log is the overview page for the whole matrix file,
        # so it is the better link target when it exists
        $overview = Join-Path $MatrixDir '00 - Troubleshooting Log.html'
        if (-not (Test-Path -LiteralPath $overview -PathType Leaf)) { $overview = '' }

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
            $site = ''
            $m = [regex]::Match($html, '<th>\s*SiteCode\s*</th>\s*<td>(.*?)</td>', $script:RxOptions)
            if ($m.Success) { $site = Get-CleanText $m.Groups[1].Value }
            $stampSource = 'setting start time'
            if ($null -eq $startTime) {
                $startTime = Get-RunFolderDate $RunFolder
                $stampSource = 'run start time'
            }

            # settings summary row: <marker cell><ID><ComputerName><Path><Action><Duration>
            # the marker cell is id="probTypeError|Warning|Info" when the setting has a
            # problem and id="" when it does not, so match on shape, not on the id
            $computer = ''; $path = ''; $settingId = ''; $action = ''
            $duration = [timespan]::Zero
            $m = [regex]::Match($html,
                '<tr>\s*<td id="[^"]*"[^>]*>\s*</td>\s*<td>\s*(\d+)\s*</td>(.*?)</tr>',
                $script:RxOptions)
            if ($m.Success) {
                $settingId = Get-CleanText $m.Groups[1].Value
                # remaining cells: ComputerName, Path, Action, Duration
                $cells = @([regex]::Matches($m.Groups[2].Value, '<td[^>]*>(.*?)</td>', $script:RxOptions) |
                    ForEach-Object { Get-CleanText $_.Groups[1].Value })
                if ($cells.Count -gt 0) { $computer = $cells[0] }
                if ($cells.Count -gt 1) { $path = $cells[1] }
                if ($cells.Count -gt 2) { $action = $cells[2] }
                if ($cells.Count -gt 3) { $duration = Convert-Duration $cells[3] }
            }
            $work += $duration
            if ($settingId -eq '') {
                $m = [regex]::Match($file.Name, 'ID (\d+) - Settings')
                if ($m.Success) { $settingId = $m.Groups[1].Value }
            }

            $setting = New-Setting -RunFolder $RunFolder -MatrixFileName $matrixName `
                -SettingId $settingId -ComputerName $computer -Path $path -Site $site `
                -Action $action -Duration $duration -StartTime $startTime `
                -LogFormat 'old' -LogFile $file.FullName
            $Settings.Add($setting)

            # problem rows
            $rx = '<tr>\s*<td id="probType(Error|Warning|Info)"[^>]*>\s*</td>\s*<td colspan="7">(.*?)</td>\s*</tr>'
            foreach ($pm in [regex]::Matches($html, $rx, $script:RxOptions)) {
                $type = Get-IssueType $pm.Groups[1].Value
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
                $detailPath = Get-DetailFilePath -Folder $MatrixDir -Href $detail

                switch ($type) {
                    'Error' { $errors++; $setting.Errors++ }
                    'Warning' { $warnings++; $setting.Warnings++ }
                    'Information' { $information++; $setting.Information++ }
                }

                $Issues.Add((New-Issue -MatrixFileName $matrixName -DateTime $startTime `
                            -Type $type -Name $problem -Description $description `
                            -ComputerName $computer -Path $path -Site $site -SettingId $settingId `
                            -RunFolder $RunFolder -LogFormat 'old' -DetailFile $detail `
                            -TimestampSource $stampSource `
                            -LogFile $file.FullName -DetailPath $detailPath))
            }
        }

        $status = if ($errors -gt 0) { 'Error' } elseif ($warnings -gt 0) { 'Warning' } else { 'Success' }
        $runStart = if ($starts.Count -gt 0) { ($starts | Measure-Object -Minimum).Minimum } else { Get-RunFolderDate $RunFolder }
        $runEnd = if ($ends.Count -gt 0) { ($ends | Measure-Object -Maximum).Maximum } else { $null }
        $logFile = if ($overview -ne '') { $overview } elseif ($settingFiles.Count -gt 0) { $settingFiles[0].FullName } else { '' }

        return [pscustomobject] @{
            RunFolder      = $RunFolder
            RunType        = (Get-RunType $RunFolder)
            MatrixFileName = $matrixName
            Status         = $status
            Settings       = $settingFiles.Count
            Errors         = $errors
            Warnings       = $warnings
            Information    = $information
            Duration       = $work
            StartTime      = $runStart
            EndTime        = $runEnd
            LogFormat      = 'old'
            LogFile        = $logFile
        }
    }

    # ------------------------------------------------ run level system logs ---

    function Read-SystemErrorLog {
        <# "SystemErrors.json" (new) / "System errors log.json" (old). These are
           run level: no matrix file and no computer, so they get MatrixFileName
           "(run level)". Despite the file name they are not all errors - the Type
           field in the JSON decides, and it also carries Warning and Information. #>
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

                $Issues.Add((New-Issue -MatrixFileName $script:RunLevelName -DateTime $stamp `
                            -Type (Get-IssueType $type) -Name $problem `
                            -Description $message -ComputerName '' -Path '' -Site '' -SettingId '' `
                            -RunFolder $RunFolder -LogFormat 'system' -DetailFile $name `
                            -TimestampSource $stampSource `
                            -LogFile $path -DetailPath $path))
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
        $settings = [System.Collections.Generic.List[object]]::new()
        $withoutLog = 0

        Read-SystemErrorLog -RunDir $RunDir -RunFolder $RunFolder -Issues $issues

        foreach ($matrixDir in (Get-ChildItem -LiteralPath $RunDir -Directory | Sort-Object Name)) {
            if ($matrixDir.Name -eq 'Export') { continue }

            $newReport = Join-Path $matrixDir.FullName '00 - Execution Report.html'
            $oldSettings = @(Get-ChildItem -LiteralPath $matrixDir.FullName -Filter 'ID * - Settings.html' -File)

            if (Test-Path -LiteralPath $newReport -PathType Leaf) {
                $runs.Add((Read-ExecutionReport -ReportPath $newReport -RunFolder $RunFolder `
                            -MatrixFolder $matrixDir.Name -Issues $issues -Settings $settings))
            }
            elseif ($oldSettings.Count -gt 0) {
                $runs.Add((Read-SettingsReports -MatrixDir $matrixDir.FullName -RunFolder $RunFolder `
                            -MatrixFolder $matrixDir.Name -Issues $issues -Settings $settings))
            }
            else {
                $withoutLog++
                $runs.Add([pscustomobject] @{
                        RunFolder = $RunFolder; RunType = (Get-RunType $RunFolder)
                        MatrixFileName = "$($matrixDir.Name).xlsx"; Status = 'no log'
                        Settings = 0; Errors = 0; Warnings = 0; Information = 0
                        Duration = [timespan]::Zero
                        StartTime = (Get-RunFolderDate $RunFolder); EndTime = $null
                        LogFormat = 'none'; LogFile = ''
                    })
            }
        }

        Write-Verbose "$RunFolder - $($runs.Count) matrix files, $($issues.Count) issues"
        return [pscustomobject] @{
            Issues            = $issues.ToArray()
            Runs              = $runs.ToArray()
            Settings          = $settings.ToArray()
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
$settingRows = [System.Collections.Generic.List[object]]::new()
$foldersWithoutLog = 0
foreach ($r in $results) {
    if ($r.Issues) { $issues.AddRange([object[]] $r.Issues) }
    if ($r.Runs) { $runs.AddRange([object[]] $r.Runs) }
    if ($r.Settings) { $settingRows.AddRange([object[]] $r.Settings) }
    $foldersWithoutLog += $r.FoldersWithoutLog
}

# -Stable keeps rows with an identical timestamp in a predictable order, so two
# runs of this script over the same logs produce byte-identical sheets.
$allIssues = @(
    $issues | Sort-Object `
        -Stable `
        -Property @{ Expression = 'DateTime'; Descending = $true }, 
    MatrixFileName, ComputerName, Name
)
$runs = @($runs | Sort-Object -Stable RunFolder, MatrixFileName)
$settingRows = @($settingRows | Sort-Object -Stable RunFolder, MatrixFileName, SettingId)

$errorsAndWarnings = @(
    $allIssues | Where-Object { 
        $_.Type -eq 'Error' -or $_.Type -eq 'Warning' 
    }
)

# split by scope: a run level issue has no matrix file to blame
$matrixIssues = @($allIssues | Where-Object MatrixFileName -NE $script:RunLevelName)
$runLevelIssues = @($allIssues | Where-Object MatrixFileName -EQ $script:RunLevelName)

$countError = @($matrixIssues | Where-Object Type -EQ 'Error').Count
$countWarning = @($matrixIssues | Where-Object Type -EQ 'Warning').Count
$countRunError = @($runLevelIssues | Where-Object Type -EQ 'Error').Count
$countRunWarning = @($runLevelIssues | Where-Object Type -EQ 'Warning').Count
$countInformation = @($allIssues | Where-Object Type -EQ 'Information').Count

Write-Host ('Matrix file runs : {0}   (new format {1}, old format {2}, no log {3})   in {4:n1}s' -f
    $runs.Count,
    @($runs | Where-Object LogFormat -EQ 'new').Count,
    @($runs | Where-Object LogFormat -EQ 'old').Count,
    $foldersWithoutLog,
    $stopwatch.Elapsed.TotalSeconds)
Write-Host ('Errors {0} (+{1} run level)   Warnings {2} (+{3} run level)   Information {4}' -f
    $countError, $countRunError, $countWarning, $countRunWarning, $countInformation)

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

    $settingRows | Export-Csv `
        -LiteralPath (Join-Path $CsvFolder 'Settings.csv') `
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
$linkBlue = [System.Drawing.Color]::FromArgb(5, 99, 193)
$solid = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$dateFormat = 'dd/mm/yyyy hh:mm:ss'
$durationFormat = '[h]:mm:ss'

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
        $color = if ($Rows[$i].Type -eq 'Error') { $script:redBg }
        elseif ($Rows[$i].Type -eq 'Warning') { $script:amber }
        else { $script:grey }
        $cells.Style.Fill.BackgroundColor.SetColor($color)
    }
}

function Convert-LinkPath {
    <# Rewrites a path so the links point at -LinkRoot instead of the folder the
       logs were actually read from. #>
    param([string] $Path)
    if (-not $LinkRoot) { return $Path }
    if (-not $Path.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) { return $Path }
    $rest = $Path.Substring($root.Length)
    if ($LinkRoot.Contains('\')) { $rest = $rest -replace '/', '\' }
    return $LinkRoot.TrimEnd('\', '/') + $rest
}

function Set-CellLink {
    <# Turns an existing cell into a hyperlink without changing what it shows.
       Setting .Hyperlink on an empty cell would overwrite the text, hence the
       save and restore. A cell that shows the path itself follows -LinkRoot. #>
    param($Worksheet, [int] $Row, [int] $Column, [string] $Target)
    if ([string]::IsNullOrEmpty($Target)) { return }
    $cell = $Worksheet.Cells[$Row, $Column]
    $linkTarget = Convert-LinkPath $Target
    $text = $cell.Value
    if ($text -is [string] -and $text -eq $Target) { $text = $linkTarget }
    try {
        $cell.Hyperlink = [System.Uri]::new($linkTarget)
        $cell.Value = $text
        $cell.Style.Font.UnderLine = $true
        $cell.Style.Font.Color.SetColor($script:linkBlue)
    }
    catch {
        Write-Verbose "Could not link '$linkTarget': $($_.Exception.Message)"
    }
}

# ---- sheet 1: errors and warnings, one Type column ------------------------
$mainRows = $errorsAndWarnings |
Select-Object MatrixFileName, DateTime, Type, Name, Description, ComputerName

$pkg = $mainRows | 
Export-Excel -Path $OutputFile `
    -WorksheetName 'Errors & Warnings' `
    -AutoFilter -FreezeTopRow -PassThru

# ---- sheet 2: full detail ------------------------------------------------
$pkg = $allIssues |
Select-Object MatrixFileName, DateTime, Type, Name, Description,
ComputerName, Path, Site, SettingId, RunFolder, RunType,
LogFormat, DetailFile, TimestampSource, LogFile |
Export-Excel -ExcelPackage $pkg -WorksheetName 'All issues (detail)' `
    -AutoFilter -FreezeTopRow -PassThru

# ---- styling -------------------------------------------------------------
$wsMain = $pkg.Workbook.Worksheets['Errors & Warnings']
$wsMain.Cells.Style.Font.Name = 'Arial'
$wsMain.Cells.Style.Font.Size = 10
Set-HeaderStyle -Worksheet $wsMain -Range 'A1:F1'
Set-ColumnWidths -Worksheet $wsMain -Widths @(42, 19, 13, 36, 70, 18)
$wsMain.Column(2).Style.Numberformat.Format = $dateFormat
$wsMain.Cells['D:E'].Style.WrapText = $true
Set-RowFills -Worksheet $wsMain -Rows $errorsAndWarnings -LastColumn 'F'

# the matrix file name opens the log page the row came from
for ($i = 0; $i -lt $errorsAndWarnings.Count; $i++) {
    Set-CellLink -Worksheet $wsMain -Row ($i + 2) -Column 1 -Target $errorsAndWarnings[$i].LogFile
}

$wsDetail = $pkg.Workbook.Worksheets['All issues (detail)']
$wsDetail.Cells.Style.Font.Name = 'Arial'
$wsDetail.Cells.Style.Font.Size = 10
Set-HeaderStyle -Worksheet $wsDetail -Range 'A1:O1'
Set-ColumnWidths -Worksheet $wsDetail -Widths @(42, 19, 14, 38, 70, 18, 46, 16, 38, 34, 13, 11, 46, 22, 80)
$wsDetail.Column(2).Style.Numberformat.Format = $dateFormat
Set-RowFills -Worksheet $wsDetail -Rows $allIssues -LastColumn 'O'

# DetailFile (M) opens the single problem, LogFile (O) opens the whole report
for ($i = 0; $i -lt $allIssues.Count; $i++) {
    $row = $i + 2
    Set-CellLink -Worksheet $wsDetail -Row $row -Column 13 -Target $allIssues[$i].DetailPath
    Set-CellLink -Worksheet $wsDetail -Row $row -Column 15 -Target $allIssues[$i].LogFile
}

# ---- sheet 3: summary, everything by formula ------------------------------
$detailLast = $allIssues.Count + 1
$det = "'All issues (detail)'"
$typeCol = "$det!`$C`$2:`$C`$$detailLast"
$matCol = "$det!`$A`$2:`$A`$$detailLast"
$nameCol = "$det!`$D`$2:`$D`$$detailLast"
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
$runLevel = """$script:RunLevelName"""
$notRunLevel = """<>$script:RunLevelName"""
foreach ($pair in @(
        @('Errors (matrix files)', "COUNTIFS($typeCol,""Error"",$matCol,$notRunLevel)"),
        @('Warnings (matrix files)', "COUNTIFS($typeCol,""Warning"",$matCol,$notRunLevel)"),
        @('Errors (run level)', "COUNTIFS($typeCol,""Error"",$matCol,$runLevel)"),
        @('Warnings (run level)', "COUNTIFS($typeCol,""Warning"",$matCol,$runLevel)"),
        @('Information (not an issue)', "COUNTIF($typeCol,""Information"")"))) {
    $ws.Cells[$row, 1].Value = $pair[0]
    $ws.Cells[$row, 2].Formula = $pair[1]
    $ws.Cells[$row, 2].Style.Font.Bold = $true
    $row++
}
$ws.Cells[$row, 1].Value = 'Errors + warnings (rows on sheet 1)'
$ws.Cells[$row, 2].Formula = "COUNTIF($typeCol,""Error"")+COUNTIF($typeCol,""Warning"")"
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
    $ws.Cells[$row, 2].Formula = "COUNTIFS($matCol,`$A$row,$typeCol,""Error"")"
    $ws.Cells[$row, 3].Formula = "COUNTIFS($matCol,`$A$row,$typeCol,""Warning"")"
    $ws.Cells[$row, 4].Formula = "B$row+C$row"
    $ws.Cells[$row, 4].Style.Font.Bold = $true
    # SUMPRODUCT keeps this a normal formula - no Ctrl+Shift+Enter needed
    $ws.Cells[$row, 5].Formula = "SUMPRODUCT(MAX(($matCol=`$A$row)*$dtCol))"
    $ws.Cells[$row, 5].Style.Numberformat.Format = $dateFormat
    # the name links to the most recent log for that matrix file
    $newest = @($g.Group | Sort-Object DateTime -Descending)[0]
    Set-CellLink -Worksheet $ws -Row $row -Column 1 -Target $newest.LogFile
    $row++
}
$row++

# per issue name
Add-BlockTitle -Worksheet $ws -Row $row -Text 'Per issue name'
$row++
Add-BlockHeader -Worksheet $ws -Row $row -Headers @('Name', 'Errors', 'Warnings', 'Total')
$row++
foreach ($g in ($errorsAndWarnings | Group-Object Name | Sort-Object Count -Descending)) {
    $ws.Cells[$row, 1].Value = $g.Name
    $ws.Cells[$row, 2].Formula = "COUNTIFS($nameCol,`$A$row,$typeCol,""Error"")"
    $ws.Cells[$row, 3].Formula = "COUNTIFS($nameCol,`$A$row,$typeCol,""Warning"")"
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
    $ws.Cells[$row, 2].Formula = "COUNTIFS($compCol,$criteria,$typeCol,""Error"")"
    $ws.Cells[$row, 3].Formula = "COUNTIFS($compCol,$criteria,$typeCol,""Warning"")"
    $ws.Cells[$row, 4].Formula = "B$row+C$row"
    $ws.Cells[$row, 4].Style.Font.Bold = $true
    $row++
}
Set-ColumnWidths -Worksheet $ws -Widths @(56, 12, 12, 12, 20)

# ---- sheet: per matrix file run data (feeds the performance sheet) --------
$runsSheet = 'Runs (per matrix file)'
$runRows = $runs |
Select-Object RunFolder, RunType, MatrixFileName, Status, Settings,
@{ Name = 'Duration'; Expression = { $_.Duration.TotalDays } },
StartTime, EndTime, Errors, Warnings, Information, LogFormat, LogFile

$pkg = $runRows | Export-Excel -ExcelPackage $pkg -WorksheetName $runsSheet `
    -AutoFilter -FreezeTopRow -PassThru

$wsRuns = $pkg.Workbook.Worksheets[$runsSheet]
$wsRuns.Cells.Style.Font.Name = 'Arial'
$wsRuns.Cells.Style.Font.Size = 10
Set-HeaderStyle -Worksheet $wsRuns -Range 'A1:M1'
Set-ColumnWidths -Worksheet $wsRuns -Widths @(34, 13, 42, 22, 9, 11, 19, 19, 8, 10, 12, 11, 80)
$wsRuns.Column(6).Style.Numberformat.Format = $durationFormat
$wsRuns.Column(7).Style.Numberformat.Format = $dateFormat
$wsRuns.Column(8).Style.Numberformat.Format = $dateFormat

# only the rows that had a problem get a colour, so the sheet stays readable
for ($i = 0; $i -lt $runs.Count; $i++) {
    $row = $i + 2
    $color = if ($runs[$i].Errors -gt 0) { $redBg } elseif ($runs[$i].Warnings -gt 0) { $amber } else { $null }
    if ($color) {
        $cells = $wsRuns.Cells["A$row" + ":M$row"]
        $cells.Style.Fill.PatternType = $solid
        $cells.Style.Fill.BackgroundColor.SetColor($color)
    }
    Set-CellLink -Worksheet $wsRuns -Row $row -Column 3 -Target $runs[$i].LogFile
}

# ---- sheet: performance --------------------------------------------------
$rs = "'$runsSheet'"
$rLast = $runs.Count + 1
$rFolder = "$rs!`$A`$2:`$A`$$rLast"
$rMatrix = "$rs!`$C`$2:`$C`$$rLast"
$rSettings = "$rs!`$E`$2:`$E`$$rLast"
$rDuration = "$rs!`$F`$2:`$F`$$rLast"
$rStart = "$rs!`$G`$2:`$G`$$rLast"
$rEnd = "$rs!`$H`$2:`$H`$$rLast"
$rErrors = "$rs!`$I`$2:`$I`$$rLast"
$rWarnings = "$rs!`$J`$2:`$J`$$rLast"

$perf = $pkg.Workbook.Worksheets.Add('Performance')
$perf.Cells.Style.Font.Name = 'Arial'
$perf.Cells.Style.Font.Size = 10
$perf.Cells['A1'].Value = 'Permission matrix - processing time'
$perf.Cells['A1'].Style.Font.Bold = $true
$perf.Cells['A1'].Style.Font.Size = 13
$perf.Cells['A1'].Style.Font.Color.SetColor($navy)
$perf.Cells['A2'].Value = "Duration is the sum of the per setting durations of that matrix file, so it is processing time, not elapsed time. Counts and averages are formulas over the '$runsSheet' sheet."
$perf.Cells['A2'].Style.Font.Size = 9
$perf.Cells['A2'].Style.Font.Color.SetColor([System.Drawing.Color]::Gray)

# average duration per matrix file decides the order; the values are formulas
$byDuration = foreach ($g in ($runs | Group-Object MatrixFileName)) {
    $timed = @($g.Group | Where-Object { $_.Duration.TotalSeconds -gt 0 })
    $average = 0.0
    if ($timed.Count -gt 0) {
        $total = 0.0
        foreach ($x in $timed) { $total += $x.Duration.TotalSeconds }
        $average = $total / $timed.Count
    }
    [pscustomobject] @{ Name = $g.Name; Average = $average; Newest = @($g.Group | Sort-Object StartTime -Descending)[0] }
}
$byDuration = @($byDuration | Sort-Object Average -Descending)

$row = 4
Add-BlockTitle -Worksheet $perf -Row $row -Text "Per matrix file ($($byDuration.Count) matrix files, slowest first)"
$row++
Add-BlockHeader -Worksheet $perf -Row $row -Headers @('MatrixFileName', 'Runs', 'Avg settings',
    'Avg duration', 'Shortest', 'Longest', 'Total time', 'Last run', 'Last duration')
$row++
foreach ($item in $byDuration) {
    $perf.Cells[$row, 1].Value = $item.Name
    $perf.Cells[$row, 2].Formula = "COUNTIF($rMatrix,`$A$row)"
    $perf.Cells[$row, 3].Formula = "AVERAGEIFS($rSettings,$rMatrix,`$A$row)"
    $perf.Cells[$row, 3].Style.Numberformat.Format = '0.0'
    # IFERROR covers a matrix file whose every run has a zero duration
    $perf.Cells[$row, 4].Formula = "IFERROR(AVERAGEIFS($rDuration,$rMatrix,`$A$row,$rDuration,"">0""),0)"
    # MAX can use the multiplication trick because the non matching rows become 0 and
    # 0 never wins a MAX of positive values. MIN needs the non matching rows removed
    # instead of zeroed, which takes an IF - and MIN(IF(..)) only evaluates as an
    # array formula, so it is written as one. SUMPRODUCT(MIN(IF(..))) returns #VALUE!
    # in Excel even though LibreOffice accepts it.
    $perf.Cells[$row, 5].CreateArrayFormula("MIN(IF(($rMatrix=`$A$row)*($rDuration>0),$rDuration))")
    $perf.Cells[$row, 6].Formula = "SUMPRODUCT(MAX(($rMatrix=`$A$row)*$rDuration))"
    $perf.Cells[$row, 7].Formula = "SUMIFS($rDuration,$rMatrix,`$A$row)"
    $perf.Cells[$row, 8].Formula = "SUMPRODUCT(MAX(($rMatrix=`$A$row)*$rStart))"
    $perf.Cells[$row, 8].Style.Numberformat.Format = $dateFormat
    $perf.Cells[$row, 9].Formula = "SUMIFS($rDuration,$rMatrix,`$A$row,$rStart,`$H$row)"
    foreach ($c in 4, 5, 6, 7, 9) { $perf.Cells[$row, $c].Style.Numberformat.Format = $durationFormat }
    Set-CellLink -Worksheet $perf -Row $row -Column 1 -Target $item.Newest.LogFile
    $row++
}
$row++

# per run
$byRun = @($runs | Group-Object RunFolder | Sort-Object Name -Descending)
Add-BlockTitle -Worksheet $perf -Row $row -Text "Per run ($($byRun.Count) runs, newest first)"
$row++
Add-BlockHeader -Worksheet $perf -Row $row -Headers @('RunFolder', 'RunType', 'Matrix files',
    'First start', 'Last end', 'Elapsed', 'Processing time', 'Errors', 'Warnings')
$row++
foreach ($g in $byRun) {
    $perf.Cells[$row, 1].Value = $g.Name
    $perf.Cells[$row, 2].Value = $g.Group[0].RunType
    $perf.Cells[$row, 3].Formula = "COUNTIF($rFolder,`$A$row)"
    $perf.Cells[$row, 4].CreateArrayFormula("MIN(IF($rFolder=`$A$row,$rStart))")
    $perf.Cells[$row, 5].Formula = "SUMPRODUCT(MAX(($rFolder=`$A$row)*$rEnd))"
    $perf.Cells[$row, 6].Formula = "E$row-D$row"
    $perf.Cells[$row, 7].Formula = "SUMIFS($rDuration,$rFolder,`$A$row)"
    $perf.Cells[$row, 8].Formula = "SUMIFS($rErrors,$rFolder,`$A$row)"
    $perf.Cells[$row, 9].Formula = "SUMIFS($rWarnings,$rFolder,`$A$row)"
    $perf.Cells[$row, 4].Style.Numberformat.Format = $dateFormat
    $perf.Cells[$row, 5].Style.Numberformat.Format = $dateFormat
    $perf.Cells[$row, 6].Style.Numberformat.Format = $durationFormat
    $perf.Cells[$row, 7].Style.Numberformat.Format = $durationFormat
    $row++
}
Set-ColumnWidths -Worksheet $perf -Widths @(46, 13, 13, 19, 19, 13, 15, 10, 10)

# ---- sheet: per setting data (feeds the per target performance) -----------
$settingsSheet = 'Settings (per run)'
$settingSheetRows = $settingRows |
Select-Object RunFolder, RunType, MatrixFileName, SettingId, ComputerName, Path, Site, Action,
@{ Name = 'Duration'; Expression = { $_.Duration.TotalDays } },
StartTime, Errors, Warnings, Information, LogFormat, LogFile

$pkg = $settingSheetRows | Export-Excel -ExcelPackage $pkg -WorksheetName $settingsSheet `
    -AutoFilter -FreezeTopRow -PassThru

$wsSettings = $pkg.Workbook.Worksheets[$settingsSheet]
$wsSettings.Cells.Style.Font.Name = 'Arial'
$wsSettings.Cells.Style.Font.Size = 10
Set-HeaderStyle -Worksheet $wsSettings -Range 'A1:O1'
Set-ColumnWidths -Worksheet $wsSettings -Widths @(34, 13, 42, 38, 18, 46, 16, 10, 11, 19, 8, 10, 12, 11, 80)
$wsSettings.Column(9).Style.Numberformat.Format = $durationFormat
$wsSettings.Column(10).Style.Numberformat.Format = $dateFormat

for ($i = 0; $i -lt $settingRows.Count; $i++) {
    $color = if ($settingRows[$i].Errors -gt 0) { $redBg } elseif ($settingRows[$i].Warnings -gt 0) { $amber } else { $null }
    if ($color) {
        $row = $i + 2
        $cells = $wsSettings.Cells["A$row" + ":O$row"]
        $cells.Style.Fill.PatternType = $solid
        $cells.Style.Fill.BackgroundColor.SetColor($color)
    }
}

# ---- sheet: performance by computer and path -----------------------------
$ss = "'$settingsSheet'"
$sLast = $settingRows.Count + 1
$sComputer = "$ss!`$E`$2:`$E`$$sLast"
$sPath = "$ss!`$F`$2:`$F`$$sLast"
$sDuration = "$ss!`$I`$2:`$I`$$sLast"
$sErrors = "$ss!`$K`$2:`$K`$$sLast"
$sWarnings = "$ss!`$L`$2:`$L`$$sLast"

$target = $pkg.Workbook.Worksheets.Add('Performance by target')
$target.Cells.Style.Font.Name = 'Arial'
$target.Cells.Style.Font.Size = 10
$target.Cells['A1'].Value = 'Permission matrix - processing time per computer and per path'
$target.Cells['A1'].Style.Font.Bold = $true
$target.Cells['A1'].Style.Font.Size = 13
$target.Cells['A1'].Style.Font.Color.SetColor($navy)
$target.Cells['A2'].Value = "One setting is one computer plus one path, so this is where the time is actually spent. Figures are formulas over the '$settingsSheet' sheet."
$target.Cells['A2'].Style.Font.Size = 9
$target.Cells['A2'].Style.Font.Color.SetColor([System.Drawing.Color]::Gray)

function Get-AverageSeconds {
    param([object[]] $Rows)
    $timed = @($Rows | Where-Object { $_.Duration.TotalSeconds -gt 0 })
    if ($timed.Count -eq 0) { return 0.0 }
    $total = 0.0
    foreach ($x in $timed) { $total += $x.Duration.TotalSeconds }
    return $total / $timed.Count
}

function Add-TargetBlock {
    <# One block of per target figures. $Range is the column of the settings sheet
       to group on, $Extra adds a value only column (used for the computer a path
       sits on). #>
    param(
        $Worksheet, [int] $StartRow, [string] $Title, [string] $KeyHeader,
        [string] $Range, [object[]] $Groups, [switch] $WithComputer
    )
    $row = $StartRow
    Add-BlockTitle -Worksheet $Worksheet -Row $row -Text $Title
    $row++
    $headers = @($KeyHeader)
    if ($WithComputer) { $headers += 'ComputerName' }
    $headers += @('Settings run', 'Avg duration', 'Shortest', 'Longest', 'Total time', 'Errors', 'Warnings')
    Add-BlockHeader -Worksheet $Worksheet -Row $row -Headers $headers
    $row++
    $offset = if ($WithComputer) { 1 } else { 0 }
    foreach ($g in $Groups) {
        # the column numbers are worked out first on purpose: inside an index,
        # [$row, 2 + $offset] parses as ($row, 2) + $offset and builds a 3 element array
        $cCount = 2 + $offset
        $cAvg = 3 + $offset
        $cMin = 4 + $offset
        $cMax = 5 + $offset
        $cTotal = 6 + $offset
        $cErrors = 7 + $offset
        $cWarnings = 8 + $offset

        $Worksheet.Cells[$row, 1].Value = if ([string]::IsNullOrEmpty($g.Key)) { '(none)' } else { $g.Key }
        if ($WithComputer) { $Worksheet.Cells[$row, 2].Value = $g.Computer }
        # the criteria is the literal value, not a reference to column A: COUNTIF
        # against a reference to an empty cell reads it as 0 and matches nothing,
        # which is what the two settings without a computer name would hit
        $key = """$($g.Key)"""
        $Worksheet.Cells[$row, $cCount].Formula = "COUNTIF($Range,$key)"
        $Worksheet.Cells[$row, $cAvg].Formula = "IFERROR(AVERAGEIFS($script:sDurationRef,$Range,$key,$script:sDurationRef,"">0""),0)"
        $Worksheet.Cells[$row, $cMin].CreateArrayFormula(
            "MIN(IF(($Range=$key)*($script:sDurationRef>0),$script:sDurationRef))")
        $Worksheet.Cells[$row, $cMax].Formula = "SUMPRODUCT(MAX(($Range=$key)*$script:sDurationRef))"
        $Worksheet.Cells[$row, $cTotal].Formula = "SUMIFS($script:sDurationRef,$Range,$key)"
        $Worksheet.Cells[$row, $cErrors].Formula = "SUMIFS($script:sErrorsRef,$Range,$key)"
        $Worksheet.Cells[$row, $cWarnings].Formula = "SUMIFS($script:sWarningsRef,$Range,$key)"
        foreach ($c in $cAvg, $cMin, $cMax, $cTotal) {
            $Worksheet.Cells[$row, $c].Style.Numberformat.Format = $script:durationFormat
        }
        $row++
    }
    return $row
}

$script:sDurationRef = $sDuration
$script:sErrorsRef = $sErrors
$script:sWarningsRef = $sWarnings

$byComputer = @(
    $settingRows | Group-Object ComputerName | ForEach-Object {
        [pscustomobject] @{ Key = $_.Name; Average = (Get-AverageSeconds $_.Group) }
    } | Sort-Object Average -Descending
)
$row = Add-TargetBlock -Worksheet $target -StartRow 4 `
    -Title "Per computer ($($byComputer.Count) computers, slowest first)" `
    -KeyHeader 'ComputerName' -Range $sComputer -Groups $byComputer
$row++

$byPath = @(
    $settingRows | Group-Object Path | ForEach-Object {
        [pscustomobject] @{
            Key      = $_.Name
            Computer = (@($_.Group | Group-Object ComputerName | Sort-Object Count -Descending)[0]).Name
            Average  = (Get-AverageSeconds $_.Group)
        }
    } | Sort-Object Average -Descending
)
$null = Add-TargetBlock -Worksheet $target -StartRow $row `
    -Title "Per path ($($byPath.Count) paths, slowest first)" `
    -KeyHeader 'Path' -Range $sPath -Groups $byPath -WithComputer

Set-ColumnWidths -Worksheet $target -Widths @(52, 18, 13, 13, 13, 13, 13, 10, 10)

# ---- sheet 4: notes ------------------------------------------------------
$wsNotes = $pkg.Workbook.Worksheets.Add('Notes & method')
$wsNotes.Cells.Style.Font.Name = 'Arial'
$wsNotes.Cells.Style.Font.Size = 10
$wsNotes.Cells['A1'].Value = 'How this workbook was built'
$wsNotes.Cells['A1'].Style.Font.Bold = $true
$wsNotes.Cells['A1'].Style.Font.Size = 13
$wsNotes.Cells['A1'].Style.Font.Color.SetColor($navy)

$linkBase = if ($LinkRoot) { $LinkRoot } else { $root }
$notes = @(
    @('Generated', "$(Get-Date -Format 'dd/MM/yyyy HH:mm:ss') by $($MyInvocation.MyCommand.Name) on $env:COMPUTERNAME"),
    @('Source', "$root - $($runs.Count) matrix file runs in $($runFolders.Count) run folders"),
    @('Sheet 1', 'One row per logged error or warning, newest first: Type says which of the two it is, Name and Description say what happened. Information level entries are not on this sheet.'),
    @('Sheet 3', 'Same columns plus Information level entries and the full context (path, site, run folder, setting id, detail file).'),
    @('Links', "Blue underlined cells open the log itself: the matrix file name on sheet 1 and on the summary, and the DetailFile and LogFile columns on sheet 3. They point at $linkBase, so they keep working for anyone who can reach that path. Old format rows link to the 'ID <n> - Settings.html' of the computer concerned; new format rows link to the execution report."),
    @('Per computer and path', 'A setting is one computer plus one path, and it is the setting that carries the duration, so that sheet is the finest grain available. A matrix file can hold many settings, which is why the per computer totals do not add up to the per matrix file totals.'),
    @('Duration', 'The duration of a matrix file is the sum of the durations of its settings, so it is processing time rather than elapsed time. It is deliberately not End time minus Start time: in the newer reports that footer holds the window of the whole run and is identical in every report of that run. On the Performance sheet, Elapsed is the run window and Processing time is the sum of the matrix file durations, so Processing time above Elapsed means settings ran in parallel.'),
    @('Two log formats', 'Old runs use "ID <n> - Settings.html", newer runs use "00 - Execution Report.html". The LogFormat column says which one a row came from.'),
    @('DateTime', 'New format rows use the exact timestamp from the problem detail JSON when the report links to one; otherwise the start time of the setting (old format) or of the run is used. See the TimestampSource column.'),
    @('ComputerName', 'Taken from the setting a problem belongs to. File level issues (for example "Excel File: Runspace processing failed") are not tied to a computer, so the cell is empty.'),
    @('Run level issues', 'SystemErrors.json / "System errors log.json" belong to the run as a whole, not to one matrix file, so they show as MatrixFileName "(run level)" with an ordinary Type. They are not all errors despite the file name: the Type field in the JSON decides, and it also records Warning and Information. Filter MatrixFileName on "(run level)" to see only these, or filter them out to count what the matrix files themselves caused.'),
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

# ---- vertical centering and sheet order ----------------------------------
# Excel puts text on the bottom of the row by default, which looks off as soon as
# one column wraps onto several lines, so centre everything that has content.
foreach ($sheet in $pkg.Workbook.Worksheets) {
    if ($sheet.Dimension) {
        $sheet.Cells[$sheet.Dimension.Address].Style.VerticalAlignment =
        [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    }
}

$pkg.Workbook.Worksheets.MoveAfter('Summary', 'Errors & Warnings')
$pkg.Workbook.Worksheets.MoveAfter('Performance', 'Summary')
$pkg.Workbook.Worksheets.MoveAfter('Performance by target', 'Performance')
$pkg.Workbook.Worksheets.MoveAfter($runsSheet, 'All issues (detail)')
$pkg.Workbook.Worksheets.MoveAfter($settingsSheet, $runsSheet)
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