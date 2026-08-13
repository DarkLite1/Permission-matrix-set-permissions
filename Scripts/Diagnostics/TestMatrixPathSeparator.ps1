#Requires -Version 7
#Requires -Modules ImportExcel

<#
    .SYNOPSIS
        Reports matrix folder paths that contain a forward slash, and proves
        whether that breaks the ignored-folder lookup on this host.

    .DESCRIPTION
        READ ONLY. Reads matrix Excel files and performs pure string operations.
        It does not read, compare or write a single ACL, it does not touch
        Active Directory, and it does not enter the permission path. Safe to run
        against production data at any time.

        Why it exists
        -------------
        SetPermissions.ps1 builds a lookup of folders it must not walk into:

            key     Join-Path -Path $Path -ChildPath $M.Path   (from the matrix)
            lookup  $IgnoredFolderPaths.ContainsKey($child.FullName)

        $child.FullName always comes back from .NET with backslashes.
        Format-PermissionsStringsHC only trims TRAILING separators, so a
        forward slash written inside a path in the Permissions sheet survives
        into the key. If Join-Path does not normalise it, the key and the lookup
        differ as strings and the hashtable misses.

        A miss means the folder is treated as an unknown child, so a row marked
        'I' (ignore) would be walked and its permissions reset to inherited
        only, which is exactly what the 'I' was there to prevent.

        Every file API involved accepts both separators, so the folders are
        found and created correctly. Only the string comparison fails, which is
        why this would never show up as an error.

        What it does
        ------------
        1. Proves empirically, on this host and this PowerShell version, whether
           Join-Path normalises an interior forward slash. This settles the
           mechanism rather than assuming it.
        2. Scans the 'Path' column of the Settings sheet and the folder column
           of the Permissions sheet of every matrix file for forward slashes.
        3. For every hit, shows the key that would be built and the FullName
           that would be looked up, and whether they match.

    .PARAMETER MatrixFolder
        Folder holding the matrix .xlsx files. Searched non-recursively.

    .PARAMETER DefaultsFile
        Optional path to the defaults .xlsx file, excluded from the scan.

    .PARAMETER ExportPath
        Optional .csv path for the full per-row result.

    .EXAMPLE
        .\TestMatrixPathSeparator.ps1 -MatrixFolder '\\server\share\Matrix'
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [String]$MatrixFolder,
    [String]$DefaultsFile,
    [String]$ExportPath
)

$ErrorActionPreference = 'Stop'

#region Prove the mechanism on this host, before looking at any data
Write-Host '--- Join-Path behaviour on this host ---' -ForegroundColor Cyan

<#
 Two shapes are probed. The UNC one matters most: matrix parent paths are
 usually UNC, and provider behaviour is not guaranteed to be identical for
 both.

 Join-Path resolves the drive through the PowerShell provider, so it throws
 for a drive that does not exist on THIS machine even though the path is
 valid on the target server. That is caught here; the real run calls
 Join-Path on the remote machine, where the drive does exist.
#>
$probe = @(
    @{ Parent = 'C:\Data'; Child = 'sub/deep' }
    @{ Parent = '\\server\share'; Child = 'sub/deep' }
)

$separatorIsNormalised = $true

foreach ($p in $probe) {
    $probeKey = try {
        Join-Path -Path $p.Parent -ChildPath $p.Child
    }
    catch {
        Write-Host "  Join-Path '$($p.Parent)' '$($p.Child)' -> could not resolve: $_" -ForegroundColor Yellow
        continue
    }

    $probeNormalised = [System.IO.Path]::GetFullPath($probeKey)
    $isSame = ($probeKey -eq $probeNormalised)

    Write-Host "  Join-Path '$($p.Parent)' '$($p.Child)'"
    Write-Host "    key built   : '$probeKey'"
    Write-Host "    normalised  : '$probeNormalised'"
    Write-Host "    match       : $isSame" -ForegroundColor $(
        if ($isSame) { 'Green' } else { 'Red' }
    )

    if (-not $isSame) { $separatorIsNormalised = $false }
}

Write-Host ''

if ($separatorIsNormalised) {
    Write-Host '  -> Join-Path normalises the separator. The ignored-folder' -ForegroundColor Green
    Write-Host '     lookup cannot miss on separator grounds. The scan below is' -ForegroundColor Green
    Write-Host '     a style check only.' -ForegroundColor Green
}
else {
    Write-Host '  -> Join-Path does NOT normalise. A forward slash in a matrix' -ForegroundColor Red
    Write-Host '     path would make the ignored-folder lookup miss.' -ForegroundColor Red
}

Write-Host ''
#endregion

if (-not (Test-Path -LiteralPath $MatrixFolder -PathType Container)) {
    throw "Matrix folder '$MatrixFolder' not found."
}

$matrixFile = @(
    Get-ChildItem -LiteralPath $MatrixFolder -Filter '*.xlsx' -File |
    Where-Object { $_.Name -notlike '~$*' }
)

if ($DefaultsFile) {
    $resolvedDefaults = (Resolve-Path -LiteralPath $DefaultsFile).ProviderPath
    $matrixFile = @(
        $matrixFile | Where-Object { $_.FullName -ne $resolvedDefaults }
    )
}

if ($matrixFile.Count -eq 0) {
    throw "No matrix .xlsx files found in '$MatrixFolder'."
}

Write-Host "Scanning $($matrixFile.Count) matrix file(s)" -ForegroundColor Cyan

$finding = [System.Collections.Generic.List[psobject]]::new()
$rowsScanned = 0
$fileScanned = 0

foreach ($file in $matrixFile) {
    try {
        $settingsSheet = @(
            Import-Excel -Path $file.FullName -Sheet 'Settings' -DataOnly
        )

        $permissionsSheet = @(
            Import-Excel -Path $file.FullName -Sheet 'Permissions' -NoHeader -DataOnly
        )
    }
    catch {
        Write-Warning "Skipped '$($file.Name)': $_"
        continue
    }

    $fileScanned++

    $enabledSetting = @(
        $settingsSheet | Where-Object { "$($_.Status)".Trim() -eq 'Enabled' }
    )

    <#
     The parent path from the Settings sheet. A forward slash here affects
     every child of that matrix, not just one row, so it is reported separately.
    #>
    foreach ($setting in $enabledSetting) {
        $settingPath = "$($setting.Path)".Trim()

        if ($settingPath -notmatch '/') { continue }

        $finding.Add(
            [PSCustomObject]@{
                File             = $file.Name
                Sheet            = 'Settings'
                Row              = ''
                Value            = $settingPath
                IsIgnored        = $false
                KeyBuilt         = $settingPath
                KeyMethod        = 'n/a (parent path, not combined)'
                Normalised       = [System.IO.Path]::GetFullPath($settingPath)
                WouldMissLookup  = (-not $separatorIsNormalised)
                HidesParent      = ''
                EscapesIgnore    = ''
            }
        )
    }

    <#
     Collect every folder row first. The separator questions below are about a
     path's relationship to the OTHER rows in the same sheet, so a single row in
     isolation cannot answer them.

     Row 1-3 are headers and row 4 is the parent folder, so sub-folder rows
     start at index 4. The parent row is included: it is written to the same
     column and a slash there is just as wrong.
    #>
    $parentPath = if ($enabledSetting.Count -gt 0) {
        "$($enabledSetting[0].Path)".Trim()
    }
    else { '' }

    $folderRow = [System.Collections.Generic.List[psobject]]::new()

    for ($i = 3; $i -lt $permissionsSheet.Count; $i++) {
        $rowsScanned++

        $cell = "$($permissionsSheet[$i].P1)".Trim()

        if ([string]::IsNullOrWhiteSpace($cell)) { continue }

        <#
         Mirrors Test-MatrixPermissionsHC: a row is ignored when ANY permission
         column holds 'I'. The sheet is already formatted, so the value is
         trimmed and upper-cased.
        #>
        $isIgnored = [bool](
            $permissionsSheet[$i].PSObject.Properties.Where({
                    $_.Name -ne 'P1' -and "$($_.Value)".Trim() -eq 'I'
                }, 'First').Count
        )

        $folderRow.Add(
            [PSCustomObject]@{
                Row       = $i + 1
                Path      = $cell
                IsIgnored = $isIgnored
            }
        )
    }

    <#
     Test-MatrixPermissionsHC trims a trailing backslash before comparing, so
     mirror that here to classify against the same strings it sees.
    #>
    $allPath = @($folderRow.Path | ForEach-Object { $_.TrimEnd('\') })

    foreach ($row in $folderRow) {
        if ($row.Path -notmatch '/') { continue }

        $normalisedRow = $row.Path.TrimEnd('\')

        <#
         Effect 1: deepest-folder detection.

         Test-MatrixPermissionsHC finds children with StartsWith("$P\"), so a
         parent row cannot see a child written with a forward slash. Report the
         parent that would be misclassified as a leaf.
        #>
        $hiddenParent = @(
            $allPath | Where-Object {
                $candidate = $_
                ($candidate -ne $normalisedRow) -and
                $normalisedRow.StartsWith(
                    ('{0}/' -f $candidate),
                    [System.StringComparison]::OrdinalIgnoreCase
                )
            }
        )

        <#
         Effect 2: ignored-subtree exclusion.

         $isInIgnoredSubtree uses the same backslash prefix, so a path below an
         'I' root written with a forward slash is not recognised as excluded.
        #>
        $missedIgnoreRoot = @(
            $folderRow | Where-Object {
                $_.IsIgnored -and
                ($_.Path.TrimEnd('\') -ne $normalisedRow) -and
                $normalisedRow.StartsWith(
                    ('{0}/' -f $_.Path.TrimEnd('\')),
                    [System.StringComparison]::OrdinalIgnoreCase
                )
            } | ForEach-Object { $_.Path }
        )

        $keyMethod = 'Join-Path (same as the real run)'

        $keyBuilt = if ($parentPath) {
            try {
                Join-Path -Path $parentPath -ChildPath $row.Path
            }
            catch {
                $keyMethod = 'Path::Combine (drive not on this host; NOT what the real run builds)'
                [System.IO.Path]::Combine($parentPath, $row.Path)
            }
        }
        else { $row.Path }

        $normalised = try { [System.IO.Path]::GetFullPath($keyBuilt) }
        catch { '<could not resolve>' }

        $finding.Add(
            [PSCustomObject]@{
                File             = $file.Name
                Sheet            = 'Permissions'
                Row              = $row.Row
                Value            = $row.Path
                IsIgnored        = $row.IsIgnored
                KeyBuilt         = $keyBuilt
                KeyMethod        = $keyMethod
                Normalised       = $normalised
                # Decided by the probe, not by this row's key
                WouldMissLookup  = (-not $separatorIsNormalised)
                # Parent row that will be wrongly treated as a leaf
                HidesParent      = ($hiddenParent -join '; ')
                # 'I' root this row should have been excluded by
                EscapesIgnore    = ($missedIgnoreRoot -join '; ')
            }
        )
    }
}

#region Report
Write-Host ''
Write-Host '--- Matrix path separator check ---' -ForegroundColor Cyan
Write-Host "  Matrix files scanned    : $fileScanned"
Write-Host "  Permissions rows scanned: $rowsScanned"
Write-Host "  Paths with a '/'        : $($finding.Count)" -ForegroundColor $(
    if ($finding.Count -and (-not $separatorIsNormalised)) { 'Red' }
    elseif ($finding.Count) { 'Yellow' }
    else { 'Green' }
)
Write-Host ''

if ($finding.Count -eq 0) {
    Write-Host 'No matrix path contains a forward slash.' -ForegroundColor Green
    Write-Host ''
}
else {
    if ($separatorIsNormalised) {
        Write-Host 'Ignored-folder lookup: NOT affected. Join-Path normalises the' -ForegroundColor Green
        Write-Host 'separator, so the key matches the enumerated FullName and' -ForegroundColor Green
        Write-Host 'permissions are applied correctly.' -ForegroundColor Green
    }
    else {
        Write-Host 'Ignored-folder lookup: AFFECTED. Join-Path does not normalise' -ForegroundColor Red
        Write-Host 'on this host, so these rows would be walked and reset.' -ForegroundColor Red
    }

    Write-Host ''

    <#
     Separate from the lookup: Test-MatrixPermissionsHC compares paths with a
     backslash prefix, so a forward slash hides a parent/child relationship
     from it. Both effects raise or suppress a 'Warning' only, so neither can
     skip a matrix or change an ACL.
    #>
    $hidesParent = @($finding | Where-Object { $_.HidesParent })
    $escapesIgnore = @($finding | Where-Object { $_.EscapesIgnore })

    Write-Host "  Rows hiding a parent from the leaf check : $($hidesParent.Count)" -ForegroundColor $(
        if ($hidesParent.Count) { 'Yellow' } else { 'Green' }
    )
    Write-Host "  Rows escaping an 'I' ignore root         : $($escapesIgnore.Count)" -ForegroundColor $(
        if ($escapesIgnore.Count) { 'Yellow' } else { 'Green' }
    )
    Write-Host ''

    if ($hidesParent.Count -eq 0 -and $escapesIgnore.Count -eq 0) {
        Write-Host 'No slashed path sits below another row in its own sheet, so the' -ForegroundColor Green
        Write-Host 'validation checks are not affected either. These paths are a' -ForegroundColor Green
        Write-Host 'consistency observation only:' -ForegroundColor Green
        Write-Host ''
        $finding | Format-Table File, Sheet, Row, Value, Normalised -AutoSize -Wrap
    }
    else {
        if ($hidesParent.Count) {
            Write-Host 'These rows hide a parent, which is then wrongly treated as a' -ForegroundColor Yellow
            Write-Host "leaf folder and may raise a false 'Inaccessible folders' warning:" -ForegroundColor Yellow
            $hidesParent | Format-Table File, Row, Value, HidesParent -AutoSize -Wrap
        }

        if ($escapesIgnore.Count) {
            Write-Host "These rows sit below an 'I' root but are not recognised as" -ForegroundColor Yellow
            Write-Host 'excluded, so they may raise a warning that should be suppressed:' -ForegroundColor Yellow
            $escapesIgnore | Format-Table File, Row, Value, EscapesIgnore -AutoSize -Wrap
        }
    }
}

if ($ExportPath) {
    $finding | Export-Csv -LiteralPath $ExportPath -NoTypeInformation -Delimiter ';'
    Write-Host "Full result written to '$ExportPath'" -ForegroundColor Cyan
}
#endregion

return $finding