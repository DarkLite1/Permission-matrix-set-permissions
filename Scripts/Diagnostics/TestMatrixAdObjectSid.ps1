#Requires -Version 7
#Requires -Modules ImportExcel

<#
    .SYNOPSIS
        Reports matrix AD objects that resolve in Active Directory but carry no
        readable ObjectSid.

    .DESCRIPTION
        READ ONLY. This script queries Active Directory and reads matrix Excel
        files. It does not read, compare or write a single ACL, it does not
        write to Active Directory, and it does not touch the permission path.
        Safe to run against production data at any time.

        Why it exists
        -------------
        Invoke-PermissionMatrixBeginHC rewrites every matrix ACL from AD object
        names to SIDs. Two conditions are checked in two different places:

            Test-AdObjectInMatrixHC   rejects an object when 'adObject' is null
            the SID rewrite           needs 'adObject.ObjectSid' as well

        An object that resolves but whose ObjectSid is null therefore passes
        validation and is then dropped from the ACL by the rewrite, with no
        error raised. On a protected ACL under Action 'Fix' that silently
        removes the group's access.

        Everything with a samAccountName is normally a security principal and
        has an objectSid, so this is expected to report nothing. The realistic
        triggers are an account that cannot READ objectSid on the target object
        (a restrictive OU ACL) or a Global Catalog partial-attribute-set edge
        case.

        Run this before deciding whether that gap is worth closing, and how.

        What it does
        ------------
        1. Reads the 'Settings' and 'Permissions' sheets of every matrix file.
        2. Rebuilds the AD object names exactly as the real run does, by
           calling the module's own Get-MatrixADObjectsMapHC for each enabled
           Settings row, so GroupName/SiteCode placeholders resolve identically.
        3. Adds the ADObjectName values from the defaults file, when given.
        4. Resolves every unique name once with the module's own
           Get-ADObjectDetailHC.
        5. Groups the results:
               OK          resolved with an ObjectSid
               NoSid       resolved WITHOUT an ObjectSid  <-- the silent drop
               NotFound    did not resolve at all (already a FatalError today)

    .PARAMETER MatrixFolder
        Folder holding the matrix .xlsx files. Searched non-recursively, to
        match how the real run picks up files.

    .PARAMETER DefaultsFile
        Optional path to the defaults .xlsx file. Its Settings sheet supplies
        additional AD object names that the real run also resolves.

    .PARAMETER ExportPath
        Optional .csv path. When given, the full per-object result is written
        there as well as summarised on screen.

    .PARAMETER MaxThreads
        Concurrent AD queries. Matches the module default. (Default: 7)

    .EXAMPLE
        .\TestMatrixAdObjectSid.ps1 -MatrixFolder '\\server\share\Matrix'

    .EXAMPLE
        .\TestMatrixAdObjectSid.ps1 `
            -MatrixFolder '\\server\share\Matrix' `
            -DefaultsFile '\\server\share\Matrix\Defaults.xlsx' `
            -ExportPath 'C:\temp\MatrixSidCheck.csv'
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [String]$MatrixFolder,
    [String]$DefaultsFile,
    [String]$ExportPath,
    [ValidateRange(1, 64)]
    [Int]$MaxThreads = 7
)

$ErrorActionPreference = 'Stop'

#region Load the module's own helpers, so the names are built the same way
$repoRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$privateFolder = Join-Path $repoRoot 'Modules\PermissionMatrix\Private'

if (-not (Test-Path -LiteralPath $privateFolder -PathType Container)) {
    throw "Could not find '$privateFolder'. Run this script from its place in the repository."
}

<#
    Matrix.ps1 references Add-ErrorHC and the Test-Matrix* validators from
    functions this script never calls. ErrorHandling.ps1 is loaded anyway so a
    stray reference can never fail at runtime.
#>
foreach ($helper in @('Utils.ps1', 'ErrorHandling.ps1', 'Matrix.ps1', 'ActiveDirectory.ps1')) {
    $helperPath = Join-Path $privateFolder $helper

    if (-not (Test-Path -LiteralPath $helperPath -PathType Leaf)) {
        throw "Could not find '$helperPath'."
    }

    . $helperPath
}
#endregion

if (-not (Test-Path -LiteralPath $MatrixFolder -PathType Container)) {
    throw "Matrix folder '$MatrixFolder' not found."
}

#region Collect the AD object names used by the matrix files
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

Write-Host "Reading $($matrixFile.Count) matrix file(s) from '$MatrixFolder'" -ForegroundColor Cyan

# AD object name -> the files and settings rows that use it
$usage = @{}

function Add-Usage {
    <#
        AllowEmptyString because a blank cell is normal: the defaults sheet has
        rows with no ADObjectName, and Mandatory alone would reject those at
        binding time, before the guard below could skip them.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyString()][AllowNull()][string]$Name,
        [Parameter(Mandatory)][string]$Source
    )

    if ([string]::IsNullOrWhiteSpace($Name)) { return }

    $key = $Name.Trim()

    if (-not $usage.ContainsKey($key)) {
        $usage[$key] = [System.Collections.Generic.List[string]]::new()
    }

    if ($usage[$key] -notcontains $Source) { $usage[$key].Add($Source) }
}

$fileProcessed = 0
$fileSkipped = [System.Collections.Generic.List[string]]::new()

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
        $fileSkipped.Add("$($file.Name) (read failed)")
        continue
    }

    # Same 'Enabled' test as Import-MatrixFileHC
    $enabledSetting = @(
        $settingsSheet | Where-Object { "$($_.Status)".Trim() -eq 'Enabled' }
    )

    if ($enabledSetting.Count -eq 0) {
        Write-Verbose "No enabled Settings row in '$($file.Name)'"
        $fileSkipped.Add("$($file.Name) (no enabled Settings row)")
        continue
    }

    if ($permissionsSheet.Count -lt 4) {
        Write-Warning "Skipped '$($file.Name)': the Permissions sheet has fewer than 4 rows."
        $fileSkipped.Add("$($file.Name) (Permissions sheet too small)")
        continue
    }

    $formattedPermissions = @($permissionsSheet | Format-PermissionsStringsHC)

    foreach ($setting in ($enabledSetting | Format-SettingStringsHC)) {
        $adObjectsMap = Get-MatrixADObjectsMapHC `
            -PermissionsSheet $formattedPermissions `
            -SettingRow $setting

        foreach ($adName in $adObjectsMap.Values) {
            Add-Usage -Name $adName -Source $file.Name
        }
    }

    $fileProcessed++
}

if ($DefaultsFile) {
    Write-Host "Reading defaults from '$DefaultsFile'" -ForegroundColor Cyan

    <#
        The try covers only the file read. Keeping the row loop outside it means
        one unexpected row can no longer abandon the rest of the sheet, which
        would silently under-report the AD objects in use.
    #>
    $defaultsSheet = $null

    try {
        $defaultsSheet = @(
            Import-Excel -Path $DefaultsFile -Sheet 'Settings' -DataOnly
        )
    }
    catch {
        Write-Warning "Could not read the defaults file: $_"
    }

    $defaultsNameCount = 0

    foreach ($row in $defaultsSheet) {
        $defaultsName = [string]$row.ADObjectName

        if ([string]::IsNullOrWhiteSpace($defaultsName)) { continue }

        Add-Usage -Name $defaultsName -Source '<Defaults>'
        $defaultsNameCount++
    }

    Write-Host "  $defaultsNameCount AD object name(s) from the defaults file"
}
#endregion

$uniqueName = @($usage.Keys | Sort-Object)

if ($uniqueName.Count -eq 0) {
    throw 'No AD object names found in the matrix files.'
}

Write-Host "Resolving $($uniqueName.Count) unique AD object name(s)" -ForegroundColor Cyan

#region Resolve every name once, with the module's own lookup
$adDetail = @(
    Get-ADObjectDetailHC `
        -ADObjectName $uniqueName `
        -Type 'SamAccountName' `
        -MaxThreads $MaxThreads
)

$detailLookup = @{}

foreach ($detail in $adDetail) {
    if ($detail.SamAccountName) {
        $detailLookup[[string]$detail.SamAccountName] = $detail
    }
}
#endregion

#region Classify
$result = foreach ($name in $uniqueName) {
    $detail = $detailLookup[$name]

    $status = if (-not $detail -or -not $detail.adObject) {
        'NotFound'
    }
    elseif (-not $detail.adObject.ObjectSid) {
        'NoSid'
    }
    else {
        'OK'
    }

    [PSCustomObject]@{
        Status            = $status
        ADObjectName      = $name
        ObjectClass       = if ($detail.adObject) { $detail.adObject.ObjectClass } else { $null }
        ObjectSid         = if ($detail.adObject) { $detail.adObject.ObjectSid } else { $null }
        DistinguishedName = if ($detail.adObject) { $detail.adObject.DistinguishedName } else { $null }
        UsedIn            = ($usage[$name] -join '; ')
    }
}

$result = @($result)
#endregion

#region Report
$ok = @($result | Where-Object Status -EQ 'OK')
$noSid = @($result | Where-Object Status -EQ 'NoSid')
$notFound = @($result | Where-Object Status -EQ 'NotFound')

Write-Host ''
Write-Host '--- Matrix AD object SID check ---' -ForegroundColor Cyan
Write-Host "  Matrix files found     : $($matrixFile.Count)"
Write-Host "  Matrix files used      : $fileProcessed"
Write-Host "  Matrix files skipped   : $($fileSkipped.Count)" -ForegroundColor $(
    if ($fileSkipped.Count) { 'Yellow' } else { 'Green' }
)
Write-Host "  Unique AD objects      : $($result.Count)"
Write-Host "  Resolved with a SID    : $($ok.Count)" -ForegroundColor Green
Write-Host "  Resolved WITHOUT a SID : $($noSid.Count)" -ForegroundColor $(
    if ($noSid.Count) { 'Red' } else { 'Green' }
)
Write-Host "  Not found in AD        : $($notFound.Count)" -ForegroundColor $(
    if ($notFound.Count) { 'Yellow' } else { 'Green' }
)
Write-Host ''

if ($noSid.Count -gt 0) {
    Write-Host 'These objects resolve in AD but have no readable ObjectSid.' -ForegroundColor Red
    Write-Host 'The real run drops them from the ACL without reporting anything:' -ForegroundColor Red
    $noSid | Format-Table ADObjectName, ObjectClass, DistinguishedName, UsedIn -AutoSize -Wrap
}
else {
    Write-Host 'No object resolves without a SID: the silent-drop path is not' -ForegroundColor Green
    Write-Host 'reachable with the current matrix files and the account running' -ForegroundColor Green
    Write-Host 'this check.' -ForegroundColor Green
    Write-Host ''
}

if ($notFound.Count -gt 0) {
    Write-Host 'These objects do not resolve at all. The real run already reports' -ForegroundColor Yellow
    Write-Host "them as 'Unknown AD Objects in Matrix' (FatalError):" -ForegroundColor Yellow
    $notFound | Format-Table ADObjectName, UsedIn -AutoSize -Wrap
}

if ($fileSkipped.Count -gt 0) {
    Write-Host 'These matrix files contributed no AD object names:' -ForegroundColor Yellow
    $fileSkipped | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
    Write-Host ''
}

if ($ExportPath) {
    $result | Sort-Object Status, ADObjectName |
    Export-Csv -LiteralPath $ExportPath -NoTypeInformation -Delimiter ';'

    Write-Host "Full result written to '$ExportPath'" -ForegroundColor Cyan
}
#endregion

return $result