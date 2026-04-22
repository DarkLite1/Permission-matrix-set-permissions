function New-ValidDefaultsExcelFixture {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    # Required columns for the script to function
    $rows = @(
        [pscustomobject]@{
            MailTo       = 'owner@example.com'
            ADObjectName = 'DefaultGroup'
            Permission   = 'R'
        }
    )

    $rows | Export-Excel `
        -Path $Path `
        -WorksheetName 'Settings' `
        -TableName 'Settings' `
        -AutoSize `
        -FreezeTopRow `
        -ErrorAction Stop

    return $Path
}
function New-MatrixSettingsExcelFixture {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Scenario
    )

    switch ($Scenario) {

        'MissingColumn' {
            $rows = @(
                [pscustomobject]@{
                    Status       = 'Enabled'
                    SiteName     = 'HQ South'
                    SiteCode     = 'CS&L'
                    ComputerName = 'BEL$FFRAN0001'
                    Path         = 'E:\DEPARTMENTS\Sagrev\GROUPS\C&S&L'
                    GroupName    = 'BEL ROL-AGS-SAGREV'
                    # Action missing
                }
            )
        }

        'InvalidAction' {
            $rows = @(
                [pscustomobject]@{
                    Status       = 'Enabled'
                    SiteName     = 'HQ South'
                    SiteCode     = 'CS&L'
                    ComputerName = 'BEL$FFRAN0001'
                    Path         = 'E:\DEPARTMENTS\Sagrev\GROUPS\C&S&L'
                    GroupName    = 'BEL ROL-AGS-SAGREV'
                    Action       = 'INVALID'
                }
            )
        }

        'MissingComputerName' {
            $rows = @(
                [pscustomobject]@{
                    Status       = 'Enabled'
                    SiteName     = 'HQ South'
                    SiteCode     = 'CS&L'
                    ComputerName = $null
                    Path         = 'E:\DEPARTMENTS\Sagrev\GROUPS\C&S&L'
                    GroupName    = 'BEL ROL-AGS-SAGREV'
                    Action       = 'Fix'
                }
            )
        }

        'MissingGroupName' {
            $rows = @(
                [pscustomobject]@{
                    Status       = 'Enabled'
                    SiteName     = 'HQ South'
                    SiteCode     = 'CS&L'
                    ComputerName = 'BEL$FFRAN0001'
                    Path         = 'E:\DEPARTMENTS\Sagrev\GROUPS\C&S&L'
                    GroupName    = $null
                    Action       = 'Fix'
                }
            )
        }

        'MissingPath' {
            $rows = @(
                [pscustomobject]@{
                    Status       = 'Enabled'
                    SiteName     = 'HQ South'
                    SiteCode     = 'CS&L'
                    ComputerName = 'BEL$FFRAN0001'
                    Path         = $null
                    GroupName    = 'BEL ROL-AGS-SAGREV'
                    Action       = 'Fix'
                }
            )
        }

        default { throw "Unknown Settings scenario '$Scenario'" }
    }

    $rows | Export-Excel `
        -Path $Path `
        -WorksheetName 'Settings' `
        -TableName 'Settings' `
        -AutoSize `
        -FreezeTopRow `
        -ClearSheet `
        -ErrorAction Stop
}
function New-MatrixSettingsFixtureRows {
    param([Parameter(Mandatory)][string]$Scenario)

    switch ($Scenario) {

        'MissingColumn' {
            return @(
                [pscustomobject]@{
                    Status    = 'Enabled'
                    SiteName  = 'HQ South'
                    SiteCode  = 'CS&L'
                    # missing ComputerName
                    # ComputerName = 'BEL$FFRAN0001' 
                    Path      = 'E:\DEPARTMENTS\Sagrev\GROUPS\C&S&L'
                    GroupName = 'BEL ROL-AGS-SAGREV'
                    Action    = 'Check'
                }
            )
        }

        'InvalidAction' {
            return @(
                [pscustomobject]@{
                    Status       = 'Enabled'
                    SiteName     = 'HQ South'
                    SiteCode     = 'CS&L'
                    ComputerName = 'BEL$FFRAN0001'
                    Path         = 'E:\DEPARTMENTS\Sagrev\GROUPS\C&S&L'
                    GroupName    = 'BEL ROL-AGS-SAGREV'
                    Action       = 'INVALID'
                }
            )
        }

        'MissingComputerName' {
            return @(
                [pscustomobject]@{
                    Status       = 'Enabled'
                    SiteName     = 'HQ South'
                    SiteCode     = 'CS&L'
                    ComputerName = $null
                    Path         = 'E:\DEPARTMENTS\Sagrev\GROUPS\C&S&L'
                    GroupName    = 'BEL ROL-AGS-SAGREV'
                    Action       = 'Fix'
                }
            )
        }

        'MissingGroupName' {
            return @(
                [pscustomobject]@{
                    Status       = 'Enabled'
                    SiteName     = 'HQ South'
                    SiteCode     = 'CS&L'
                    ComputerName = 'BEL$FFRAN0001'
                    Path         = 'E:\DEPARTMENTS\Sagrev\GROUPS\C&S&L'
                    GroupName    = $null
                    Action       = 'Fix'
                }
            )
        }

        'MissingPath' {
            return @(
                [pscustomobject]@{
                    Status       = 'Enabled'
                    SiteName     = 'HQ South'
                    SiteCode     = 'CS&L'
                    ComputerName = 'BEL$FFRAN0001'
                    Path         = $null
                    GroupName    = 'BEL ROL-AGS-SAGREV'
                    Action       = 'Fix'
                }
            )
        }

        'Valid' {
            return @(
                [pscustomobject]@{
                    Status       = 'Enabled'
                    SiteName     = 'HQ South'
                    SiteCode     = 'CS&L'
                    ComputerName = 'BEL$FFRAN0001'
                    Path         = 'E:\DEPARTMENTS\Sagrev\GROUPS\C&S&L'
                    GroupName    = 'BEL ROL-AGS-SAGREV'
                    Action       = 'Fix'
                }
            )
        }

        default { throw "Unknown Settings scenario: $Scenario" }
    }
}
function New-MatrixPermissionsExcelFixture {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][hashtable]$Spec
    )

    # Build rows 1–4 exactly
    $rows = @()

    foreach ($i in 1..4) {
        $rowData = $Spec["Row$i"]
        $obj = [ordered]@{}
        for ($c = 0; $c -lt $rowData.Count; $c++) {
            $obj["Column$($c+1)"] = $rowData[$c]
        }
        $rows += [pscustomobject]$obj
    }

    # Add data rows (row5+)
    foreach ($d in $Spec.Data) {
        $obj = [ordered]@{}
        $obj['Column1'] = $d.Path
        $obj['Column2'] = $d.Col2
        $obj['Column3'] = $d.Col3
        $rows += [pscustomobject]$obj
    }

    # Output to Excel
    $rows |
    Export-Excel `
        -Path $Path `
        -WorksheetName 'Permissions' `
        -ClearSheet `
        -FreezeTopRow:$false `
        -AutoSize:$false `
        -NoHeader `
        -ErrorAction Stop

    return $Path
}
function New-MatrixPermissionsFixtureRows {
    param([Parameter()][string]$Scenario = 'Valid')

    switch ($Scenario) {

        'Valid' {
            return @{
                Row1 = @('', '', '')                 
                Row2 = @('', '', '')                 
                Row3 = @('', 'Bob', 'Mike')          
                Row4 = @('Path', 'L', 'L')           
                Data = @(
                    @{ Path = 'Finance'       ; Col2 = 'R' ; Col3 = 'R' }
                    @{ Path = 'Finance\Docs'  ; Col2 = 'W' ; Col3 = 'W' }
                )
            }
        }

        'WithGroupNamePlaceholder' {
            # No GroupName or SiteCode
            return @{
                Row1 = @('', '')
                Row2 = @('', 'Director')
                Row3 = @('', 'GroupName')  # placeholder that should be replaced by actual GroupName from Settings         
                Row4 = @('Path', '')
                Data = @(
                    @{ Path = 'Finance'      ; Col2 = 'L' }
                    @{ Path = 'Finance\Docs' ; Col2 = 'W' }
                )
            }
        }

        'WithSiteCodePlaceholder' {
            # No GroupName or SiteCode
            return @{
                Row1 = @('', 'Director')
                Row2 = @('', 'SiteCode')
                Row3 = @('', 'BEL')  # placeholder that should be replaced by actual SiteCode from Settings         
                Row4 = @('Path', '')
                Data = @(
                    @{ Path = 'Finance'      ; Col2 = 'L' }
                    @{ Path = 'Finance\Docs' ; Col2 = 'W' }
                )
            }
        }

        'MissingADObjectName' {
            return @{
                Row1 = @('', '', '')
                Row2 = @('', '', '')
                Row3 = @('', '', 'Mike')               # Column 2 = missing
                Row4 = @('Path', '', 'L')              # Column 2 header = missing
                Data = @(
                    @{ Path = 'Finance'      ; Col2 = 'R' ; Col3 = 'R' }
                    @{ Path = 'Finance\Docs' ; Col2 = 'W' ; Col3 = 'W' }
                )
            }
        }

        'InvalidPermissionChar' {
            return @{
                Row1 = @('', '', '')
                Row2 = @('', '', '')
                Row3 = @('', 'Bob', 'Mike')
                Row4 = @('Path', 'L', 'L')
                Data = @(
                    @{ Path = 'Finance'      ; Col2 = 'X'   ; Col3 = 'R' }  # invalid
                    @{ Path = 'Finance\Docs' ; Col2 = 'R'   ; Col3 = 'YYY' }  # invalid
                )
            }
        }

        default { throw "Unknown Permissions scenario '$Scenario'" }
    }
}
function New-MatrixExcelFixture {
    param(
        [Parameter(Mandatory)][string]$Path,
        [array]$SettingsRows,
        [hashtable]$PermissionsRows = (
            New-MatrixPermissionsFixtureRows -Scenario 'Valid'
        ),
        [array]$FormDataRows,
        [switch]$Disabled
    )

    if (-not $SettingsRows) {
        $SettingsRows = New-MatrixSettingsFixtureRows -Scenario 'Valid'
    }

    if (-not $PermissionsRows) {
        $PermissionsRows = New-MatrixPermissionsFixtureRows -Scenario 'Valid'
    }

    if ($Disabled) {
        foreach ($row in $SettingsRows) {
            $row.Status = 'Disabled'
        }
    }

    # SETTINGS
    $SettingsRows |
    Export-Excel -Path $Path -WorksheetName 'Settings' -TableName 'Settings' `
        -ClearSheet -AutoSize -FreezeTopRow

    # PERMISSIONS
    New-MatrixPermissionsExcelFixture -Path $Path -Spec $PermissionsRows | Out-Null

    # FORMDATA
    if (-not $FormDataRows) {
        $fileName = Split-Path $Path -Leaf
        $FormDataRows = @(
            [pscustomobject]@{
                MatrixFileName        = $fileName
                MatrixFolderPath      = 'E:\Folder'
                MatrixCategoryName    = 'Default'
                MatrixSubCategoryName = 'General'
                MatrixResponsible     = 'owner@example.com'
                MatrixFormStatus      = $Disabled ? 'Disabled' : 'Enabled'
            }
        )
    }

    $FormDataRows |
    Export-Excel -Path $Path -WorksheetName 'FormData' -TableName 'FormData' `
        -ClearSheet -AutoSize -FreezeTopRow
}
