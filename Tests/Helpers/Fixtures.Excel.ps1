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

        'MissingAction' {
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

        'MissingAction' {
            return @(
                [pscustomobject]@{
                    Status                  = 'Enabled'
                    SiteName                = 'HQ South'
                    SiteCode                = 'CS&L'
                    ComputerName            = 'BEL$FFRAN0001'
                    Path                    = 'E:\DEPARTMENTS\Sagrev\GROUPS\C&S&L'
                    GroupName               = 'BEL ROL-AGS-SAGREV'
                    # Action    = 'Check'
                    ApplyDefaultPermissions = $false
                }
            )
        }

        'InvalidAction' {
            return @(
                [pscustomobject]@{
                    Status                  = 'Enabled'
                    SiteName                = 'HQ South'
                    SiteCode                = 'CS&L'
                    ComputerName            = 'BEL$FFRAN0001'
                    Path                    = 'E:\DEPARTMENTS\Sagrev\GROUPS\C&S&L'
                    GroupName               = 'BEL ROL-AGS-SAGREV'
                    Action                  = 'INVALID'
                    ApplyDefaultPermissions = $false
                }
            )
        }

        'MissingComputerName' {
            return @(
                [pscustomobject]@{
                    Status                  = 'Enabled'
                    SiteName                = 'HQ South'
                    SiteCode                = 'CS&L'
                    ComputerName            = $null
                    Path                    = 'E:\DEPARTMENTS\Sagrev\GROUPS\C&S&L'
                    GroupName               = 'BEL ROL-AGS-SAGREV'
                    Action                  = 'Fix'
                    ApplyDefaultPermissions = $true
                }
            )
        }

        'MissingGroupName' {
            return @(
                [pscustomobject]@{
                    Status                  = 'Enabled'
                    SiteName                = 'HQ South'
                    SiteCode                = 'CS&L'
                    ComputerName            = 'BEL$FFRAN0001'
                    Path                    = 'E:\DEPARTMENTS\Sagrev\GROUPS\C&S&L'
                    GroupName               = $null
                    Action                  = 'Fix'
                    ApplyDefaultPermissions = $true

                }
            )
        }

        'MissingSiteCode' {
            return @(
                [pscustomobject]@{
                    Status                  = 'Enabled'
                    SiteName                = 'HQ South'
                    SiteCode                = $null
                    ComputerName            = 'BEL$FFRAN0001'
                    Path                    = 'E:\DEPARTMENTS\Sagrev\GROUPS\C&S&L'
                    GroupName               = $null
                    Action                  = 'Fix'
                    ApplyDefaultPermissions = $true
                }
            )
        }

        'MissingPath' {
            return @(
                [pscustomobject]@{
                    Status                  = 'Enabled'
                    SiteName                = 'HQ South'
                    SiteCode                = 'CS&L'
                    ComputerName            = 'BEL$FFRAN0001'
                    Path                    = $null
                    GroupName               = 'BEL ROL-AGS-SAGREV'
                    Action                  = 'Fix'
                    ApplyDefaultPermissions = $true
                }
            )
        }

        'MissingApplyDefaultPermissions' {
            return @(
                [pscustomobject]@{
                    Status       = 'Enabled'
                    SiteName     = 'HQ South'
                    SiteCode     = 'CS&L'
                    ComputerName = 'BEL$FFRAN0001'
                    Path         = 'E:\DEPARTMENTS\Sagrev\GROUPS\C&S&L'
                    GroupName    = 'BEL ROL-AGS-SAGREV'
                    Action       = 'Fix'
                    # ApplyDefaultPermissions = $true
                }
            )
        }

        'Valid' {
            return @(
                [pscustomobject]@{
                    Status                  = 'Enabled'
                    SiteName                = 'HQ South'
                    SiteCode                = 'CS&L'
                    ComputerName            = 'BEL$FFRAN0001'
                    Path                    = 'E:\DEPARTMENTS\Sagrev\GROUPS\C&S&L'
                    GroupName               = 'BEL ROL-AGS-SAGREV'
                    Action                  = 'Fix'
                    ApplyDefaultPermissions = $false
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
                    @{ Path = 'Finance'      ; Col2 = 'R' ; Col3 = 'R' }
                    @{ Path = 'Finance\Docs' ; Col2 = 'W' ; Col3 = 'W' }
                )
            }
        }

        'WithGroupNamePlaceholder' {
            return @{
                Row1 = @('', '')
                Row2 = @('', 'Director')
                Row3 = @('', 'GroupName') 
                Row4 = @('Path', '')
                Data = @(
                    @{ Path = 'Finance'      ; Col2 = 'L' }
                    @{ Path = 'Finance\Docs' ; Col2 = 'W' }
                )
            }
        }

        'WithSiteCodePlaceholder' {
            return @{
                Row1 = @('', 'Director')
                Row2 = @('', 'SiteCode')
                Row3 = @('', 'BEL') 
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
                Row3 = @('', '', 'Mike')               
                Row4 = @('Path', '', 'L')              
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
                    @{ Path = 'Finance'      ; Col2 = 'X'   ; Col3 = 'R' }
                    @{ Path = 'Finance\Docs' ; Col2 = 'R'   ; Col3 = 'YYY' } 
                )
            }
        }

        # NEW SCENARIOS ADDED BELOW:

        'MissingRows' {
            # Only sending 2 rows total to trigger the '< 4' check
            return @{
                Row1 = @('', 'Bob', 'Mike')
                Row2 = @('Path', 'L', 'L')
            }
        }

        'MissingColumns' {
            # Only sending the Path column to trigger the '< 2' check
            return @{
                Row1 = @('')
                Row2 = @('')
                Row3 = @('')
                Row4 = @('Path')
                Data = @(
                    @{ Path = 'Finance' }
                )
            }
        }

        'MissingFolderName' {
            return @{
                Row1 = @('', '', '')
                Row2 = @('', '', '')
                Row3 = @('', 'Bob', 'Mike')
                Row4 = @('Path', 'L', 'L')
                Data = @(
                    @{ Path = 'Finance' ; Col2 = 'R' ; Col3 = 'R' }
                    @{ Path = ''        ; Col2 = 'W' ; Col3 = 'W' } # Blank path
                )
            }
        }

        'DuplicateFolderName' {
            return @{
                Row1 = @('', '', '')
                Row2 = @('', '', '')
                Row3 = @('', 'Bob', 'Mike')
                Row4 = @('Path', 'L', 'L')
                Data = @(
                    @{ Path = 'Finance' ; Col2 = 'R' ; Col3 = 'R' }
                    @{ Path = 'Finance' ; Col2 = 'W' ; Col3 = 'W' } # Exact duplicate
                )
            }
        }

        'InaccessibleFolders' {
            # Parent folder has no valid permission ('L'), and deep folder also only has 'L'
            return @{
                Row1 = @('', '', '')
                Row2 = @('', '', '')
                Row3 = @('', 'Bob', 'Mike')
                Row4 = @('Path', 'L', 'L')
                Data = @(
                    @{ Path = 'Finance'      ; Col2 = 'L' ; Col3 = 'L' }
                    @{ Path = 'Finance\Docs' ; Col2 = 'L' ; Col3 = '' } 
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
        # $fileName = Split-Path $Path -Leaf
        $FormDataRows = @(
            [pscustomobject]@{
                # MatrixFileName        = $fileName
                MatrixFolderDisplayName = 'Default Folder'
                MatrixFolderPath        = 'E:\Folder'
                MatrixCategoryName      = 'Default'
                MatrixSubCategoryName   = 'General'
                MatrixResponsible       = 'owner@example.com'
                MatrixFormStatus        = $Disabled ? 'Disabled' : 'Enabled'
            }
        )
    }

    $FormDataRows |
    Export-Excel -Path $Path -WorksheetName 'FormData' -TableName 'FormData' `
        -ClearSheet -AutoSize -FreezeTopRow
}
