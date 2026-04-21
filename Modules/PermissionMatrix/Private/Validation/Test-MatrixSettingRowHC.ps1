function Test-MatrixSettingRowHC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$SettingRow
    )

    $checks = [System.Collections.Generic.List[pscustomobject]]::new()

    $validActions = @('Fix', 'New', 'Check')
    if ($SettingRow.Action -notin $validActions) {
        $checks.Add([pscustomobject]@{
                Type        = 'FatalError'
                Name        = 'Invalid Action'
                Description = "The Action must be one of: $($validActions -join ', ')"
                Value       = "Found: '$($SettingRow.Action)'"
            })
    }

    if ([string]::IsNullOrWhiteSpace($SettingRow.Path)) {
        $checks.Add([pscustomobject]@{
                Type        = 'FatalError'
                Name        = 'Missing Path'
                Description = 'The Path column cannot be empty.'
                Value       = "Found: '$($SettingRow.Path)'"
            })
    }

    if ([string]::IsNullOrWhiteSpace($SettingRow.ComputerName)) {
        $checks.Add([pscustomobject]@{
                Type        = 'FatalError'
                Name        = 'Missing ComputerName'
                Description = 'The ComputerName column cannot be empty.'
                Value       = "Found: '$($SettingRow.ComputerName)'"
            })
    }

    <# 
        # Only required when permissions header hold this

        if ([string]::IsNullOrWhiteSpace($SettingRow.SiteName)) {
            $checks.Add([pscustomobject]@{
                    Type        = 'FatalError'
                    Name        = 'Missing SiteName'
                    Description = 'The SiteName column cannot be empty.'
                    Value       = "Found: '$($SettingRow.SiteName)'"
                })
        }

        if ([string]::IsNullOrWhiteSpace($SettingRow.SiteCode)) {
            $checks.Add([pscustomobject]@{
                    Type        = 'FatalError'
                    Name        = 'Missing SiteCode'
                    Description = 'The SiteCode column cannot be empty.'
                    Value       = "Found: '$($SettingRow.SiteCode)'"
                })
        }

        if ([string]::IsNullOrWhiteSpace($SettingRow.GroupName)) {
            $checks.Add([pscustomobject]@{
                    Type        = 'FatalError'
                    Name        = 'Missing GroupName'
                    Description = 'The GroupName column cannot be empty.'
                    Value       = "Found: '$($SettingRow.GroupName)'"
                })
        } 
    #>

    return $checks
}