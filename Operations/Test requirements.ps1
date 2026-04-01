#Requires -RunAsAdministrator
#Requires -Version 7

<#
    .SYNOPSIS
        Test if a computer is capable of running the Permission Matrix script.

    .DESCRIPTION
        Test the current computer for administrator permissions, .NET version,
        PowerShell version, ...

        Also correct the smb share permissions when they are incorrect.

    .PARAMETER Path
        Shared folder paths.

    .PARAMETER Flag
        Valid values:
        - True  : ABE will be enabled
        - False : ABE will be disabled

    .PARAMETER RequiredSharePermissions
        The smb share permissions that are required on the share. If the
        current smb share permissions are not matching, they will be replaced
        with the correct ones.

        If a folder in Path is not configured as an smb share, it will be
        ignored.

    .PARAMETER MinimumPowerShellVersion
        The minimal required PowerShell version to run the Permission Matrix
        script.
#>

[OutputType([PSCustomObject])]
[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [String[]]$Path,
    
    [Parameter(Mandatory)]
    [Boolean]$Flag,
    
    [HashTable[]]$RequiredSharePermissions = @(
        @{
            AccountName = 'BUILTIN\Administrators'
            AccessRight = 'Full'
        }
        @{
            AccountName = 'NT AUTHORITY\Authenticated Users'
            AccessRight = 'Change'
        }
    ),
    
    [HashTable]$MinimumPowerShellVersion = @{
        Major = 7
        Minor = 1
    }
)

#region Helper Functions
function Test-IsAdminHC {
    try {
        $Identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $Principal = New-Object Security.Principal.WindowsPrincipal($Identity)
        $Result = $Principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        Write-Verbose "Administrator permissions: $Result"
        return $Result
    }
    catch {
        throw "Failed to determine if the current user is local admin: $_"
    }
}

function Test-IsRequiredDotNetVersionHC {
    # High-speed direct read without the pipeline
    $dotNetRelease = Get-ItemPropertyValue -LiteralPath 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\' -Name 'Release' -ErrorAction Ignore
    return ($null -ne $dotNetRelease -and $dotNetRelease -ge 394802)
}

function Test-IsRequiredPowerShellVersionHC {
    # Fixed mathematical logic: (Major > MinMajor) OR (Major == MinMajor AND Minor >= MinMinor)
    return (
        ($PSVersionTable.PSVersion.Major -gt $MinimumPowerShellVersion.Major) -or 
        (($PSVersionTable.PSVersion.Major -eq $MinimumPowerShellVersion.Major) -and ($PSVersionTable.PSVersion.Minor -ge $MinimumPowerShellVersion.Minor))
    )
}
#endregion

#region System Requirement Checks (Early Exits)
if (-not (Test-IsAdminHC)) {
    return [PSCustomObject]@{
        Type        = 'FatalError'
        Name        = 'Administrator privileges'
        Description = 'Administrator privileges are required to be able to apply permissions.'
        Value       = "Account '$([Security.Principal.WindowsIdentity]::GetCurrent().Name)'"
    }
}

if (-not (Test-IsRequiredPowerShellVersionHC)) {
    return [PSCustomObject]@{
        Type        = 'FatalError'
        Name        = 'PowerShell version'
        Description = "PowerShell version $($MinimumPowerShellVersion.Major).$($MinimumPowerShellVersion.Minor) or higher is required."
        Value       = "PowerShell $($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor)"
    }
}

if (-not (Test-IsRequiredDotNetVersionHC)) {
    return [PSCustomObject]@{
        Type        = 'FatalError'
        Name        = '.NET Framework version'
        Description = 'Microsoft .NET Framework version 4.6.2 or higher is required to be able to traverse long path names and use advanced PowerShell methods.'
        Value       = $null
    }
}
#endregion

#region Core Processing
$smbShares = Get-SmbShare
$uniquePaths = $Path | Sort-Object -Unique

$abeCorrected = @{}
$permissionsCorrected = @{}

foreach ($p in $uniquePaths) {
    # Fast intrinsic filtering based on exact match or subfolder match
    $matchingShares = $smbShares.Where({ $_.Path -eq $p -or $_.Path.StartsWith("$p\") })

    foreach ($share in $matchingShares) {
        Write-Verbose "Smb share '$($share.Name)' path '$($share.Path)'"

        #region Set Access Based Enumeration (ABE)
        try {
            # FolderEnumerationMode: 0 = AccessBased, 1 = Unrestricted
            $isAbeEnabled = ($share.FolderEnumerationMode -eq 'AccessBased') -or ($share.FolderEnumerationMode -eq 0)
            
            if ($isAbeEnabled -ne $Flag) {
                $abeMode = if ($Flag) { 'AccessBased' } else { 'Unrestricted' }
                
                Write-Verbose "Set FolderEnumerationMode to '$abeMode'"

                Set-SmbShare -Name $share.Name -FolderEnumerationMode $abeMode -ErrorAction Stop -Force

                # Use index assignment to prevent duplicate key errors
                $abeCorrected[$share.Name] = $share.Path
            }
        }
        catch {
            throw "Failed setting FolderEnumerationMode to '$abeMode' for path '$p' on '$env:COMPUTERNAME': $_"
        }
        #endregion

        #region Set SMB Share Permissions
        try {
            $smbShareAccess = Get-SmbShareAccess -InputObject $share -ErrorAction Stop

            $smbSharePermissions = $smbShareAccess.ForEach({
                    @{
                        AccountName       = $_.AccountName
                        AccessControlType = switch ($_.AccessControlType) {
                            0 { 'Allow' }
                            1 { 'Deny' }
                            default { [String]$_ }
                        }
                        AccessRight       = switch ($_.AccessRight) {
                            0 { 'Full' }
                            1 { 'Change' }
                            2 { 'Read' }
                            default { [String]$_ }
                        }
                    }
                })

            $correctPermissionsCount = 0

            foreach ($permission in $RequiredSharePermissions) {
                $isCorrect = $smbSharePermissions.Where({
                        $_.AccountName -eq $permission.AccountName -and
                        $_.AccessControlType -eq 'Allow' -and
                        $_.AccessRight -eq $permission.AccessRight
                    }, 'First')

                if ($isCorrect) {
                    $correctPermissionsCount++
                }
            }

            # If the share doesn't have the EXACT match of required permissions, rebuild it
            if (($RequiredSharePermissions.Count -ne $smbShareAccess.Count) -or ($RequiredSharePermissions.Count -ne $correctPermissionsCount)) {
                
                $incorrectPermissions = @{}

                # Revoke all existing permissions
                $smbSharePermissions.ForEach({
                        Write-Verbose "Remove incorrect smb share permission '$($_.AccountName):$($_.AccessRight)'"
                        $incorrectPermissions[$_.AccountName] = [String]$_.AccessRight
                        $null = Revoke-SmbShareAccess -Name $share.Name -AccountName $_.AccountName -ErrorAction Stop -Force
                    })

                $permissionsCorrected[$share.Name] = $incorrectPermissions

                # Grant the exact required baseline permissions
                $RequiredSharePermissions.ForEach({
                        Write-Verbose "Add correct smb share permission '$($_.AccountName): $($_.AccessRight)'"
                    
                        $grantParams = $_
                        $null = Grant-SmbShareAccess -Name $share.Name @grantParams -ErrorAction Stop -Force
                    })
            }
        }
        catch {
            throw "Failed setting share permissions on path '$p' on '$env:COMPUTERNAME': $_"
        }
        #endregion
    }
}
#endregion

#region Return Result Objects
if ($abeCorrected.Count -gt 0) {
    [PSCustomObject]@{
        Type        = 'Warning'
        Name        = 'Access Based Enumeration'
        Description = "Access Based Enumeration should be set to '$Flag'. This will hide files and folders where the users don't have access to. We fixed this now."
        Value       = $abeCorrected
    }
}

if ($permissionsCorrected.Count -gt 0) {
    $requiredString = ($RequiredSharePermissions.ForEach({ "'$($_.AccountName): $($_.AccessRight)'" })) -join ', '
    
    [PSCustomObject]@{
        Type        = 'Warning'
        Name        = 'Share permissions'
        Description = "The share permissions are now set to $requiredString. The effective permissions are managed on NTFS level."
        Value       = $permissionsCorrected
    }
}
#endregion