#Requires -RunAsAdministrator
#Requires -Version 7

<#
    .SYNOPSIS
        Validates system requirements and enforces standardized SMB share 
        settings on a target computer.

    .DESCRIPTION
        This script acts as the prerequisite validation and baseline 
        configuration step for the Permission Matrix pipeline. 
        It evaluates the local computer to ensure it meets the minimum 
        execution requirements (Administrator privileges, PowerShell version, 
        and .NET Framework 4.6.2+ for long-path support).

        Additionally, it audits any SMB shares matching the provided paths. If 
        discrepancies are found, it automatically corrects them by:
        1. Toggling Access-Based Enumeration (ABE) to the desired state.
        2. Resetting the SMB Share Permissions to a standardized baseline 
        (defaulting to Administrators:Full, Authenticated Users:Change). 
        
        By standardizing the SMB layer, the script ensures that all effective 
        access control is securely and exclusively managed at the NTFS file 
        system layer.

    .PARAMETER Path
        An array of local directory paths to audit. If a path is actively 
        shared via SMB, its share properties will be validated and corrected.

    .PARAMETER Flag
        Determines the Access-Based Enumeration (ABE) state for the matching 
        SMB shares.
        - $true  : ABE is enabled (Users only see files/folders they have 
        permission to access).
        - $false : ABE is disabled (Unrestricted enumeration).

    .PARAMETER RequiredSharePermissions
        An array of hashtables defining the exact baseline permissions required 
        on the SMB share. If the current share permissions deviate from this 
        baseline, the script will revoke all existing share access and 
        forcefully apply these exact rules.

    .PARAMETER MinimumPowerShellVersion
        A hashtable defining the absolute minimum required version of 
        PowerShell (Default: Major 7, Minor 1).

    .EXAMPLE
        $paths = @('E:\Data\HR', 'E:\Data\Finance')
        .\Test-Requirements.ps1 -Path $paths -Flag $true
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

$abeCorrected = [ordered]@{}
$permissionsCorrected = [ordered]@{}

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
                
                $incorrectPermissions = [ordered]@{}

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