#Requires -RunAsAdministrator
#Requires -Version 5.1

<#
    .SYNOPSIS
        Test if a computer is capable of running the Permission Matrix script.

    .DESCRIPTION
        Test the current computer for administrator permissions, .NET version,
        PowerShell version, ...

        Also correct the smb share permissions when they are incorrect.
        Required smb share permissions:
        - Enabled Access Based Enumeration (ABE)
        - Administrator: FullControl
        - Everyone: Modify

    .PARAMETER Path
        Shared folder paths.

    .PARAMETER Flag
        Valid values:
        - True  : ABE will be enabled
        - False : ABE will be disabled

    .NOTES
        Computers with an OS older than Windows Server 2012 are not supported
        as the module 'SmbShare' would not be available.
#>

[OutputType([PSCustomObject])]
[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String[]]$Path,
    [Parameter(Mandatory)]
    [Boolean]$Flag,
    [hashtable[]]$requiredSharePermissions = @(
        @{
            AccountName = 'BUILTIN\Administrators'
            AccessRight = 'Full'
        }
        @{
            AccountName = 'NT AUTHORITY\Authenticated Users'
            AccessRight = 'Change'
        }
    )
)

Begin {
    Function Test-IsRequiredDotNetVersionHC {
        $dotNet = Get-ChildItem 'HKLM:SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\' -ErrorAction 'Ignore' |
        Get-ItemPropertyValue -Name Release | ForEach-Object { $_ -ge 394802 }

        if ($dotNet) {
            $true
        }
        else {
            $false
        }
    }

    Function Test-IsAdminHC {
        <#
            .SYNOPSIS
                Check if a user is local administrator.

            .DESCRIPTION
                Check if a user is member of the local group 'Administrators' and returns
                TRUE if he is, FALSE if not.

            .EXAMPLE
                Test-IsAdminHC -SamAccountName SrvBatch
                Returns TRUE in case SrvBatch is admin on this machine

            .EXAMPLE
                Test-IsAdminHC
                Returns TRUE if the current user is admin on this machine

            .NOTES
                CHANGELOG
                2017/05/29 Added parameter to check for a specific user
        #>

        [CmdLetBinding()]
        [OutputType([Boolean])]
        Param (
            $SamAccountName = [Security.Principal.WindowsIdentity]::GetCurrent()
        )

        Try {
            $Identity = [Security.Principal.WindowsIdentity]$SamAccountName
            $Principal = New-Object Security.Principal.WindowsPrincipal -ArgumentList $Identity
            $Result = $Principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
            Write-Verbose "Administrator permissions: $Result"
            $Result
        }
        Catch {
            throw "Failed to determine if the user '$SamAccountName' is local admin: $_"
        }
    }

    Function Test-IsRequiredPowerShellVersionHC {
        [CmdLetBinding()]
        [OutputType([Boolean])]
        Param (
            [Int]$Major = 5,
            [Int]$Minor = 1
        )

        (
            ($PSVersionTable.PSVersion.Major -ge $Major) -and
            ($PSVersionTable.PSVersion.Minor -ge $Minor)
        )
    }
}

Process {
    #region Require administrator privileges
    if (-not (Test-IsAdminHC)) {
        Return [PSCustomObject]@{
            Type        = 'FatalError'
            Name        = 'Administrator privileges'
            Description = "Administrator privileges are required to be able to apply permissions."
            Value       = "SamAccountName '$env:USERNAME'"
        }
    }
    #endregion

    #region Require at least PowerShell 5.1
    if (-not (Test-IsRequiredPowerShellVersionHC)) {
        Return [PSCustomObject]@{
            Type        = 'FatalError'
            Name        = 'PowerShell version'
            Description = "PowerShell version 5.1 or higher is required to be able to use advanced methods."
            Value       = "PowerShell $($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor)"
        }
    }
    #endregion

    #region Require at least .NET 4.6.2
    if (-not (Test-IsRequiredDotNetVersionHC)) {
        Return [PSCustomObject]@{
            Type        = 'FatalError'
            Name        = '.NET Framework version'
            Description = "Microsoft .NET Framework version 4.6.2 or higher is required to be able to traverse long path names and use advanced PowerShell methods."
            Value       = $null
        }
    }
    #endregion

    $smbShares = Get-SmbShare

    $Path = $Path | Sort-Object -Unique

    $abeCorrected = @{}
    $permissionsCorrected = @{}

    foreach ($p in $Path) {
        foreach (
            $share in
            $smbShares | Where-Object {
                ($_.Path -like "$p\*") -or ($_.Path -eq $p)
            }
        ) {
            Write-Verbose "Smb share '$($share.Name)' path '$($share.Path)'"

            #region Set Access based enumeration
            if (($share.FolderEnumerationMode -eq 'AccessBased') -ne $Flag) {
                Try {
                    $params = @{
                        Name                  = $share.Name
                        FolderEnumerationMode = if ($Flag) {
                            'AccessBased'
                        }
                        else {
                            'Unrestricted'
                        }
                        ErrorAction           = 'Stop'
                        Force                 = $true
                    }
                    Write-Verbose "Set FolderEnumerationMode to '$($params.FolderEnumerationMode)'"

                    Set-SmbShare @params

                    $abeCorrected.Add($share.Name, $share.Path)
                }
                Catch {
                    throw "Failed setting FolderEnumerationMode to '$($params.FolderEnumerationMode)' for path '$p' on '$env:COMPUTERNAME': $_"
                }
            }
            #endregion

            #region Set smb share permissions
            $smbShareAccess = Get-SmbShareAccess -InputObject $share

            $correctPermissions = 0

            foreach ($permission in $requiredSharePermissions) {
                $smbShareAccess | Where-Object {
                    ($_.AccountName -eq $permission.AccountName) -and
                    ($_.AccessRight -eq $permission.AccessRight)
                } | ForEach-Object {
                    $correctPermissions++
                }
            }

            if (
                ($requiredSharePermissions.Count -ne $smbShareAccess.Count) -or
                ($requiredSharePermissions.Count -ne $correctPermissions)
            ) {
                try {
                    #region Remove incorrect smb share permissions
                    $incorrectPermissions = @{}

                    $smbShareAccess.ForEach(
                        {
                            Write-Verbose "Remove incorrect smb share permission '$($_.AccountName):$($_.AccessRight)'"

                            $incorrectPermissions[$_.AccountName] = $_.AccessRight.ToString()

                            Revoke-SmbShareAccess -Name $share.Name -AccountName $_.AccountName -Force
                        }
                    )

                    $permissionsCorrected.Add(
                        $share.Name, $incorrectPermissions
                    )
                    #endregion

                    #region Add correct smb share permissions
                    $requiredSharePermissions.ForEach(
                        {
                            Write-Verbose "Add correct smb share permission '$($_.AccountName):$($_.AccessRight)'"

                            $params = $_
                            Grant-SmbShareAccess -Name $share.Name @params -Force
                        }
                    )
                    #endregion
                }
                Catch {
                    throw "Failed setting share permissions on path '$Path' on '$env:COMPUTERNAME': $_"
                }
            }
            #endregion
        }
    }

    #region Return result objects
    if ($abeCorrected.Count -ne 0) {
        [PSCustomObject]@{
            Type        = 'Warning'
            Name        = 'Access Based Enumeration'
            Description = "Access Based Enumeration should be set to '$flag'. This will hide files and folders where the users don't have access to. We fixed this now."
            Value       = $abeCorrected
        }
    }

    if ($permissionsCorrected.Count -ne 0) {
        [PSCustomObject]@{
            Type        = 'Warning'
            Name        = 'Share permissions'
            Description = "The share permissions are now set to 'Administrators: FullControl' and 'Authenticated users: Change'. The effective permissions are managed on NTFS level."
            Value       = $permissionsCorrected
        }
    }
    #endregion
}