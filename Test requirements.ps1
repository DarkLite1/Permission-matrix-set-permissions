#Requires -RunAsAdministrator
#Requires -Version 5.1

<# 
.SYNOPSIS   
    Set Access Based Enumeration and permissions on a shared folder.

.DESCRIPTION
    Correct the following settings on shared folder:
    - Access Based Enumeration: Enabled/Disabled based on the Flag parameter
    - Share permission 'Everyone' set to 'FullControl' and all others are removed

    Most of these functions are available in the module 'SmbShare'. However, this module
    is only available from Windows Server 2012+. The current environment still has older
    servers that need to be supported too. That's why this script is built.

.PARAMETER Path
    Shared folder paths.

.PARAMETER Flag
    If set to TRUE, ABE will be enabled. Otherwise it will be disabled.

.NOTE
    CHANGELOG
        2018/06/29 Script born
        2018/07/09 Added test for local admin
        2018/10/02 Added check for .NET version 4.6.2 (needed for long path name support)

    AUTHOR Brecht.Gijbels@heidelbergcement.com #>

[OutputType([PSCustomObject])]
[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String[]]$Path,
    [Parameter(Mandatory)]
    [Boolean]$Flag
)

Begin {
    Function Get-SharePermissionHC {
        <#
        .SYNOPSIS   
            Retrieve all shares on the local machine with their permissions.

        .DESCRIPTION
            Retrieve all shares and their share permissions (not NTFS permissions) of folders
            on the local machine.

        .EXAMPLE
            Retrieve all shares and their permissions

            $Shares = Get-SharePermissionHC

            $Shares
            Name    Path             Acl                                                     
            ----    ----             ---                                                     
            Log     L:\              {Everyone}                                              
            Scripts T:\Input\Scripts {GROUPHC\Domain Users, Everyone, BUILTIN\Administrators}

            $Shares[0].Acl
            Name  : Everyone
            Value : FullControl

            $Shares[0].Acl
            Name  : GROUPHC\Domain Users
            Value : FullControl

            Name  : Everyone
            Value : FullControl

            Name  : BUILTIN\Administrators
            Value : FullControl

        .NOTE
            CHANGELOG
                2016/02/26 Function born
                2018/06/29 Optimized code with 'CIM' CmdLets instead of 'WMI'
                           Speed up with '.Foreach' method
                           Added error handling
                2019/08/22 Fixed a bug with weird characters in the share name

            AUTHOR Brecht.Gijbels@heidelbergcement.com #>
        
        [CmdletBinding()]
        Param ()

        Try {
            $lss = Get-CimInstance -ClassName Win32_LogicalShareSecuritySetting -Verbose:$false
            $Shares = Get-CimInstance -ClassName Win32_Share -Verbose:$false

            foreach ($s in $lss) {
                Write-Verbose "Share name '$($S.Name)'"

                $sh = $Shares | Where-Object { $_.Name -eq $S.Name }
                $sd = (Invoke-CimMethod -InputObject $s -MethodName GetSecurityDescriptor -Verbose:$false).Descriptor
        
                $Obj = [PSCustomObject]@{
                    Name = $S.Name
                    Path = $sh.Path
                    Acl  = @{}
                }

                ($sd.DACL).Foreach( {
                        $permission = Switch ($_.AccessMask) {
                            1179817 { ‘Read’; Break }
                            1245631 { ‘Change’; Break }
                            2032127 { ‘FullControl’; Break }
                            default { ‘Special’ }
                        }

                        $domain = $_.Trustee.Domain
                        $userName = if ($tn = $_.Trustee.Name) { $tn } else { $_.Trustee.SID }

                        $User = if ($domain) { "$domain\$userName" } else { $userName }
                        $Permission = $permission

                        $Obj.Acl.Add($user, $permission)
                    
                    })

                $Obj
            }
        }
        Catch {
            throw "Failed retrieving the share permissions on '$ENV:ComputerName': $_"
        }
    }

    Function Set-AccessBasedEnumerationHC {
        <# 
        .SYNOPSIS   
            Set Access Based Enumeration on a shared folder.

        .DESCRIPTION
            Set Access Based Enumeration enabled status to 'true' or 'false' on a 
            shared folder.

        .PARAMETER ComputerName
            Specifies the target computer.

        .PARAMETER Name
            Name of the share.

        .PARAMETER Flag
            If 'Enabled' we set the ABE to TRUE in case of 'Disabled' we set it to FALSE.

        .EXAMPLE
            Enabled ABE on share 'Log' of the localhost
            
            Set-AccessBasedEnumerationHC -Name log -Flag $true -Verbose

            VERBOSE: Access Based Enumeration enabled on share 'log' for 'ServerName'

        .NOTE
            CHANGELOG
                2018/06/29 Rewrote the function to just return a boolean
                Optimized for speed by using the method '.Foreach'
                Removed object creation as it's overkill
                Improved help and added OutputType

            AUTHOR Brecht.Gijbels@heidelbergcement.com #>

        [CmdLetbinding()]
        [OutputType()]
        Param (
            [Parameter(Mandatory, ValueFromPipeline)]
            [String[]]$Name,
            [Parameter(Mandatory)]
            [Boolean]$Flag,
            [ValidateNotNullOrEmpty()]
            [String]$ComputerName = $env:COMPUTERNAME
        )

        Process {
            $Name.Foreach( {
                    Try {
                        $ShareInfo = [NetApi32]::NetShareGetInfo_1005($ComputerName, $_)

                        if ($Flag -eq $true) {
                            $ShareInfo.Shi1005_flags = ($ShareInfo.Shi1005_flags -bor [Shi1005_flags]::SHI1005_FLAGS_ACCESS_BASED_DIRECTORY_ENUM)

                            if (([NetApi32]::NetShareSetInfo_1005($ComputerName, $_, $ShareInfo)) -eq 0) {
                                Write-Verbose "ABE enabled on share '$_' "
                            }
                            else {
                                throw "Couldn't verify the Access Based Enumeration permissions"
                            }
                        }
                        else {    
                            $ShareInfo.Shi1005_flags = 0

                            if (([NetApi32]::NetShareSetInfo_1005($ComputerName, $_, $ShareInfo)) -eq 0) {
                                Write-Verbose "ABE disabled on share '$_'"
                            }
                            else {
                                throw "Couldn't verify the Access Based Enumeration permissions"
                            }
                        }
                    }
                    Catch {
                        throw "Failed setting Access Based Enumeration to '$Flag' for share '$Name' on '$ComputerName': $_"
                    }
                })
        }
    }

    Function Set-SharePermissionsHC {
        <# 
    .SYNOPSIS
        Correct the permissions on a share.
    
    .DESCRIPTION
        To make sure that permissions are only managed on NTFS level, the share permissions
        on a shared folder need to be set as following:
        - Authenticated users: Change & Read
        - Administrators: FullControl

        This function is created because the module 'SmbShare', which already contains these
        features, is not available on Windows Server 2008.

        https://helgeklein.com/blog/2009/01/how-to-restrict-users-from-changing-permissions-on-file-servers/

        How to Prevent Users from Changing Permissions on File Servers

        On file servers in corporate environments one typically does not want users to change permissions, even on their own files. It might seem that it would be sufficient to simply grant change permissions instead of full control, but unfortunately that is not the case. The problem is that whenever a new file gets created, the user creating the file will be its owner. And owners can always change permissions, regardless of the contents of the DACL.

        The Solution

        In order to prevent “orderly” users from “tidying” the permissions on their files and directories and thus messing things up, often removing administrators from the DACL, too, the following needs to be done:

        Only grant change (aka modify) permissions in the NTFS file system. “Change” does not include the specific right “change permissions”.
        Do not grant full share permissions. Use change + read instead. This masks out the right “change permissions” which owners are implicitly granted. This obviously applies to network access only.
        The clever part is not granting “full control” in the share permissions to users. Since administrators still want to be able to modify permissions, I suggest adding a second ACE to each share’s DACL. The resulting DACL now contains the following two entries:

        - Authenticated users: change + read
        - Administrators: full control
    
    .PARAMETER Name
        Name of the shared folder.
   
    .NOTES
    	CHANGELOG
    	2018/07/03 Function born
    
    	AUTHOR Brecht.Gijbels@heidelbergcement.com #>

        [CmdLetbinding()]
        [OutputType()]
        Param (
            [Parameter(Mandatory)]
            [String]$Name
        )

        Try {
            #region Authenticated users: Change & Read
            $trustee = ([WMIClass]'Win32_trustee').PSBase.CreateInstance()
            $trustee.Domain = $null
            $trustee.Name = 'Authenticated users'

            $ace = ([WMIClass]'Win32_ACE').PSBase.CreateInstance()
            $ace.AccessMask = [Uint32][System.Security.AccessControl.FileSystemRights]'1245631'
            $ace.AceFlags = [Uint32][System.Security.AccessControl.AceFlags]::None
            $ace.AceType = [Uint32][System.Security.AccessControl.AceType]::AccessAllowed
            $ace.Trustee = $trustee
            #endregion

            #region Adminstrators: FullControl
            $trustee = ([WMIClass]'Win32_trustee').PSBase.CreateInstance()
            $trustee.Domain = $null
            $trustee.Name = 'Administrators'

            $ace2 = ([WMIClass]'Win32_ACE').PSBase.CreateInstance()
            $ace2.AccessMask = [Uint32][System.Security.AccessControl.FileSystemRights]::FullControl
            $ace2.AceFlags = [Uint32][System.Security.AccessControl.AceFlags]::None
            $ace2.AceType = [Uint32][System.Security.AccessControl.AceType]::AccessAllowed
            $ace2.Trustee = $trustee
            #endregion

            $sd = ([WMIClass]'Win32_SecurityDescriptor').PSBase.CreateInstance()
            $sd.ControlFlags = 4
            $sd.DACL = $ace, $ace2
            $sd.group = $trustee
            $sd.owner = $trustee

            $lss = Get-WmiObject -ClassName Win32_LogicalShareSecuritySetting -Filter "Name='$Name'"
            $null = $lss.SetSecurityDescriptor($sd)

            $lss = Get-WmiObject -ClassName Win32_LogicalShareSecuritySetting -Filter "Name='$Name'"
            $null = $lss.SetSecurityDescriptor($sd)
            
            Write-Verbose "Share permissions set to 'FullControl' for 'Everyone' on '$Name'"
        }
        Catch {
            throw "Failed setting 'FullControl' for 'Everyone' on share '$Name': $_"
        }
    }

    Function Test-AccessBasedEnumerationHC {
        <# 
        .SYNOPSIS   
            Check if a shared folder has ABE enabled.

        .DESCRIPTION
            Check if a shared folder has Access Based Enumeration enabled. Returns
            TRUE if ABE is enabled or FALSE when it's not.

        .PARAMETER ComputerName
            Specifies the target computer.

        PARAMETER Name
            Name of the share.

        .EXAMPLE
            Retrieve the Access Based Enumeration status for the shared folder 'Log'.
                        
            Test-AccessBasedEnumerationHC -Name Log

            The boolean 'True' is returned because the share has ABE set to enabled.

        .NOTE
            CHANGELOG
                2016/09/28 Function born
                2018/06/29 Rewrote the function to just return a boolean
                           Optimized for speed by using the method '.Foreach'
                           Removed object creation as it's overkill
                           Improved help and added OutputType

            AUTHOR Brecht.Gijbels@heidelbergcement.com #>

        [CmdLetbinding()]
        [OutputType([Boolean])] 
        Param (
            [Parameter(Mandatory, ValueFromPipeline)]
            [String[]]$Name,
            [ValidateNotNullOrEmpty()]
            [String]$ComputerName = $env:COMPUTERNAME
        )
      
        Process {
            $Name.ForEach( {
                    Try {
                        $ShareInfo = [NetApi32]::NetShareGetInfo_1005($ComputerName, $_)

                        if ($ShareInfo.Shi1005_flags -eq ($ShareInfo.Shi1005_flags -bor [Shi1005_flags]::SHI1005_FLAGS_ACCESS_BASED_DIRECTORY_ENUM)) {
                            $true
                        }
                        else {
                            $false
                        }
                    }
                    Catch {
                        throw "Failed retrieving Access Based Enumeration status on share '$S' for '$ComputerName': $_"
                    }
                })
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
            [Int]$Major,
            [Int]$Minor
        )

        (($Major -gt 5) -or (($Major -eq 5) -and ($Minor -ge 1)))
    }
}

Process {
    <# 
        the require parameter at the top is not supported with `Invoke-Command -FilePath`
        https://stackoverflow.com/questions/51185882/invoke-command-ignores-requires-in-the-script-file
    #>
    
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
    if (-not (Test-IsRequiredPowerShellVersionHC -Major $PSVersionTable.PSVersion.Major -Minor $PSVersionTable.PSVersion.Minor)) {
        Return [PSCustomObject]@{
            Type        = 'FatalError'
            Name        = 'PowerShell version'
            Description = "PowerShell version 5.1 or higher is required to be able to use advanced methods."
            Value       = "PowerShell $($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor)"
        }
    }
    #endregion

    #region Require at least .NET 4.6.2
    $DotNet = Get-ChildItem 'HKLM:SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\' | 
    Get-ItemPropertyValue -Name Release | ForEach-Object { $_ -ge 394802 }
 
    if (-not $DotNet) {
        Return [PSCustomObject]@{
            Type        = 'FatalError'
            Name        = '.NET Framework version'
            Description = "Microsoft .NET Framework version 4.6.2 or higher is required to be able to traverse long path names and use advanced PowerShell methods."
            Value       = $null
        }
    }
    #endregion

    #region Load NetApi32
    Try {
        if (-not ([System.Management.Automation.PSTypeName]'NetApi32').Type) {
            Add-Type -TypeDefinition @"
                    using System;
                    using System.Collections.Generic;
                    using System.Runtime.InteropServices;
                    using System.Text;
 
 
                    public enum Share_Type : uint
                    {
                        STYPE_DISKTREE = 0x00000000,   // Disk Drive
                        STYPE_PRINTQ = 0x00000001,   // Print Queue
                        STYPE_DEVICE = 0x00000002,   // Communications Device
                        STYPE_IPC = 0x00000003,   // InterProcess Communications
                        STYPE_SPECIAL = 0x80000000,   // Special share types (C$, ADMIN$, IPC$, etc)
                        STYPE_TEMPORARY = 0x40000000   // Temporary share 
                    }
 
                    public enum Share_ReturnValue : int
                    {
                        NERR_Success = 0,
                        ERROR_ACCESS_DENIED = 5,
                        ERROR_NOT_ENOUGH_MEMORY = 8,
                        ERROR_INVALID_PARAMETER = 87,
                        ERROR_INVALID_LEVEL = 124, // unimplemented level for info
                        ERROR_MORE_DATA = 234,
                        NERR_BufTooSmall = 2123, // The API return buffer is too small.
                        NERR_NetNameNotFound = 2310 // This shared resource does not exist.
                    }
 
                    [System.Flags]
                    public enum Shi1005_flags
                    {
                        SHI1005_FLAGS_DFS = 0x0001,  // Part of a DFS tree (Cannot be set)
                        SHI1005_FLAGS_DFS_ROOT = 0x0002,  // Root of a DFS tree (Cannot be set)
                        SHI1005_FLAGS_RESTRICT_EXCLUSIVE_OPENS = 0x0100,  // Disallow Exclusive file open
                        SHI1005_FLAGS_FORCE_SHARED_DELETE = 0x0200,  // Open files can be force deleted
                        SHI1005_FLAGS_ALLOW_NAMESPACE_CACHING = 0x0400,  // Clients can cache the namespace
                        SHI1005_FLAGS_ACCESS_BASED_DIRECTORY_ENUM = 0x0800,  // Only directories for which a user has FILE_LIST_DIRECTORY will be listed
                        SHI1005_FLAGS_FORCE_LEVELII_OPLOCK = 0x1000,  // Prevents exclusive caching
                        SHI1005_FLAGS_ENABLE_HASH = 0x2000,  // Used for server side support for peer caching
                        SHI1005_FLAGS_ENABLE_CA = 0X4000   // Used for Clustered shares
                    }
 
                    public static class NetApi32
                    {
 
                        // ********** Structures **********
 
                        // SHARE_INFO_502
                        [StructLayout(LayoutKind.Sequential)]
                        public struct SHARE_INFO_502
                        {
                            [MarshalAs(UnmanagedType.LPWStr)]
                            public string shi502_netname;
                            public uint shi502_type;
                            [MarshalAs(UnmanagedType.LPWStr)]
                            public string shi502_remark;
                            public Int32 shi502_permissions;
                            public Int32 shi502_max_uses;
                            public Int32 shi502_current_uses;
                            [MarshalAs(UnmanagedType.LPWStr)]
                            public string shi502_path;
                            public IntPtr shi502_passwd;
                            public Int32 shi502_reserved;
                            public IntPtr shi502_security_descriptor;
                        }
 
                        // SHARE_INFO_1005
                        [StructLayout(LayoutKind.Sequential)]
                        public struct SHARE_INFO_1005
                        {
                            public Int32 Shi1005_flags;
                        }
 
       
 
                        private class unmanaged
                        {
 
                            //NetShareGetInfo
                            [DllImport("Netapi32.dll", SetLastError = true)]
                            internal static extern int NetShareGetInfo(
                                [MarshalAs(UnmanagedType.LPWStr)] string serverName,
                                [MarshalAs(UnmanagedType.LPWStr)] string netName,
                                Int32 level,
                                ref IntPtr bufPtr
                            );
 
                            [DllImport("Netapi32.dll", SetLastError = true)]
                            public extern static Int32 NetShareSetInfo(
                                [MarshalAs(UnmanagedType.LPWStr)] string servername,
                                [MarshalAs(UnmanagedType.LPWStr)] string netname, Int32 level,IntPtr bufptr, out Int32 parm_err);
 
 
                        }
 
                        // ***** Functions *****
                        public static SHARE_INFO_502 NetShareGetInfo_502(string ServerName, string ShareName)
                        {
                            Int32 level = 502;
                            IntPtr lShareInfo = IntPtr.Zero;
                            SHARE_INFO_502 shi502_Info = new SHARE_INFO_502();
                            Int32 result = unmanaged.NetShareGetInfo(ServerName, ShareName, level, ref lShareInfo);
                            if ((Share_ReturnValue)result == Share_ReturnValue.NERR_Success)
                            {
                                shi502_Info = (SHARE_INFO_502)Marshal.PtrToStructure(lShareInfo, typeof(SHARE_INFO_502));
                            }
                            else
                            {
                                throw new Exception("Unable to get 502 structure.  Function returned: " + (Share_ReturnValue)result);
                            }
                            return shi502_Info;
                        }
 
                        public static SHARE_INFO_1005 NetShareGetInfo_1005(string ServerName, string ShareName)
                        {
                            Int32 level = 1005;
                            IntPtr lShareInfo = IntPtr.Zero;
                            SHARE_INFO_1005 shi1005_Info = new SHARE_INFO_1005();
                            Int32 result = unmanaged.NetShareGetInfo(ServerName, ShareName, level, ref lShareInfo);
                            if ((Share_ReturnValue)result == Share_ReturnValue.NERR_Success)
                            {
                                shi1005_Info = (SHARE_INFO_1005)Marshal.PtrToStructure(lShareInfo, typeof(SHARE_INFO_1005));
                            }
                            else
                            {
                                throw new Exception("Unable to get 1005 structure.  Function returned: " + (Share_ReturnValue)result);
                            }
                            return shi1005_Info;
                        }
 
                        public static int NetShareSetInfo_1005(string ServerName, string ShareName, SHARE_INFO_1005 shi1005_Info) //  Int32 Shi1005_flags
                        {
                            Int32 level = 1005;
                            Int32 err;
             
                            IntPtr ptr = Marshal.AllocHGlobal(Marshal.SizeOf(shi1005_Info));
                            Marshal.StructureToPtr(shi1005_Info, ptr, false);
 
                            var result = unmanaged.NetShareSetInfo(ServerName, ShareName, level, ptr, out err);
 
                            return result;
                        }
 
                    }
"@
        }
    }
    Catch {
        throw "Failed loading type defintion 'NetApi32': $_"
    }
    #endregion

    $Shares = Get-SharePermissionHC

    #region Get unique Path
    $Path = $Path | Sort-Object -Unique
    #endregion

    #region Set Access Based Enumeration
    Try {
        $AbeCorrected = @{}
        foreach ($P in $Path) {
            ($Shares).Where( { (
                        (($_.Path -like "$P\*") -or ($_.Path -eq $P)) -and 
                        (Test-AccessBasedEnumerationHC -Name $_.Name) -ne $Flag) }).ForEach( {
                        
                    Set-AccessBasedEnumerationHC -Name $_.Name -Flag $Flag
                    $AbeCorrected.Add($_.Name, $_.Path)
                })
        }

        if ($AbeCorrected.Count -ne 0) {
            [PSCustomObject]@{
                Type        = 'Warning'
                Name        = 'Access Based Enumeration'
                Description = "Access Based Enumeration should be set to '$flag'. This will hide files and folders where the users don't have access to. We fixed this now."
                Value       = $AbeCorrected
            }
        }
    }
    Catch {
        throw "Failed setting Access Based Enumeration to '$Flag' for path '$Path' on '$env:COMPUTERNAME': $_"
    }
    #endregion
    
    #region Set share permissions to to 'FullControl for Administrators' and 'Read & Executed for Authenticated users'
    Try {
        $SharePermCorrected = @{}
        foreach ($P in $Path) {
            ($Shares).Where( {
                    (($_.Path -like "$P\*") -or ($_.Path -eq $P)) -and
                    (
                        ($_.Acl.Count -ne 2) -or
                        ($_.Acl.'BUILTIN\Administrators' -ne 'FullControl') -or 
                        ($_.Acl.'NT AUTHORITY\Authenticated users' -ne 'Change')
                    )
                }).ForEach( {
                    Set-SharePermissionsHC -Name $_.Name
                    $SharePermCorrected.Add($_.Name, $_.ACL)
                })
        }

        if ($SharePermCorrected.Count -ne 0) {
            [PSCustomObject]@{
                Type        = 'Warning'
                Name        = 'Share permissions'
                Description = "The share permissions are now set to 'Administrators: FullControl' and 'Authenticated users: Change'. The effective permissions are managed on NTFS level."
                Value       = $SharePermCorrected
            }
        }
    }
    Catch {
        throw "Failed setting share permissions on path '$Path' on '$env:COMPUTERNAME': $_"
    }
    #endregion
}