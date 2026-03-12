#Requires -Version 7
#Requires -RunAsAdministrator

<#
    .SYNOPSIS
        Scan an NTFS folder structure and create, check or fix the permissions.

    .DESCRIPTION
        All files and folders are checked with the permissions defined in the
        matrix.

    .PARAMETER Path
        The parent folder on the localhost where the folder tree starts.

    .PARAMETER Action
        Valid values:
        - New   : Creates a new folder structure with the correct permissions
        - Check : Only check if the permissions are correct
        - Fix   : Check and fix incorrect permissions

    .PARAMETER Matrix
        The array containing the correct folder names and their permissions.

    .PARAMETER DetailedLog
        When incorrect permissions are found only the FullName of the path is
        reported. However, when DetailedLog is enabled the current and desired
        permissions are reported.

        For performance reason, only enable this for troubleshooting.
#>

[OutputType([PSCustomObject[]])]
[CmdLetBinding()]
param (
    [Parameter(Mandatory)]
    [String]$Path,
    [Parameter(Mandatory)]
    [ValidateSet('New', 'Check', 'Fix')]
    [String]$Action,
    [Parameter(Mandatory)]
    [PSCustomObject[]]$Matrix,
    [Parameter(Mandatory)]
    [Int]$JobThrottleLimit,
    [Boolean]$DetailedLog
)

begin {
    #region Function New-AceHC
    function New-AceHC {
        [CmdLetBinding()]
        param (
            [Parameter(Mandatory)]
            [ValidateSet('L', 'R', 'W', 'F', 'M')]
            [String]$Access,

            [Parameter(Mandatory)]
            [String]$Name,

            [Parameter(Mandatory)]
            [ValidateSet('Folder', 'InheritedFile', 'InheritedFolder')]
            [String]$Type
        )

        $identity = "$env:USERDOMAIN\$Name"
        $allow = [System.Security.AccessControl.AccessControlType]::Allow
        $rules = [System.Collections.Generic.List[System.Security.AccessControl.FileSystemAccessRule]]::new()

        $createRule = {
            param($rights, $inheritance, $propagation)
            $rules.Add([System.Security.AccessControl.FileSystemAccessRule]::new($identity, $rights, $inheritance, $propagation, $allow))
        }

        switch ($Access) {
            'L' {
                if ($Type -in 'Folder', 'InheritedFolder') {
                    &$createRule 'ReadAndExecute' 'ContainerInherit' 'None'
                }
            }
            'W' {
                if ($Type -eq 'Folder') {
                    &$createRule 'CreateFiles, AppendData, DeleteSubdirectoriesAndFiles, ReadAndExecute, Synchronize' 'None' 'InheritOnly'
                    &$createRule 'DeleteSubdirectoriesAndFiles, Modify, Synchronize' 'ContainerInherit, ObjectInherit' 'InheritOnly'
                }
                elseif ($Type -eq 'InheritedFolder') {
                    &$createRule 'DeleteSubdirectoriesAndFiles, Modify, Synchronize' 'ContainerInherit, ObjectInherit' 'InheritOnly'
                }
                elseif ($Type -eq 'InheritedFile') {
                    &$createRule 'DeleteSubdirectoriesAndFiles, Modify, Synchronize' 'None' 'None'
                }
            }
            default {
                $rights = switch ($Access) {
                    'R' { 'ReadAndExecute' }
                    'F' { 'FullControl' }
                    'M' { 'Modify' }
                }

                if ($Type -in 'Folder', 'InheritedFolder') {
                    &$createRule $rights 'ContainerInherit, ObjectInherit' 'None'
                }
                elseif ($Type -eq 'InheritedFile') {
                    &$createRule $rights 'None' 'None'
                }
            }
        }
        return $rules.ToArray()
    }
    #endregion

    #region Function Test-AclEqualHC (Main Thread)
    function Test-AclEqualHC {
        [OutputType([Boolean])]
        param (
            [Parameter(Mandatory)]
            [AllowNull()]
            [AllowEmptyCollection()]
            [System.Object[]]$ReferenceAce = @(),

            [Parameter(Mandatory)]
            [AllowNull()]
            [AllowEmptyCollection()]
            [System.Object[]]$DifferenceAce = @()
        )

        try {
            if ($ReferenceAce.Count -ne $DifferenceAce.Count) { return $false }

            # OPTIMIZATION: Use O(1) HashSet for fast matching instead of nested loops
            $refSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($R in $ReferenceAce) {
                [void]$refSet.Add("$([int]$R.FileSystemRights)|$([int]$R.AccessControlType)|$($R.IdentityReference.ToString())|$([int]$R.InheritanceFlags)")
            }

            foreach ($D in $DifferenceAce) {
                $id = "$([int]$D.FileSystemRights)|$([int]$D.AccessControlType)|$($D.IdentityReference.ToString())|$([int]$D.InheritanceFlags)"
                if (-not $refSet.Contains($id)) { return $false }
            }
            return $true
        }
        catch {
            throw "Failed testing the ACL for equality: $_"
        }
    }
    #endregion

    #region ScriptBlock InheritedPermissionsScriptBlock
    $inheritedPermissionsScriptBlock = {
        [OutputType([PSCustomObject[]])]
        [CmdLetBinding()]
        param (
            [Parameter(Mandatory)]
            [String]$Path,
            [Parameter(Mandatory)]
            [ValidateSet('Check', 'Fix')]
            [String]$Action,

            [Parameter(Mandatory)]
            [AllowNull()]
            [AllowEmptyCollection()]
            [System.Object[]]$FolderAclAccessList = @(),

            [Parameter(Mandatory)]
            [AllowNull()]
            [AllowEmptyCollection()]
            [System.Object[]]$FileAclAccessList = @(),

            [Parameter(Mandatory)]
            [HashTable]$IgnoredFolderPaths,
            [Parameter(Mandatory)]
            [String]$TokenPrivileges,
            [Boolean]$DetailedLog
        )

        $ErrorActionPreference = 'Stop'

        try { Import-Module -Name 'Microsoft.PowerShell.Security' } catch { throw "Failed loading .NET library: $_" }

        # OPTIMIZATION: Setup HashSets ONCE per runspace to avoid repeating work for every file
        $folderRulesSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        if ($FolderAclAccessList) { foreach ($r in $FolderAclAccessList) { [void]$folderRulesSet.Add($r) } }

        $fileRulesSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        if ($FileAclAccessList) { foreach ($r in $FileAclAccessList) { [void]$fileRulesSet.Add($r) } }

        #region Function Test-AclEqualHC (Parallel Thread)
        function Test-AclEqualHC {
            [OutputType([Boolean])]
            param (
                [Parameter(Mandatory)]
                [System.Collections.Generic.HashSet[string]]$ReferenceSet,

                [Parameter(Mandatory)]
                [AllowNull()]
                [AllowEmptyCollection()]
                [System.Object[]]$DifferenceAce = @()
            )

            try {
                if ($ReferenceSet.Count -ne $DifferenceAce.Count) { return $false }

                foreach ($D in $DifferenceAce) {
                    # Generate the fingerprint using [int] to bypass slow string evaluations
                    $id = "$([int]$D.FileSystemRights)|$([int]$D.AccessControlType)|$($D.IdentityReference.ToString())|$([int]$D.InheritanceFlags)"
                    
                    if (-not $ReferenceSet.Contains($id)) { return $false }
                }
                return $true
            }
            catch {
                throw "Failed testing the ACL for equality: $_"
            }
        }
        #endregion

        #region Function Get-FolderContentHC
        function Get-FolderContentHC {
            param (
                [Parameter(Mandatory)]
                [String]$Path
            )

            try {
                Write-Verbose "Get content of folder '$Path'"
                $dirInfo = [System.IO.DirectoryInfo]::new($Path)
                $enumerator = $dirInfo.EnumerateFileSystemInfos()
            }
            catch {
                throw "Failed retrieving the folder content of '$Path': $_"
            }

            foreach ($child in $enumerator) {
                # Skip DFS links, Reparse Points and System directories
                if (
                    ($child.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -or
                    ($child.Attributes -band [System.IO.FileAttributes]::System)
                    # ($child.Attributes -band [System.IO.FileAttributes]::Hidden)
                ) {
                    continue
                }

                if ($IgnoredFolderPaths.ContainsKey($child.FullName)) { 
                    continue 
                }

                $isContainer = $child -is [System.IO.DirectoryInfo]

                $accessDenied = $false
                $acl = $null
                try {
                    # FAST .NET API Call bypassing PowerShell provider overhead
                    if ($isContainer) {
                        $acl = [System.IO.FileSystemAclExtensions]::GetAccessControl([System.IO.DirectoryInfo]$child)
                    }
                    else {
                        $acl = [System.IO.FileSystemAclExtensions]::GetAccessControl([System.IO.FileInfo]$child)
                    }
                }
                catch [System.UnauthorizedAccessException] {
                    $accessDenied = $true
                }
                catch {
                    # FALLBACK: Use classic Get-Acl if .NET method fails
                    try {
                        $acl = Get-Acl -LiteralPath $child.FullName -ErrorAction Stop
                    }
                    catch [System.UnauthorizedAccessException] {
                        $accessDenied = $true
                    }
                    catch {
                        if (-not (Test-Path -LiteralPath $child.FullName)) {
                            Write-Verbose "Item '$($child.FullName)' removed"
                            $Error.RemoveAt(0)
                        }
                        else {
                            $errorMessage = "Failed retrieving the ACL of '$child': $_"
                            
                            Write-Warning $errorMessage

                            if ($DetailedLog) {
                                $incorrectInheritedAcl[$child.FullName] = $errorMessage
                            }
                            else {
                                $incorrectInheritedAcl.Add($child.FullName)
                            }
                        }
                        continue
                    }
                }

                $testedInheritedFilesAndFolders[$child.FullName] = $true

                $diffAce = if (-not $accessDenied -and $acl) { @($acl.Access) } else { @() }

                if ($isContainer) {
                    if ($accessDenied -or (-not (Test-AclEqualHC -ReferenceSet $folderRulesSet -DifferenceAce $diffAce))) {
                        & $incorrectAclInheritedOnly
                    }

                    if ((-not $accessDenied) -or ($Action -eq 'Fix')) {
                        Get-FolderContentHC -Path $child.FullName
                    }
                }
                else {
                    if ($accessDenied -or (-not (Test-AclEqualHC -ReferenceSet $fileRulesSet -DifferenceAce $diffAce))) {
                        & $incorrectAclInheritedOnly
                    }
                }
            }
        }
        #endregion

        #region ScriptBlock IncorrectAclInheritedOnly
        $incorrectAclInheritedOnly = {
            Write-Warning "Incorrect ACL '$($child.FullName)'"

            if ($DetailedLog) {
                $incorrectInheritedAcl[$child.FullName] = if ($accessDenied) {
                    'Access Denied'
                }
                else {
                    $acl.AccessToString
                }
            }
            else {
                $incorrectInheritedAcl.Add($child.FullName)
            }

            if ($Action -eq 'Fix') {
                Write-Verbose "Set ACL to inherited only '$($child.FullName)'"

                if ($isContainer) {
                    $dirInfo = [System.IO.DirectoryInfo]::new($child.FullName)

                    if ($accessDenied) {
                        [TokenManipulator]::SetOwner($child.FullName, 'BUILTIN\Administrators')
                    }

                    try {
                        [System.IO.FileSystemAclExtensions]::SetAccessControl($dirInfo, $inheritedDirAcl)
                    }
                    catch [System.UnauthorizedAccessException] {
                        [TokenManipulator]::SetOwner($child.FullName, 'BUILTIN\Administrators')
                        [System.IO.FileSystemAclExtensions]::SetAccessControl($dirInfo, $inheritedDirAcl)
                    }
                }
                else {
                    $fileInfo = [System.IO.FileInfo]::new($child.FullName)

                    if ($accessDenied) {
                        [TokenManipulator]::SetOwner($child.FullName, 'BUILTIN\Administrators')
                    }

                    try {
                        [System.IO.FileSystemAclExtensions]::SetAccessControl($fileInfo, $inheritedFileAcl)
                    }
                    catch [System.UnauthorizedAccessException] {
                        [TokenManipulator]::SetOwner($child.FullName, 'BUILTIN\Administrators')
                        [System.IO.FileSystemAclExtensions]::SetAccessControl($fileInfo, $inheritedFileAcl)
                    }
                }
            }
        }
        #endregion

        try {
            #region Logging Setup
            $testedInheritedFilesAndFolders = @{ }

            if ($DetailedLog) {
                $incorrectInheritedAcl = @{ }
            }
            else {
                $incorrectInheritedAcl = [System.Collections.Generic.List[String]]::New()
            }
            #endregion

            #region Get super powers
            try {
                Write-Verbose 'Get super powers'

                if (-not ('TokenManipulator' -as [type])) {
                    try {
                        Add-Type $tokenPrivileges -ErrorAction Stop
                    }
                    catch {
                        if ($_.Exception.Message -notmatch 'already exists') {
                            throw $_
                        }
                    }
                }

                [void][TokenManipulator]::AddPrivilege('SeRestorePrivilege')
                [void][TokenManipulator]::AddPrivilege('SeBackupPrivilege')
                [void][TokenManipulator]::AddPrivilege('SeTakeOwnershipPrivilege')
            }
            catch { throw "Failed getting super powers: $_" }
            #endregion

            #region Create inherited folder and file acl
            Write-Verbose 'Inherited permissions'
            $builtinAdmin = [System.Security.Principal.NTAccount]'BUILTIN\Administrators'

            $inheritedDirAcl = New-Object System.Security.AccessControl.DirectorySecurity
            $inheritedDirAcl.SetOwner($builtinAdmin)
            $inheritedDirAcl.SetAccessRuleProtection($false, $false)

            $inheritedFileAcl = New-Object System.Security.AccessControl.FileSecurity
            $inheritedFileAcl.SetOwner($builtinAdmin)
            $inheritedFileAcl.SetAccessRuleProtection($false, $false)
            #endregion

            #region Check or fix folder and file permissions
            try { Get-FolderContentHC -Path $Path } catch { throw "Failed checking or setting the inheritance in folder '$Path': $_" }
            #endregion
        }
        catch { throw "Failed setting permissions for '$Path': $_" }
        finally {
            [PSCustomObject]@{ testedInheritedFilesAndFolders = $testedInheritedFilesAndFolders; IncorrectInheritedAcl = $incorrectInheritedAcl }
        }
    }
    #endregion

    #region TokenManipulator C# Class
    $tokenPrivileges = @'
using System;
using System.Runtime.InteropServices;
using System.Security.Principal;

public class TokenManipulator
{
    [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
    internal static extern bool AdjustTokenPrivileges(IntPtr htok, bool disall, ref TokPriv1Luid newst, int len, IntPtr prev, IntPtr relen);

    [DllImport("kernel32.dll", ExactSpelling = true)]
    internal static extern IntPtr GetCurrentProcess();

    [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
    internal static extern bool OpenProcessToken(IntPtr h, int acc, ref IntPtr phtok);

    [DllImport("advapi32.dll", SetLastError = true)]
    internal static extern bool LookupPrivilegeValue(string host, string name, ref long pluid);

    [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    internal static extern uint SetNamedSecurityInfo(string pObjectName, int objectType, uint securityInfo, byte[] psidOwner, byte[] psidGroup, IntPtr pDacl, IntPtr pSacl);

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    internal struct TokPriv1Luid
    {
        public int Count;
        public long Luid;
        public int Attr;
    }

    internal const int SE_PRIVILEGE_DISABLED = 0x00000000;
    internal const int SE_PRIVILEGE_ENABLED = 0x00000002;
    internal const int TOKEN_QUERY = 0x00000008;
    internal const int TOKEN_ADJUST_PRIVILEGES = 0x00000020;

    internal const uint OWNER_SECURITY_INFORMATION = 0x00000001;
    internal const int SE_FILE_OBJECT = 1;

    public static bool AddPrivilege(string privilege)
    {
        try
        {
            bool retVal;
            TokPriv1Luid tp;
            IntPtr hproc = GetCurrentProcess();
            IntPtr htok = IntPtr.Zero;
            retVal = OpenProcessToken(hproc, TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, ref htok);
            tp.Count = 1;
            tp.Luid = 0;
            tp.Attr = SE_PRIVILEGE_ENABLED;
            retVal = LookupPrivilegeValue(null, privilege, ref tp.Luid);
            retVal = AdjustTokenPrivileges(htok, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
            return retVal;
        }
        catch (Exception ex)
        {
            throw ex;
        }
    }

    public static bool RemovePrivilege(string privilege)
    {
        try
        {
            bool retVal;
            TokPriv1Luid tp;
            IntPtr hproc = GetCurrentProcess();
            IntPtr htok = IntPtr.Zero;
            retVal = OpenProcessToken(hproc, TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, ref htok);
            tp.Count = 1;
            tp.Luid = 0;
            tp.Attr = SE_PRIVILEGE_DISABLED;
            retVal = LookupPrivilegeValue(null, privilege, ref tp.Luid);
            retVal = AdjustTokenPrivileges(htok, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
            return retVal;
        }
        catch (Exception ex)
        {
            throw ex;
        }
    }

    public static void SetOwner(string path, string accountName)
    {
        NTAccount account = new NTAccount(accountName);
        SecurityIdentifier sid = (SecurityIdentifier)account.Translate(typeof(SecurityIdentifier));
        byte[] sidBytes = new byte[sid.BinaryLength];
        sid.GetBinaryForm(sidBytes, 0);

        uint result = SetNamedSecurityInfo(path, SE_FILE_OBJECT, OWNER_SECURITY_INFORMATION, sidBytes, null, IntPtr.Zero, IntPtr.Zero);
        if (result != 0)
        {
            throw new System.ComponentModel.Win32Exception((int)result);
        }
    }
}
'@
    #endregion
}

process {
    try {
        $ErrorActionPreference = 'Stop'

        #region Pre-process the Matrix properties
        $missingFolders = [System.Collections.Generic.List[String]]::New()

        if ($Matrix) {
            foreach ($M in $Matrix) {
                if (-not $M.PSObject.Properties.Match('Parent').Count) {
                    $M | Add-Member -NotePropertyName 'Parent' -NotePropertyValue $false
                }
                if (-not $M.PSObject.Properties.Match('Ignore').Count) {
                    $M | Add-Member -NotePropertyName 'Ignore' -NotePropertyValue $false
                }

                if (($null -ne $M.ACL) -and ($M.ACL -isnot [System.Collections.IDictionary])) {
                    $realHash = @{}
                    foreach ($prop in $M.ACL.PSObject.Properties) {
                        if ($prop.MemberType -match 'NoteProperty') { $realHash[$prop.Name] = $prop.Value }
                    }
                    $M.ACL = $realHash
                }
            }
        }
        #endregion

        #region Logging Setup
        $testedInheritedFilesAndFolders = @{ }

        if ($DetailedLog) {
            $incorrectAclNonInheritedFolders = @{ }
            $incorrectInheritedAcl = @{ }
        }
        else {
            $incorrectAclNonInheritedFolders = [System.Collections.Generic.List[String]]::New()
            $incorrectInheritedAcl = [System.Collections.Generic.List[String]]::New()
        }
        #endregion

        #region Get super powers
        try {
            Write-Verbose 'Get super powers'

            if (-not ('TokenManipulator' -as [type])) {
                try {
                    Add-Type $tokenPrivileges -ErrorAction Stop
                }
                catch {
                    if ($_.Exception.Message -notmatch 'already exists') {
                        throw $_
                    }
                }
            }

            [void][TokenManipulator]::AddPrivilege('SeRestorePrivilege')
            [void][TokenManipulator]::AddPrivilege('SeBackupPrivilege')
            [void][TokenManipulator]::AddPrivilege('SeTakeOwnershipPrivilege')
        }
        catch { throw "Failed getting super powers: $_" }
        #endregion

        #region Import library for .NET calls
        try { Import-Module -Name 'Microsoft.PowerShell.Security' -Force } catch { throw "Failed loading .NET library: $_" }
        #endregion

        #region Create the parent folder when action is New
        try {
            if ($Action -eq 'New') {
                try { $missingFolders.Add((New-Item -Path $Path -ItemType Directory -EA Stop).FullName) }
                catch {
                    $Error.RemoveAt(0)
                    return [PSCustomObject]@{
                        Type        = 'FatalError'
                        Name        = 'Parent folder exists already'
                        Description = "The folder defined as 'Path' in the worksheet 'Settings' cannot be present on the remote machine when 'Action=New' is used. Please use 'Action' with value 'Check' or 'Fix' instead."
                        Value       = $Path
                    }
                }
            }
            elseif (-not (Test-Path -LiteralPath $Path -PathType Container)) {
                return [PSCustomObject]@{
                    Type        = 'FatalError'
                    Name        = 'Parent folder missing'
                    Description = "The folder defined as 'Path' in the worksheet 'Settings' needs to be available on the remote machine. In case the folder structure needs to be created, please use 'Action=New' instead."
                    Value       = $Path
                }
            }

            Write-Verbose "Parent folder '$Path'"
        }
        catch { throw "Failed checking the existence of the parent folder: $_" }
        #endregion

        #region Add the FullName for each path
        foreach ($M in $Matrix) {
            $tmpPath = if ($M.Parent) { $Path } else { Join-Path -Path $Path -ChildPath $M.Path }
            $M.Path = $tmpPath
        }
        #endregion

        #region Remove ignored folders from the matrix
        $ignoredFolders, $Matrix = $Matrix.Where( { $_.Ignore }, 'Split')
        $ignoredFolderPaths = @{}

        if ($ignoredFolders) {
            $IgnoredFolders.Path.ForEach({
                    Write-Verbose "Ignored folder '$_'"
                    $ignoredFolderPaths[$_] = $true
                })

            [PSCustomObject]@{
                Type        = 'Information'
                Name        = 'Ignored folder'
                Description = "All rows in the worksheet 'Permissions' that have the character 'i' defined are ignored. These folders are not checked for incorrect permissions."
                Value       = $IgnoredFolders.Path
            }
        }
        #endregion

        #region Inaccessible files Regex
        $FoldersListOnlyAclRegex = $Matrix.Where({ (-not ($_.Acl.Values.Where( { $_ -ne 'L' }))) -and ($_.ACL.Count -ne 0) }).ForEach( { [Regex]::Escape("$_") }) -join '|'
        $FoldersWithPermissionsRegex = $Matrix.Where( { ($_.Acl.Values.Where( { $_ -ne 'L' })) }).ForEach( { [Regex]::Escape("$_") }) -join '|'
        #endregion

        #region Create file and folder ACL for each path in the matrix
        try {
            Write-Verbose "Create ACE 'BUILTIN\Administrators' : 'FullControl'"
            $builtinAdmin = [System.Security.Principal.NTAccount]'BUILTIN\Administrators'

            $adminFullControlAce = @{
                Folder = New-Object System.Security.AccessControl.FileSystemAccessRule($builtinAdmin, [System.Security.AccessControl.FileSystemRights]::FullControl, [System.Security.AccessControl.InheritanceFlags]'ContainerInherit,ObjectInherit', [System.Security.AccessControl.PropagationFlags]::None, [System.Security.AccessControl.AccessControlType]::Allow)
                File   = New-Object System.Security.AccessControl.FileSystemAccessRule($builtinAdmin, [System.Security.AccessControl.FileSystemRights]::FullControl, [System.Security.AccessControl.AccessControlType]::Allow)
            }

            foreach ($M in $Matrix) {
                $M | Add-Member -NotePropertyMembers @{ FolderAcl = $null; InheritedFileAcl = $null; InheritedFolderAcl = $null }
            }

            $Matrix.Where( { $_.ACL.Count -eq 0 }).ForEach( { $_.ACL = $null })

            $aceCache = @{ }

            foreach ($M in $Matrix.Where( { $_.ACL })) {
                Write-Verbose "Create ACL for path '$($M.Path)'"

                $acl = @{
                    Folder          = New-Object System.Security.AccessControl.DirectorySecurity
                    InheritedFolder = New-Object System.Security.AccessControl.DirectorySecurity
                    InheritedFile   = New-Object System.Security.AccessControl.FileSecurity
                }

                $acl.Folder.SetAccessRuleProtection($true, $false)
                $acl.Folder.SetOwner($builtinAdmin)

                $acl.InheritedFolder.SetAccessRuleProtection($false, $false)
                $acl.InheritedFolder.SetOwner($builtinAdmin)

                $acl.InheritedFile.SetAccessRuleProtection($false, $false)
                $acl.InheritedFile.SetOwner($builtinAdmin)

                $M.ACL.GetEnumerator().Foreach({
                        try {
                            $ID = "$($_.Key)@$($_.Value)"

                            if (-not $aceCache.ContainsKey($ID)) {
                                $param = @{ Access = $_.Value; Name = $_.Key }
                                $aceCache[$ID] = @{
                                    Folder          = @( New-AceHC @param -Type 'Folder' )
                                    InheritedFolder = @( New-AceHC @param -Type 'InheritedFolder' )
                                    InheritedFile   = @( New-AceHC @param -Type 'InheritedFile' )
                                }
                            }

                            $aceCache[$ID]['Folder'].ForEach({ $acl.Folder.AddAccessRule($_) })
                            $aceCache[$ID]['InheritedFolder'].ForEach({ $acl.InheritedFolder.AddAccessRule($_) })
                            $aceCache[$ID]['InheritedFile'].ForEach({ $acl.InheritedFile.AddAccessRule($_) })
                        }
                        catch { throw "AD object '$($ID.split('@')[0])' with permission character '$($ID.split('@')[1])' probably doesn't exist in AD: $_" }
                    })

                $acl.Folder.AddAccessRule($adminFullControlAce.Folder)
                $acl.InheritedFolder.AddAccessRule($adminFullControlAce.Folder)
                $acl.InheritedFile.AddAccessRule($adminFullControlAce.File)

                $M.FolderAcl = $acl.Folder
                $M.inheritedFolderAcl = $acl.InheritedFolder
                $M.inheritedFileAcl = $acl.InheritedFile
            }
        }
        catch { throw "Failed creating the AccessControlList: $_" }
        #endregion

        #region Create Missing Folders (Check/Fix Matrix)
        try {
            $pathsToCreate = @()
            foreach ($M in $Matrix) {
                if (($M.Parent -eq $false) -and (-not (Test-Path -LiteralPath $M.Path -PathType Container))) {
                    $pathsToCreate += $M.Path
                }
            }

            foreach ($nonExistingPath in $pathsToCreate) {
                if ($Action -eq 'Check') {
                    Write-Verbose "Missing folder '$nonExistingPath'"
                    $missingFolders.Add($nonExistingPath)
                }
                else {
                    Write-Verbose "Create missing folder '$nonExistingPath'"
                    $missingFolders.Add((New-Item -Path $nonExistingPath -ItemType Directory -Force -EA Stop).FullName)
                }
            }

            if ($Action -eq 'Check' -and $missingFolders.Count -gt 0) {
                $Matrix = $Matrix.Where({ $_.Path -notin $missingFolders })
            }

            if ($missingFolders.Count -ne 0) {
                $Obj = [PSCustomObject]@{
                    Type        = 'Warning'
                    Name        = $null
                    Description = $null
                    Value       = $missingFolders.ToArray()
                }

                switch ($Action) {
                    'New' { $Obj.Name = 'Child folder created'; $Obj.Description = "All folders defined in the worksheet 'Permissions' have been created with the correct permissions underneath the parent folder defined in the worksheet 'Settings'."; break }
                    'Fix' { $Obj.Name = 'Child folder created'; $Obj.Description = 'The missing folders underneath the parent folder have been created.'; break }
                    'Check' { $Obj.Name = 'Child folder missing'; $Obj.Description = "Not all folders defined in the worksheet 'Permissions' were found underneath the parent folder."; break }
                    default { throw "Action '$_' is not supported." }
                }

                $Obj
            }
            else { Write-Verbose 'All folders present, no missing folders' }
        }
        catch { throw "Failed checking/creating the missing child folders: $_" }
        #endregion

        #region Non-Inherited folder permissions check and apply
        $testedNonInheritedFolders = @{}
        Write-Verbose 'Folders with ACL in the matrix that are not ignored'

        [array]$foldersWithAcl = $Matrix.Where({ ($_.FolderAcl) -and (-not $_.ignore) }) | Sort-Object -Property 'Path'

        foreach ($folder in $foldersWithAcl) {
            try {
                $ignoredFolderPaths[$folder.Path] = $true
                Write-Verbose "Matrix ACL folder '$($folder.Path)'"

                $dirInfo = [System.IO.DirectoryInfo]::new($folder.Path)
                $testedNonInheritedFolders[$folder.Path] = $folder

                $accessDenied = $false
                $acl = $null
                try {
                    # FAST .NET API Call bypassing PowerShell provider overhead
                    $acl = [System.IO.FileSystemAclExtensions]::GetAccessControl($dirInfo)
                }
                catch [System.UnauthorizedAccessException] {
                    $accessDenied = $true
                }
                catch {
                    # FALLBACK: Use classic Get-Acl if .NET method fails
                    try {
                        $acl = Get-Acl -LiteralPath $folder.Path -ErrorAction Stop
                    }
                    catch [System.UnauthorizedAccessException] {
                        $accessDenied = $true
                    }
                }

                $diffAce = if (-not $accessDenied -and $acl) { @($acl.Access) } else { @() }

                if ($accessDenied -or (-not $acl.AreAccessRulesProtected) -or (-not (Test-AclEqualHC -ReferenceAce ($folder.FolderAcl).Access -DifferenceAce $diffAce))) {
                    Write-Warning "Incorrect folder ACL '$($folder.Path)'"

                    #region Log Incorrect ACL
                    if ($Action -ne 'New') {
                        if ($DetailedLog) {
                            $incorrectAclNonInheritedFolders[$folder.Path] = @{
                                'Old' = if ($accessDenied) { 'Access Denied' } else { $acl.AccessToString }
                                'New' = ($folder.FolderAcl).AccessToString
                            }
                        }
                        else {
                            $incorrectAclNonInheritedFolders.Add($folder.Path)
                        }
                    }
                    #endregion

                    #region Set corrected ACL
                    if ($Action -ne 'Check') {
                        Write-Verbose 'Set correct ACL'

                        if ($accessDenied) { [TokenManipulator]::SetOwner($folder.Path, 'BUILTIN\Administrators') }

                        $newAcl = [System.Security.AccessControl.DirectorySecurity]::new()
                        $newAcl.SetOwner($builtinAdmin)
                        $newAcl.SetAccessRuleProtection($true, $false)
                        foreach ($rule in $folder.FolderAcl.Access) { $newAcl.AddAccessRule($rule) }

                        try {
                            [System.IO.FileSystemAclExtensions]::SetAccessControl($dirInfo, $newAcl)
                        }
                        catch [System.UnauthorizedAccessException] {
                            [TokenManipulator]::SetOwner($folder.Path, 'BUILTIN\Administrators')
                            [System.IO.FileSystemAclExtensions]::SetAccessControl($dirInfo, $newAcl)
                        }

                        Write-Verbose 'ACL corrected'
                    }
                    #endregion
                }
            }
            catch { throw "Failed checking/setting the permissions on non inherited folder '$($folder.Path)': $_" }
        }

        if ($incorrectAclNonInheritedFolders.Count -ne 0) {
            [PSCustomObject]@{
                Type        = 'Warning'
                Name        = 'Non inherited folder incorrect permissions'
                Description = "The folders that have permissions defined in the worksheet 'Permissions' are not matching with the permissions found on the folders of the remote machine."
                Value       = if ($DetailedLog) { $incorrectAclNonInheritedFolders } else { $incorrectAclNonInheritedFolders.ToArray() }
            }
        }
        #endregion

        #region Inherited folder and file permissions check and apply
        try {
            Write-Verbose 'Inherited permissions'
            if ($Action -ne 'New') {

                $ErrorActionPreference = 'Continue'
                $scriptBlockString = $inheritedPermissionsScriptBlock.ToString()

                $safeFolders = foreach ($folder in $foldersWithAcl) {
                    $extractRules = {
                        param($acl)
                        if (-not $acl) { return @() }
                        $arr = @()
                        foreach ($r in $acl.Access) {
                            # OPTIMIZATION: Extract to primitive string before sending into the runspace!
                            $arr += "$([int]$r.FileSystemRights)|$([int]$r.AccessControlType)|$($r.IdentityReference.ToString())|$([int]$r.InheritanceFlags)"
                        }
                        return $arr
                    }

                    [PSCustomObject]@{
                        Path               = $folder.Path
                        FolderRules        = &$extractRules $folder.InheritedFolderAcl
                        FileRules          = &$extractRules $folder.InheritedFileAcl
                        Action             = $Action
                        IgnoredFolderPaths = $ignoredFolderPaths
                        TokenPrivileges    = $tokenPrivileges
                        DetailedLog        = $DetailedLog
                        ScriptString       = $scriptBlockString
                    }
                }

                $jobResults = $safeFolders | ForEach-Object -Parallel {
                    $folderDto = $_

                    $params = @{
                        Path                = $folderDto.Path
                        Action              = $folderDto.Action
                        FolderAclAccessList = $folderDto.FolderRules
                        FileAclAccessList   = $folderDto.FileRules
                        IgnoredFolderPaths  = $folderDto.IgnoredFolderPaths
                        TokenPrivileges     = $folderDto.TokenPrivileges
                        DetailedLog         = $folderDto.DetailedLog
                    }

                    $rehydratedBlock = [scriptblock]::Create($folderDto.ScriptString)
                    & $rehydratedBlock @params

                } -ThrottleLimit $JobThrottleLimit

                foreach ($jobResult in $jobResults) {
                    foreach ($j in $jobResult.testedInheritedFilesAndFolders) {
                        foreach ($i in $j.GetEnumerator()) { $testedInheritedFilesAndFolders[$i.Key] = $i.Value }
                    }
                    foreach ($j in $jobResult.IncorrectInheritedAcl) {
                        if ($DetailedLog) {
                            foreach ($i in $j.GetEnumerator()) { $IncorrectInheritedAcl[$i.Key] = $i.Value }
                        }
                        else { $IncorrectInheritedAcl.Add($j) }
                    }
                }

                if ($IncorrectInheritedAcl.Count -ne 0) {
                    [PSCustomObject]@{
                        Type        = 'Warning'
                        Name        = 'Inherited permissions incorrect'
                        Description = "All folders that don't have permissions assigned to them in the worksheet 'Permissions' are supposed to inherit their permissions from the parent folder. Files can only inherit permissions from the parent folder and are not allowed to have explicit permissions."
                        Value       = if ($DetailedLog) { $IncorrectInheritedAcl } else { $IncorrectInheritedAcl.ToArray() }
                    }
                }
            }
        }
        catch { throw "Failed checking/setting the inheritance on folders and files: $_" }
        #endregion
    }
    catch { throw "Failed setting the permissions: $_" }
}