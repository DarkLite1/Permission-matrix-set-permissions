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
    function New-AceHC {
        <#
        .SYNOPSIS
            Convert an AD Object name and a permission character to a valid ACE.

        .DESCRIPTION
            Convert an AD Object name and a permission character to a valid Access Control List Entry.
        #>
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

        # 1. Define the static properties used by EVERY rule
        $identity = "$env:USERDOMAIN\$Name"
        $allow = [System.Security.AccessControl.AccessControlType]::Allow
        $rules = [System.Collections.Generic.List[System.Security.AccessControl.FileSystemAccessRule]]::new()

        # 2. Helper to instantly stamp out rules
        $createRule = {
            param($rights, $inheritance, $propagation)
            $rules.Add([System.Security.AccessControl.FileSystemAccessRule]::new($identity, $rights, $inheritance, $propagation, $allow))
        }

        # 3. Matrix logic
        switch ($Access) {
            'L' {
                if ($Type -in 'Folder', 'InheritedFolder') {
                    &$createRule 'ReadAndExecute' 'ContainerInherit' 'None'
                }
            }
            'W' {
                if ($Type -eq 'Folder') {
                    # Write on the root folder creates two distinct rules
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
                # Maps R, F, and M which all share identical inheritance logic
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

        # Return the generated rules (unwrapped safely if there's multiple)
        return $rules.ToArray()
    }

    function Test-AclEqualHC {
        <#
        .SYNOPSIS
            Compare two Access Control Entries.

        .DESCRIPTION
            Checks if two ACE's are matching. Returns True if both ACE lists are equal and
            False when they don't.
        #>

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
            # 1. High-speed early exit: If the number of ACEs don't match, they aren't equal
            if ($ReferenceAce.Count -ne $DifferenceAce.Count) { 
                return $false 
            }

            # 2. Iterate through and fail instantly on the first mismatch
            foreach ($D in $DifferenceAce) {
                $aclMatch = $ReferenceAce.Where({
                        ($D.FileSystemRights -eq $_.FileSystemRights) -and
                        ($D.AccessControlType -eq $_.AccessControlType) -and
                        ($D.IdentityReference -eq $_.IdentityReference) -and
                        ($D.InheritanceFlags -eq $_.InheritanceFlags) 
                        # ($D.PropagationFlags -eq $_.PropagationFlags) # NTFS alters this on files natively
                    }, 'First')

                # If no exact match was found for this specific rule, the ACLs are not equal
                if (-not $aclMatch) { 
                    return $false 
                }
            }

            # If we made it through the entire loop, every ACE perfectly matched!
            return $true
        }
        catch {
            throw "Failed testing the ACL for equality: $_"
        }
    }
    
    function Wait-MaxRunningJobsHC {
        <#
        .SYNOPSIS
            Limit how many jobs can run at the same time
        #>
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [System.Management.Automation.Job[]]$Name,
            
            [Parameter(Mandatory)]
            [Int]$MaxThreads,
            
            [Int]$MaxAllowedCpuLoadPercentage = 80
        )

        process {
            # 1. High-speed check: Wait for a job slot to open up FIRST
            while (@($Name).Where({ $_.State -eq 'Running' }).Count -ge $MaxThreads) {
                $null = Wait-Job -Job $Name -Any
            }

            # 2. Once a slot is open, ensure CPU isn't overwhelmed before launching the next one
            while ((Get-Counter '\Processor(_Total)\% Processor Time' -ErrorAction Ignore).CounterSamples.CookedValue -gt $MaxAllowedCpuLoadPercentage) {
                Start-Sleep -Milliseconds 500
            }
        }
    }

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
            [System.Collections.ArrayList]$FolderAclAccessList,
            [Parameter(Mandatory)]
            [System.Collections.ArrayList]$FileAclAccessList,
            [Parameter(Mandatory)]
            [HashTable]$IgnoredFolderPaths,
            [Parameter(Mandatory)]
            [String]$TokenPrivileges,
            [Boolean]$DetailedLog
        )

        $ErrorActionPreference = 'Stop'

        #region Import library for .NET calls
        try {
            Import-Module -Name 'Microsoft.PowerShell.Security'
        }
        catch {
            throw "Failed loading .NET library: $_"
        }
        #endregion

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

                foreach ($D in $DifferenceAce) {
                    $aclMatch = $ReferenceAce.Where({
                            ($D.FileSystemRights -eq $_.FileSystemRights) -and
                            ($D.AccessControlType -eq $_.AccessControlType) -and
                            ($D.IdentityReference -eq $_.IdentityReference) -and
                            ($D.InheritanceFlags -eq $_.InheritanceFlags)
                        }, 'First')

                    if (-not $aclMatch) { return $false }
                }
                return $true
            }
            catch {
                throw "Failed testing the ACL for equality: $_"
            }
        }

        function Get-FolderContentHC {
            param (
                [Parameter(Mandatory)]
                [String]$Path
            )

            try {
                Write-Verbose "Get content of folder '$Path'"
                $childItems = (Get-ChildItem -LiteralPath $Path -EA Stop).Where(
                    { -not ($IgnoredFolderPaths.ContainsKey($_.FullName)) }
                )
            }
            catch {
                throw "Failed retrieving the folder content of '$Path': $_"
            }

            foreach ($child in $childItems) {
                $accessDenied = $false
                try {
                    # Explicit casting prevents ambiguous overload errors!
                    $info = if ($child.PSIsContainer) {
                        [System.IO.DirectoryInfo]::new($child.FullName)
                    }
                    else { 
                        [System.IO.FileInfo]::new($child.FullName) 
                    }

                    $acl = [System.IO.FileSystemAclExtensions]::GetAccessControl($info)
                }
                catch [System.UnauthorizedAccessException] {
                    $accessDenied = $true
                }
                catch {
                    if (-not (Test-Path -LiteralPath $child.FullName)) {
                        Write-Verbose "Item '$child' removed"
                        $Error.RemoveAt(0)
                    }
                    else {
                        $ErrorActionPreference = 'Continue'
                        Write-Error "Failed retrieving the ACL of '$child': $_"
                        $ErrorActionPreference = 'Stop'
                    }
                    continue
                }

                # Only for Pester testing:
                $testedInheritedFilesAndFolders[$child.FullName] = $true

                # Safely get difference ACE or default to empty array
                $diffAce = if (-not $accessDenied -and $acl) { @($acl.Access) } else { @() }

                if ($child.PSIsContainer) {
                    if ($accessDenied -or (-not (Test-AclEqualHC -ReferenceAce $FolderAclAccessList -DifferenceAce $diffAce))) {
                        & $incorrectAclInheritedOnly
                    }

                    # If we had access natively, OR if we just forced a fix, recurse into it!
                    if ((-not $accessDenied) -or ($Action -eq 'Fix')) {
                        Get-FolderContentHC -Path $child.FullName
                    }
                }
                else {
                    if ($accessDenied -or (-not (Test-AclEqualHC -ReferenceAce $FileAclAccessList -DifferenceAce $diffAce))) {
                        & $incorrectAclInheritedOnly
                    }
                }
            }
        }

        $incorrectAclInheritedOnly = {
            Write-Warning "Incorrect ACL '$($child.FullName)'"
            
            #region Log
            if ($DetailedLog) {
                $incorrectInheritedAcl[$child.FullName] = if ($accessDenied) { 'Access Denied' } else { $acl.AccessToString }
            }
            else {
                $incorrectInheritedAcl.Add($child.FullName)
            }
            #endregion

            #region Set permissions
            if ($Action -eq 'Fix') {
                Write-Verbose "Set ACL to inherited only '$($child.FullName)'"

                if ($child.PSIsContainer) {
                    $dirInfo = [System.IO.DirectoryInfo]::new($child.FullName)
                    
                    # Break-in logic bypassing Set-Acl
                    if ($accessDenied) {
                        Write-Verbose "Access denied. Taking ownership of '$($child.FullName)'"
                        $takeOwnAcl = [System.Security.AccessControl.DirectorySecurity]::new()
                        $takeOwnAcl.SetOwner($builtinAdmin)
                        [System.IO.FileSystemAclExtensions]::SetAccessControl([System.IO.DirectoryInfo]$dirInfo, [System.Security.AccessControl.DirectorySecurity]$takeOwnAcl)
                    }
                    # Apply final matrix logic
                    [System.IO.FileSystemAclExtensions]::SetAccessControl([System.IO.DirectoryInfo]$dirInfo, [System.Security.AccessControl.DirectorySecurity]$inheritedDirAcl)
                }
                else {
                    $fileInfo = [System.IO.FileInfo]::new($child.FullName)
                    
                    # Break-in logic bypassing Set-Acl
                    if ($accessDenied) {
                        Write-Verbose "Access denied. Taking ownership of '$($child.FullName)'"
                        $takeOwnAcl = [System.Security.AccessControl.FileSecurity]::new()
                        $takeOwnAcl.SetOwner($builtinAdmin)
                        [System.IO.FileSystemAclExtensions]::SetAccessControl([System.IO.FileInfo]$fileInfo, [System.Security.AccessControl.FileSecurity]$takeOwnAcl)
                    }
                    # Apply final matrix logic
                    [System.IO.FileSystemAclExtensions]::SetAccessControl([System.IO.FileInfo]$fileInfo, [System.Security.AccessControl.FileSecurity]$inheritedFileAcl)
                }
            }
            #endregion
        }

        try {
            #region Logging
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
                Add-Type $TokenPrivileges
                [void][TokenManipulator]::AddPrivilege('SeRestorePrivilege')
                [void][TokenManipulator]::AddPrivilege('SeBackupPrivilege')
                [void][TokenManipulator]::AddPrivilege('SeTakeOwnershipPrivilege')
            }
            catch {
                throw "Failed getting super powers: $_"
            }
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
            try {
                Get-FolderContentHC -Path $Path
            }
            catch {
                throw "Failed checking or setting the inheritance in folder '$Path': $_"
            }
            #endregion
        }
        catch {
            throw "Failed setting permissions for '$Path': $_"
        }
        finally {
            [PSCustomObject]@{
                testedInheritedFilesAndFolders = $testedInheritedFilesAndFolders
                IncorrectInheritedAcl          = $incorrectInheritedAcl
            }
        }
    }

    $tokenPrivileges = @'
using System;
using System.Runtime.InteropServices;

public class TokenManipulator
{
[DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
internal static extern bool AdjustTokenPrivileges(IntPtr htok, bool disall,
ref TokPriv1Luid newst, int len, IntPtr prev, IntPtr relen);
[DllImport("kernel32.dll", ExactSpelling = true)]
internal static extern IntPtr GetCurrentProcess();
[DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
internal static extern bool OpenProcessToken(IntPtr h, int acc, ref IntPtr
phtok);
[DllImport("advapi32.dll", SetLastError = true)]
internal static extern bool LookupPrivilegeValue(string host, string name,
ref long pluid);
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
}
'@
}

process {
    try {
        $ErrorActionPreference = 'Stop'

        #region Rehydrate Deserialized ACLs
        # When passed over PS Remoting, Hashtables are stripped of their type and become flat PSCustomObjects.
        # This converts them safely back into native Hashtables
        if ($Matrix) {
            foreach ($M in $Matrix) {
                if (
                    ($null -ne $M.ACL) -and
                    ($M.ACL -isnot [System.Collections.IDictionary])
                ) {
                    $realHash = @{}
                    foreach ($prop in $M.ACL.PSObject.Properties) {
                        if ($prop.MemberType -match 'NoteProperty') {
                            $realHash[$prop.Name] = $prop.Value
                        }
                    }
                    $M.ACL = $realHash
                }
            }
        }
        #endregion

        #region Logging
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
            Add-Type $tokenPrivileges
            [void][TokenManipulator]::AddPrivilege('SeRestorePrivilege')
            [void][TokenManipulator]::AddPrivilege('SeBackupPrivilege')
            [void][TokenManipulator]::AddPrivilege('SeTakeOwnershipPrivilege')
        }
        catch {
            throw "Failed getting super powers: $_"
        }
        #endregion

        #region Import library for .NET calls
        try {
            Import-Module -Name 'Microsoft.PowerShell.Security' -Force
        }
        catch {
            throw "Failed loading .NET library: $_"
        }
        #endregion

        #region Create the parent folder when action is New
        try {
            $missingFolders = [System.Collections.Generic.List[String]]::New()

            if ($Action -eq 'New') {
                try {
                    $missingFolders.Add((New-Item -Path $Path -ItemType Directory -EA Stop).FullName)
                }
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
        catch {
            throw "Failed checking the existence of the parent folder: $_"
        }
        #endregion

        #region Add the FullName for each path
        foreach ($M in $Matrix) {
            $tmpPath = if ($M.Parent) { $Path }
            else { Join-Path -Path $Path -ChildPath $M.Path }

            $M.Path = $tmpPath
        }
        #endregion

        #region Remove ignored folders from the matrix
        $ignoredFolders, $Matrix = $Matrix.Where( { $_.Ignore }, 'Split')
        $ignoredFolderPaths = @{}

        if ($ignoredFolders) {
            $IgnoredFolders.Path.ForEach(
                {
                    Write-Verbose "Ignored folder '$_'"
                    $ignoredFolderPaths[$_] = $true
                }
            )

            [PSCustomObject]@{
                Type        = 'Information'
                Name        = 'Ignored folder'
                Description = "All rows in the worksheet 'Permissions' that have the character 'i' defined are ignored. These folders are not checked for incorrect permissions."
                Value       = $IgnoredFolders.Path
            }
        }
        #endregion

        #region Inaccessible files
        $FoldersListOnlyAclRegex = $Matrix.Where(
            {
                (-not ($_.Acl.Values.Where( { $_ -ne 'L' }))) -and
                ($_.ACL.Count -ne 0)
            }
        ).ForEach( {
                [Regex]::Escape("$_")
            }) -join '|'

        $FoldersWithPermissionsRegex = $Matrix.Where( {
                ($_.Acl.Values.Where( { $_ -ne 'L' }))
            }).ForEach( {
                [Regex]::Escape("$_")
            }) -join '|'
        #endregion

        #region Create file and folder ACL for each path in the matrix
        try {
            Write-Verbose "Create ACE 'BUILTIN\Administrators' : 'FullControl'"
            $builtinAdmin = [System.Security.Principal.NTAccount]'BUILTIN\Administrators'

            $adminFullControlAce = @{
                Folder = New-Object System.Security.AccessControl.FileSystemAccessRule(
                    $builtinAdmin,
                    [System.Security.AccessControl.FileSystemRights]::FullControl,
                    [System.Security.AccessControl.InheritanceFlags]'ContainerInherit,ObjectInherit',
                    [System.Security.AccessControl.PropagationFlags]::None,
                    [System.Security.AccessControl.AccessControlType]::Allow
                )
                File   = New-Object System.Security.AccessControl.FileSystemAccessRule(
                    $builtinAdmin,
                    [System.Security.AccessControl.FileSystemRights]::FullControl,
                    [System.Security.AccessControl.AccessControlType]::Allow
                )
            }

            foreach ($M in $Matrix) {
                $M | Add-Member -NotePropertyMembers @{
                    FolderAcl          = $null
                    InheritedFileAcl   = $null
                    InheritedFolderAcl = $null
                }
            }

            $Matrix.Where( { $_.ACL.Count -eq 0 }).ForEach( { $_.ACL = $null })

            #region Create the folder ACL's
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

                $M.ACL.GetEnumerator().Foreach(
                    {
                        try {
                            $ID = "$($_.Key)@$($_.Value)"

                            if (-not $aceCache.ContainsKey($ID)) {
                                $param = @{
                                    Access = $_.Value
                                    Name   = $_.Key
                                }

                                $aceCache[$ID] = @{
                                    Folder          = @(
                                        New-AceHC @param -Type 'Folder'
                                    )
                                    InheritedFolder = @(
                                        New-AceHC @param -Type 'InheritedFolder'
                                    )
                                    InheritedFile   = @(
                                        New-AceHC @param -Type 'InheritedFile'
                                    )
                                }
                            }

                            $aceCache[$ID]['Folder'].ForEach(
                                { $acl.Folder.AddAccessRule($_) }
                            )
                            $aceCache[$ID]['InheritedFolder'].ForEach(
                                { $acl.InheritedFolder.AddAccessRule($_) }
                            )
                            $aceCache[$ID]['InheritedFile'].ForEach(
                                { $acl.InheritedFile.AddAccessRule($_) }
                            )
                        }
                        catch {
                            throw "AD object '$($ID.split('@')[0])' with permission character '$($ID.split('@')[1])' probably doesn't exist in AD: $_"
                        }
                    }
                )

                $acl.Folder.AddAccessRule($adminFullControlAce.Folder)
                $acl.InheritedFolder.AddAccessRule($adminFullControlAce.Folder)
                $acl.InheritedFile.AddAccessRule($adminFullControlAce.File)

                $M.FolderAcl = $acl.Folder
                $M.inheritedFolderAcl = $acl.InheritedFolder
                $M.inheritedFileAcl = $acl.InheritedFile
            }
            #endregion
        }
        catch {
            throw "Failed creating the AccessControlList: $_"
        }
        #endregion

        #region Missing folders
        try {
            foreach (
                $nonExistingPath in
                ($Matrix.Where( { (-not (Test-Path -LiteralPath $_.Path -PathType Container)) -and (-not $_.Parent) })).Path) {

                if ($Action -eq 'Check') {
                    Write-Verbose "Missing folder '$nonExistingPath'"
                    $missingFolders.Add($nonExistingPath)
                    $Matrix = $Matrix.Where( { $_.Path -ne $nonExistingPath })
                }
                else {
                    Write-Verbose "Create missing folder '$nonExistingPath'"
                    $missingFolders.Add((New-Item -Path $nonExistingPath -ItemType Directory -Force -EA Stop).FullName)
                }
            }

            if ($missingFolders.Count -ne 0) {
                $Obj = [PSCustomObject]@{
                    Type        = 'Warning'
                    Name        = $null
                    Description = $null
                    Value       = $missingFolders.ToArray()
                }

                switch ($Action) {
                    'New' {
                        $Obj.Name = 'Child folder created'
                        $Obj.Description = "All folders defined in the worksheet 'Permissions' have been created with the correct permissions underneath the parent folder defined in the worksheet 'Settings'."
                        break
                    }
                    'Fix' {
                        $Obj.Name = 'Child folder created'
                        $Obj.Description = 'The missing folders underneath the parent folder have been created.'
                        break
                    }
                    'Check' {
                        $Obj.Name = 'Child folder missing'
                        $Obj.Description = "Not all folders defined in the worksheet 'Permissions' were found underneath the parent folder."
                        break
                    }
                    default {
                        throw "Action '$_' is not supported."
                    }
                }

                $Obj
            }
            else {
                Write-Verbose 'All folders present, no missing folders'
            }
        }
        catch {
            throw "Failed checking/creating the missing child folders: $_"
        }
        #endregion

        #region Non inherited folder permissions
        $testedNonInheritedFolders = @{}

        Write-Verbose 'Folders with ACL in the matrix that are not ignored'

        [array]$foldersWithAcl = $Matrix.Where(
            { ($_.FolderAcl) -and (-not $_.ignore) }
        ) | Sort-Object -Property 'Path'

        foreach ($folder in $foldersWithAcl) {
            try {
                $ignoredFolderPaths[$folder.Path] = $true

                Write-Verbose "Matrix ACL folder '$($folder.Path)'"
                
                $dirInfo = [System.IO.DirectoryInfo]::new($folder.Path)

                # Only for Pester testing:
                $testedNonInheritedFolders[$folder.Path] = $folder

                $accessDenied = $false
                try {
                    $acl = [System.IO.FileSystemAclExtensions]::GetAccessControl([System.IO.DirectoryInfo]$dirInfo)
                }
                catch [System.UnauthorizedAccessException] {
                    $accessDenied = $true
                }

                # Safely get difference ACE or default to empty array
                $diffAce = if (-not $accessDenied -and $acl) { @($acl.Access) } else { @() }

                if ($accessDenied -or (-not $acl.AreAccessRulesProtected) -or (-not (Test-AclEqualHC -ReferenceAce ($folder.FolderAcl).Access -DifferenceAce $diffAce))) {
                    Write-Warning "Incorrect folder ACL '$($folder.Path)'"
                    
                    #region Log
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

                    #region Set permissions
                    if ($Action -ne 'Check') {
                        Write-Verbose 'Set correct ACL'

                        # BREAK-IN LOGIC using strongly typed .NET 
                        if ($accessDenied) {
                            Write-Verbose "Access denied. Taking ownership of '$($folder.Path)'"
                            $takeOwnAcl = [System.Security.AccessControl.DirectorySecurity]::new()
                            $takeOwnAcl.SetOwner($builtinAdmin)
                            [System.IO.FileSystemAclExtensions]::SetAccessControl([System.IO.DirectoryInfo]$dirInfo, [System.Security.AccessControl.DirectorySecurity]$takeOwnAcl)
                        }

                        [System.IO.FileSystemAclExtensions]::SetAccessControl([System.IO.DirectoryInfo]$dirInfo, [System.Security.AccessControl.DirectorySecurity]$folder.FolderAcl)
                        Write-Verbose 'ACL corrected'
                    }
                    #endregion
                }
            }
            catch {
                throw "Failed checking/setting the permissions on non inherited folder '$($folder.Path)': $_"
            }
        }

        if ($incorrectAclNonInheritedFolders.Count -ne 0) {
            [PSCustomObject]@{
                Type        = 'Warning'
                Name        = 'Non inherited folder incorrect permissions'
                Description = "The folders that have permissions defined in the worksheet 'Permissions' are not matching with the permissions found on the folders of the remote machine."
                Value       = if ($DetailedLog) { $incorrectAclNonInheritedFolders }
                else { $incorrectAclNonInheritedFolders.ToArray() }
            }
        }
        #endregion

        #region Inherited folder and file permissions
        try {
            Write-Verbose 'Inherited permissions'
            if ($Action -ne 'New') {
                $jobs = @()

                foreach ($folder in $foldersWithAcl) {
                    $InvokeParams = @{
                        ScriptBlock  = $inheritedPermissionsScriptBlock
                        ArgumentList = $folder.Path, $Action, @($folder.InheritedFolderAcl.Access), @($folder.InheritedFileAcl.Access), $ignoredFolderPaths, $tokenPrivileges, $DetailedLog
                    }
                    # $testArg = $InvokeParams.ArgumentList
                    # & $InvokeParams.ScriptBlock @testArg -Verbose

                    # <#
                    $jobs += Start-Job @InvokeParams

                    #region Wait for max running jobs
                    $waitParams = @{
                        Name       = $jobs | Where-Object { $_ }
                        MaxThreads = $JobThrottleLimit
                    }
                    Wait-MaxRunningJobsHC @waitParams
                    #endregion #>
                }

                $ErrorActionPreference = 'Continue'

                $jobResults = $jobs | Wait-Job | Receive-Job

                #region Combine results of jobs into one object
                foreach ($jobResult in $jobResults) {
                    foreach ($j in $jobResult.testedInheritedFilesAndFolders) {
                        foreach ($i in $j.GetEnumerator()) {
                            $testedInheritedFilesAndFolders[$i.Key] = $i.Value
                        }
                    }
                    foreach ($j in $jobResult.IncorrectInheritedAcl) {
                        if ($DetailedLog) {
                            foreach ($i in $j.GetEnumerator()) {
                                $IncorrectInheritedAcl[$i.Key] = $i.Value
                            }
                        }
                        else {
                            $IncorrectInheritedAcl.Add($j)
                        }
                    }
                }
                #endregion

                if ($IncorrectInheritedAcl.Count -ne 0) {
                    [PSCustomObject]@{
                        Type        = 'Warning'
                        Name        = 'Inherited permissions incorrect'
                        Description = "All folders that don't have permissions assigned to them in the worksheet 'Permissions' are supposed to inherit their permissions from the parent folder. Files can only inherit permissions from the parent folder and are not allowed to have explicit permissions."
                        Value       = if ($DetailedLog) {
                            $IncorrectInheritedAcl
                        }
                        else {
                            $IncorrectInheritedAcl.ToArray()
                        }
                    }
                }
            }
        }
        catch {
            throw "Failed checking/setting the inheritance on folders and files: $_"
        }
        #endregion
    }
    catch {
        throw "Failed setting the permissions: $_"
    }
}