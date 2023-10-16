#Requires -Version 5.1
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
Param (
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

Begin {
    Function New-AceHC {
        <#
        .SYNOPSIS
            Convert an AD Object name and a permission character to a valid ACE.

        .DESCRIPTION
            Convert an AD Object name and a permission character to a valid Access Control List Entry.

        .PARAMETER Type
            The permission character defining the access to the folder.

        .PARAMETER Name
            Name of the AD object, used to identify the user or group within AD.
        #>

        [CmdLetBinding()]
        Param (
            [Parameter(Mandatory)]
            [ValidateSet('L', 'R', 'W', 'F', 'M')]
            [String]$Access,
            [Parameter(Mandatory)]
            [String]$Name,
            [Parameter(Mandatory)]
            [ValidateSet('Folder', 'InheritedFile', 'InheritedFolder')]
            [String]$Type
        )

        #Write-Verbose "Create ACE name '$Name' type '$Type'"

        Switch ($Access) {
            'L' {
                if (($type -eq 'Folder') -or ($type -eq 'InheritedFolder')) {
                    New-Object System.Security.AccessControl.FileSystemAccessRule(
                        "$env:USERDOMAIN\$Name",
                        [System.Security.AccessControl.FileSystemRights]::ReadAndExecute,
                        [System.Security.AccessControl.InheritanceFlags]::ContainerInherit,
                        [System.Security.AccessControl.PropagationFlags]::None,
                        [System.Security.AccessControl.AccessControlType]::Allow
                    )    
                }

                Break
            }
            'W' {
                if ($type -eq 'Folder') {
                    # This folder only
                    New-Object System.Security.AccessControl.FileSystemAccessRule(
                        "$env:USERDOMAIN\$Name",
                        [System.Security.AccessControl.FileSystemRights]'CreateFiles, AppendData, DeleteSubdirectoriesAndFiles, ReadAndExecute, Synchronize',
                        [System.Security.AccessControl.InheritanceFlags]::None,
                        [System.Security.AccessControl.PropagationFlags]::InheritOnly,
                        [System.Security.AccessControl.AccessControlType]::Allow
                    )
                    # Subfolders and files only
                    New-Object System.Security.AccessControl.FileSystemAccessRule(
                        "$env:USERDOMAIN\$Name",
                        [System.Security.AccessControl.FileSystemRights]'DeleteSubdirectoriesAndFiles, Modify, Synchronize',
                        [System.Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit',
                        [System.Security.AccessControl.PropagationFlags]::InheritOnly,
                        [System.Security.AccessControl.AccessControlType]::Allow
                    )
                }
                elseif ($type -eq 'InheritedFolder') {
                    # Subfolders and files only
                    New-Object System.Security.AccessControl.FileSystemAccessRule(
                        "$env:USERDOMAIN\$Name",
                        [System.Security.AccessControl.FileSystemRights]'DeleteSubdirectoriesAndFiles, Modify, Synchronize',
                        [System.Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit',
                        [System.Security.AccessControl.PropagationFlags]::InheritOnly,
                        [System.Security.AccessControl.AccessControlType]::Allow
                    )
                }
                elseif ($Type -eq 'InheritedFile') {
                    New-Object System.Security.AccessControl.FileSystemAccessRule(
                        "$env:USERDOMAIN\$Name",
                        [System.Security.AccessControl.FileSystemRights]'DeleteSubdirectoriesAndFiles, Modify, Synchronize',
                        [System.Security.AccessControl.InheritanceFlags]::None,
                        [System.Security.AccessControl.PropagationFlags]::None,
                        [System.Security.AccessControl.AccessControlType]::Allow
                    )
                }
                
                Break
            }
            'R' {
                if (($type -eq 'Folder') -or ($type -eq 'InheritedFolder')) {
                    New-Object System.Security.AccessControl.FileSystemAccessRule(
                        "$env:USERDOMAIN\$Name",
                        [System.Security.AccessControl.FileSystemRights]::ReadAndExecute,
                        [System.Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit',
                        [System.Security.AccessControl.PropagationFlags]::None,
                        [System.Security.AccessControl.AccessControlType]::Allow
                    )
                }
                elseif ($Type -eq 'InheritedFile') {
                    New-Object System.Security.AccessControl.FileSystemAccessRule(
                        "$env:USERDOMAIN\$Name",
                        [System.Security.AccessControl.FileSystemRights]::ReadAndExecute,
                        [System.Security.AccessControl.InheritanceFlags]::None,
                        [System.Security.AccessControl.PropagationFlags]::None,
                        [System.Security.AccessControl.AccessControlType]::Allow
                    )
                }
                Break
            }
            'F' {
                if (($type -eq 'Folder') -or ($type -eq 'InheritedFolder')) {
                    New-Object System.Security.AccessControl.FileSystemAccessRule(
                        "$env:USERDOMAIN\$Name",
                        [System.Security.AccessControl.FileSystemRights]::FullControl,
                        [System.Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit',
                        [System.Security.AccessControl.PropagationFlags]::None,
                        [System.Security.AccessControl.AccessControlType]::Allow
                    )
                }
                elseif ($Type -eq 'InheritedFile') {
                    New-Object System.Security.AccessControl.FileSystemAccessRule(
                        "$env:USERDOMAIN\$Name",
                        [System.Security.AccessControl.FileSystemRights]::FullControl,
                        [System.Security.AccessControl.InheritanceFlags]::None,
                        [System.Security.AccessControl.PropagationFlags]::None,
                        [System.Security.AccessControl.AccessControlType]::Allow
                    )
                }
                Break
            }
            'M' {
                if (($type -eq 'Folder') -or ($type -eq 'InheritedFolder')) {
                    New-Object System.Security.AccessControl.FileSystemAccessRule(
                        "$env:USERDOMAIN\$Name",
                        [System.Security.AccessControl.FileSystemRights]::Modify,
                        [System.Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit',
                        [System.Security.AccessControl.PropagationFlags]::None,
                        [System.Security.AccessControl.AccessControlType]::Allow
                    )
                }
                elseif ($Type -eq 'InheritedFile') {
                    New-Object System.Security.AccessControl.FileSystemAccessRule(
                        "$env:USERDOMAIN\$Name",
                        [System.Security.AccessControl.FileSystemRights]::Modify,
                        [System.Security.AccessControl.InheritanceFlags]::None,
                        [System.Security.AccessControl.PropagationFlags]::None,
                        [System.Security.AccessControl.AccessControlType]::Allow
                    )
                }
                Break
            }
            Default {
                throw "Permission character '$_' not supported."
            }
        }
    }

    Function Test-AclEqualHC {
        <#
	    .SYNOPSIS
		    Compare two Access Control Entries.

	    .DESCRIPTION
		    Checks if two ACE's are matching. Returns True if both ACE lists are equal and
            False when they don't.

        .PARAMETER ReferenceAce
            Reference collection of Access Control Entries of the first list

        .PARAMETER DifferenceAce
            Difference collection of Access Control Entries of the second list
#>

        [OutputType([Boolean])]
        Param (
            [Parameter(Mandatory)]
            [System.Object[]]$ReferenceAce,
            [System.Object[]]$DifferenceAce
        )

        Try {
            $aclMatchCount = 0

            foreach ($D in $DifferenceAce) {
                $aclMatch = $ReferenceAce.Where( 
                    {
                        ($D.FileSystemRights -eq $_.FileSystemRights) -and
                        ($D.AccessControlType -eq $_.AccessControlType) -and
                        ($D.IdentityReference -eq $_.IdentityReference) -and
                        ($D.InheritanceFlags -eq $_.InheritanceFlags) #-and
                        # ($D.PropagationFlags -eq $_.PropagationFlags) # <<<< issue
                    }, 'First'
                )

                if ($aclMatch) {
                    $aclMatchCount++
                }
                else {
                    # Write-Verbose "ACL equal 'false'"
                    Return $False
                }
            }

            if ($aclMatchCount -ne $ReferenceAce.Count) {
                # Write-Verbose "ACL equal 'false'"
                Return $False
            }

            # Write-Verbose "ACL equal 'true'"
            Return $True
        }
        Catch {
            throw "Failed testing the ACL for equality: $_"
        }
    }

    Function Wait-MaxRunningJobsHC {
        <# 
        .SYNOPSIS   
            Limit how many jobs can run at the same time
    
        .DESCRIPTION
            Only allow a specific quantity of jobs to run at the same time.
            Also wait for launching new jobs when there is not enough free 
            memory.
    
        .PARAMETER Name
            Name of the variable holding the jobs returned by 'Start-Job' or
            'Invoke-Command -AsJob'.
    
        .PARAMETER MaxThreads
            The number of jobs that are allowed to run at the same time.
    
        .PARAMETER MaxAllowedCpuLoadPercentage
            The CPU load must be below this percentage before we exit the 
            function to start a new job.
    
        .EXAMPLE
            $jobs = @()
    
            $scriptBlock = {
                Write-Output 'do work'
                Start-Sleep -Seconds 30
            }
    
            foreach ($i in 1..20) {
                Write-Verbose "Start job $i"
                $jobs += Start-Job -ScriptBlock $ScriptBlock
                Wait-MaxRunningJobsHC -Name $jobs -MaxThreads 3
            }
    
            Only allow 3 jobs to run at the same time. Wait to launch the next
            job until one is finished.
        #>
        
        [CmdletBinding()]
        Param (
            [Parameter(Mandatory)]
            [System.Management.Automation.Job[]]$Name,
            [Parameter(Mandatory)]
            [Int]$MaxThreads,
            [Int]$MaxAllowedCpuLoadPercentage = 80
        )
    
        Begin {
            Function Get-CPUloadHC {
                (
                    Get-Counter '\Processor(_Total)\% Processor Time'
                ).CounterSamples.CookedValue
            }
            Function Get-RunningJobsHC {
                @($Name).Where( { $_.State -eq 'Running' })
            }
        }
    
        Process {
            while ((Get-CPUloadHC) -gt $MaxAllowedCpuLoadPercentage) {
                Start-Sleep -Milliseconds 500
            }
    
            while ((Get-RunningJobsHC).Count -ge $MaxThreads) {
                $null = Wait-Job -Job $Name -Any
            }
        }
    }

    $inheritedPermissionsScriptBlock = {
        [OutputType([PSCustomObject[]])]
        [CmdLetBinding()]
        Param (
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

        Function Test-AclEqualHC {
            <#
            .SYNOPSIS
                Compare two Access Control Entries.
    
            .DESCRIPTION
                Checks if two ACE's are matching. Returns True if both ACE lists are equal and
                False when they don't.
    
            .PARAMETER ReferenceAce
                Reference collection of Access Control Entries of the first list
    
            .PARAMETER DifferenceAce
                Difference collection of Access Control Entries of the second list
    #>
    
            [OutputType([Boolean])]
            Param (
                [Parameter(Mandatory)]
                [System.Object[]]$ReferenceAce,
                [System.Object[]]$DifferenceAce
            )
    
            Try {
                $aclMatchCount = 0
    
                foreach ($D in $DifferenceAce) {
                    $aclMatch = $ReferenceAce.Where( 
                        {
                            ($D.FileSystemRights -eq $_.FileSystemRights) -and
                            ($D.AccessControlType -eq $_.AccessControlType) -and
                            ($D.IdentityReference -eq $_.IdentityReference) -and
                            ($D.InheritanceFlags -eq $_.InheritanceFlags)
                        }, 'First'
                    )
    
                    if ($aclMatch) {
                        $aclMatchCount++
                    }
                    else {
                        # Write-Verbose "ACL equal 'false'"
                        Return $False
                    }
                }
    
                if ($aclMatchCount -ne $ReferenceAce.Count) {
                    # Write-Verbose "ACL equal 'false'"
                    Return $False
                }
    
                # Write-Verbose "ACL equal 'true'"
                Return $True
            }
            Catch {
                throw "Failed testing the ACL for equality: $_"
            }
        }
    
        Function Get-FolderContentHC {
            Param (
                [Parameter(Mandatory)]
                [String]$Path
            )
    
            Try {
                $childItems = (Get-ChildItem -LiteralPath $Path -EA Stop).Where( 
                    { -not ($IgnoredFolderPaths.ContainsKey($_.FullName)) }
                )
            }
            Catch {
                throw "Failed retrieving the folder content of '$Path': $_"
            }
    
            foreach ($child in $childItems) {
                Try {
                    $acl = $child.GetAccessControl()
                }
                Catch {
                    if (-not (Test-Path -LiteralPath $child.FullName)) {
                        Write-Verbose "Item '$($child.FullName)' removed"
                        $Error.RemoveAt(0)
                    }
                    else {
                        $ErrorActionPreference = 'Continue'

                        Write-Error "Failed retrieving the ACL of '$($child.FullName)': $_"
                        
                        $Error.RemoveAt(1)
                        $ErrorActionPreference = 'Stop'
                    }
                    Continue
                }
    
                if (-not $child.PSIsContainer) {
                    # Only for Pester testing:
                    $testedInheritedFilesAndFolders[$child.FullName] = $true
    
                    if (
                        -not (Test-AclEqualHC -ReferenceAce $FileAclAccessList -DifferenceAce $acl.Access)
                    ) {
                        & $incorrectAclInheritedOnly
                    }
                }
                else {
                    # Only for Pester testing:
                    $testedInheritedFilesAndFolders[$child.FullName] = $true
    
                    if (
                        -not (Test-AclEqualHC -ReferenceAce $FolderAclAccessList -DifferenceAce $acl.Access)
                    ) {
                        & $incorrectAclInheritedOnly
                    }
    
                    Get-FolderContentHC -Path $child.FullName
                }
            }
        }
    
        $incorrectAclInheritedOnly = {
            Write-Warning "Incorrect ACL '$($child.FullName)'"
            #region Log
            if ($DetailedLog) {
                $incorrectInheritedAcl.($child.FullName.TrimStart('\\?\')) = $acl.AccessToString
            }
            else {
                $incorrectInheritedAcl.Add($child.FullName.TrimStart('\\?\'))
            }
            #endregion
    
            #region Set permissions
            if ($Action -eq 'Fix') {
                Write-Verbose "Set ACL to inherited only '$($child.FullName)'"
    
                if ($child.PSIsContainer) {
                    # This is a workaround for non inherited permissions
                    # that do not get properly removed
                    $acl.Access | ForEach-Object {
                        $acl.RemoveAccessRuleSpecific($_)
                    }
                    $child.SetAccessControl($acl)
                    # for one reason or another the below does not work repetitively
                    # so we use Set-Acl instead
                    # $child.SetAccessControl($inheritedDirAcl)
                    Set-Acl -Path $child.FullName -AclObject $inheritedDirAcl
                }
                else {
                    $acl.Access | ForEach-Object {
                        $acl.RemoveAccessRuleSpecific($_)
                    }
                    $child.SetAccessControl($acl)
                    # for one reason or another the below does not work repetitively
                    # so we use Set-Acl instead
                    # $child.SetAccessControl($inheritedFileAcl)
                    Set-Acl -Path $child.FullName -AclObject $inheritedFileAcl
                }
            }
            #endregion
        }
    
        Try {
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
            Try {
                Write-Verbose 'Get super powers'
                Add-Type $TokenPrivileges
                [void][TokenManipulator]::AddPrivilege('SeRestorePrivilege')
                [void][TokenManipulator]::AddPrivilege('SeBackupPrivilege')
                [void][TokenManipulator]::AddPrivilege('SeTakeOwnershipPrivilege')
            }
            Catch {
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
            Try {
                Get-FolderContentHC -Path $Path
            }
            Catch {
                throw "Failed checking or setting the inheritance in folder '$Path': $_"
            }
            #endregion
        }
        Catch {
            throw "Failed setting permissions for '$Path': $_"
        }
        Finally {
            [PSCustomObject]@{
                testedInheritedFilesAndFolders = $testedInheritedFilesAndFolders
                IncorrectInheritedAcl          = $incorrectInheritedAcl
            }
        }
    }

    $tokenPrivileges = @"
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
"@
}

Process {
    Try {
        $ErrorActionPreference = 'Stop'

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
        Try {
            Write-Verbose 'Get super powers'
            Add-Type $tokenPrivileges
            [void][TokenManipulator]::AddPrivilege('SeRestorePrivilege')
            [void][TokenManipulator]::AddPrivilege('SeBackupPrivilege')
            [void][TokenManipulator]::AddPrivilege('SeTakeOwnershipPrivilege')
        }
        Catch {
            throw "Failed getting super powers: $_"
        }
        #endregion

        #region Create the parent folder when action is New
        Try {
            $missingFolders = [System.Collections.Generic.List[String]]::New()

            if ($Action -eq 'New') {
                Try {
                    $missingFolders.Add((New-Item -Path $Path -ItemType Directory -EA Stop).FullName)
                }
                Catch {
                    $Error.RemoveAt(0)
                    Return [PSCustomObject]@{
                        Type        = 'FatalError'
                        Name        = 'Parent folder exists already'
                        Description = "The folder defined as 'Path' in the worksheet 'Settings' cannot be present on the remote machine when 'Action=New' is used. Please use 'Action' with value 'Check' or 'Fix' instead."
                        Value       = $Path
                    }
                }
            }
            elseif (-not (Test-Path -LiteralPath $Path -PathType Container)) {
                Return [PSCustomObject]@{
                    Type        = 'FatalError'
                    Name        = 'Parent folder missing'
                    Description = "The folder defined as 'Path' in the worksheet 'Settings' needs to be available on the remote machine. In case the folder structure needs to be created, please use 'Action=New' instead."
                    Value       = $Path
                }
            }
            
            Write-Verbose "Parent folder '$Path'"
            
            # Set-Location -Path $Path
        }
        Catch {
            throw "Failed checking the existence of the parent folder: $_"
        }
        #endregion
        
        #region Add the FullName for each path
        foreach ($M in $Matrix) {
            $tmpPath = if ($M.Parent) { $Path }
            else { Join-Path -Path $Path -ChildPath $M.Path }

            $M.Path = "\\?\$tmpPath"
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
                Value       = $IgnoredFolders.Path.TrimStart('\\?\')
            }
        }
        #endregion

        #region Inaccessible files
        $FoldersListOnlyAclRegex = $Matrix.Where( {
                (-not ($_.Acl.Values.Where( { $_ -ne 'L' }))) -and ($_.ACL.Count -ne 0)
            }).ForEach( {
                [Regex]::Escape("\\?\$_")
            }) -join '|'

        $FoldersWithPermissionsRegex = $Matrix.Where( {
                ($_.Acl.Values.Where( { $_ -ne 'L' }))
            }).ForEach( {
                [Regex]::Escape("\\?\$_")
            }) -join '|'
        #endregion

        #region Create file and folder ACL for each path in the matrix
        Try {
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
                        Try {
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
                        Catch {
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
        Catch {
            throw "Failed creating the AccessControlList: $_"
        }
        #endregion

        #region Missing folders
        Try {
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
                    Value       = $missingFolders.ToArray().TrimStart('\\?\')
                }

                Switch ($Action) {
                    'New' {
                        $Obj.Name = 'Child folder created'
                        $Obj.Description = "All folders defined in the worksheet 'Permissions' have been created with the correct permissions underneath the parent folder defined in the worksheet 'Settings'."
                        Break
                    }
                    'Fix' {
                        $Obj.Name = 'Child folder created'
                        $Obj.Description = "The missing folders underneath the parent folder have been created."
                        Break
                    }
                    'Check' {
                        $Obj.Name = 'Child folder missing'
                        $Obj.Description = "Not all folders defined in the worksheet 'Permissions' were found underneath the parent folder."
                        Break
                    }
                    Default {
                        throw "Action '$_' is not supported."
                    }
                }

                $Obj
            }
            else {
                Write-Verbose 'All folders present, no missing folders'
            }
        }
        Catch {
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
                $folderItem = Get-Item -Path $folder.Path -EA Stop

                # Only for Pester testing:
                $testedNonInheritedFolders[$folder.Path] = $folder

                $acl = $folderItem.GetAccessControl()

                $testEqualParams = @{
                    ReferenceAce  = ($folder.FolderAcl).Access
                    DifferenceAce = ($acl).Access
                }

                if (
                    (-not $acl.AreAccessRulesProtected) -or
                    (-not (Test-AclEqualHC @testEqualParams))
                ) {
                    Write-Warning "Incorrect folder ACL '$($folder.Path)'"
                    #region Log
                    if ($Action -ne 'New') {
                        if ($DetailedLog) {
                            $incorrectAclNonInheritedFolders.($folder.Path.TrimStart('\\?\')) = @{
                                'Old' = $acl.AccessToString
                                'New' = ($folder.FolderAcl).AccessToString
                            }
                        }
                        else {
                            $incorrectAclNonInheritedFolders.Add($folder.Path.TrimStart('\\?\'))
                        }
                    }
                    #endregion

                    #region Set permissions
                    if ($Action -ne 'Check') {
                        Write-Verbose 'Set correct ACL'
                        
                        # workaround for non inherited permissions
                        # that do not get properly removed
                        $acl.Access | ForEach-Object {
                            $acl.RemoveAccessRuleSpecific($_)
                        }
                        $folderItem.SetAccessControl($acl)
                        $folderItem.SetAccessControl($folder.FolderAcl)

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
        Try {
            Write-Verbose 'Inherited permissions'
            if ($Action -ne 'New') {
                $jobs = @()

                foreach ($folder in $foldersWithAcl) {
                    $InvokeParams = @{
                        ScriptBlock  = $inheritedPermissionsScriptBlock
                        ArgumentList = $folder.Path, $Action, @($folder.InheritedFolderAcl.Access), @($folder.InheritedFileAcl.Access), $ignoredFolderPaths, $tokenPrivileges, $DetailedLog
                    }
                    $jobs += Start-Job @InvokeParams

                    #region Wait for max running jobs
                    $waitParams = @{
                        Name       = $jobs | Where-Object { $_ }
                        MaxThreads = $JobThrottleLimit
                    }
                    Wait-MaxRunningJobsHC @waitParams
                    #endregion
                }

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
        Catch {
            throw "Failed checking/setting the inheritance on folders and files: $_"
        }
        #endregion
    }
    Catch {
        throw "Failed setting the permissions: $_"
    }
}