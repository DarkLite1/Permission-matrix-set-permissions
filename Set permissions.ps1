#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
    .SYNOPSIS
        Create, check or fix a folder structure based on folder names and
        folder permission characters in an Excel file.

    .DESCRIPTION
        Check/Fix:
        We create a new folder structure in the temporary cache. After the folders have been created we
        will apply the default permissions first and then the permissions defined in the matrix. So the
        permissions defined in the matrix will always win over the default permissions.
        Then we will compare the permissions on the destination with the ones we applied in the temporary
        cache. All permission issues and other problems will be reported by returning objects. In case of 'Fix'
        we also fix these problems.

        New:
        The new folder structure is immediately created on the destination.

    .PARAMETER Path
        The parent folder on the localhost where the folder tree starts.

    .PARAMETER Action
        Accepted values are:
        'New'    > Creates a new folder structure on the destination with the correct permissions.
        'Check'  > Checks if an existing folder structure has the correct permissions.
        'Fix'    > Checks and fixes an existing folder structure with the correct permissions.

    .PARAMETER Matrix
        The array containing the correct folder names and their permissions.

    .PARAMETER DetailedLog
        When set to true, the script will be able to log more details for better troubleshooting.
        In case of ACL's that are incorrect, in normal circumstances only the FullName of the path
        is reported. When DetailedLog is enabled, the ACL's that are on the folder and the expected
        ACL are also reported.

        Keep in mind that when enabling DetailedLog, the script execution performance will drop. For
        this reason it is only advised to use this feature only in case of troubleshooting.

    .NOTES
        CHANGELOG
        2014/08/25 Script born
        2014/10/27 Added checks for existing folder structures
        2014/10/28 Added fixing of incorrect ACL's on the target folders
        2014/10/30 Added switch 'Action' for 'New', 'Check' and 'Fix'
        2014/11/26 Code clean-up
        2014/11/28 Added reading in first line as target path instead of Excel field A1
        2014/12/02 Added remoting capability, so we can run on another server
        2014/12/03 Updated functions to use 'LiteralPath' instead of 'Path'
        2014/12/05 Improved script logic & execution order
        2014/12/05 Made '$ComputerName' a variable in the input file
        2014/12/11 Added '$Mail' to allow sending e-mails when needed
        2015/01/26 Added 'New-LogFileNameHC'
        2015/02/09 Added 'IgnoredFolders' to not treat everything we find in the Excel sheet
        2015/02/09 Added support for users in the Excel sheet
        2015/03/13 Updated e-mail content when error found in the input file
        2015/03/18 Added preflight checks 'online, remoting enabled, PS Version up-to-date'
        2015/03/18 Changed 'MailTo' as input from the Excel sheet
        2015/03/18 Removed BNL Naming Convention check, for use with other countries
        2015/03/19 Updated descriptions to be more clear for end-users
        2015/03/23 Added check for OS version if remoting fails
        2015/04/16 Switched UserName and Password for a PS Credential object in the parameters '$Credential'
        2015/04/29 Rewritten the whole thing to output objects and remove all the different 'Invoke-Commands'
        2015/10/19 Moved to new script server
        2015/12/03 Added check for incorrect paths in the worksheet 'Settings' that end with '\' $PathIncorrect
        2015/12/04 Improved 'Test-Input' to update the object immediately instead of creating extra variables
        2015/12/08 Improved searching for files
        2016/01/21 Improved error checking for the line 'Path' in the worksheet 'Permissions'
                    The option 'ignore' ('i') is not allowed on this line
        2016/01/21 Fixed a bug where we check file inheritance in case the target folder is empty
        2016/02/05 Improved file ACL check to make sure we catch the files that Windows reports being 'Inherited'
                    but actually aren't (SR-859049)
        2016/02/05 Improved speed by removing 'Add-Member' in 'Compare-PermissionsHC'
        2016/02/09 Fixed 'PathTooLong' in case we can't create a file in the folder anymore
        2016/02/10 Added parameter 'NrFilesToReport' to avoid huge lists of files that have incorrect permissions
        2016/02/10 Simplified 'Test-AclInheritedOnlyHC'
        2016/02/15 Rewrote 'Compare-PermissionsHC' completely
        2016/02/18 Improved code with break for switch statements
        2016/02/18 Added '$NoPermissions'
        2016/02/19 Added '$DeadFiles', files that are in folders with list only permissions, so no one 
                   has access to them
        2016/02/26 Added 'Get-SharePermissionsHC' to check the permissions on share level, not only NTFS
        2016/03/08 Added 'GroupName' parameter for HTML log file
        2016/03/08 Improved 'Get-SharePermissionsHC' to collect all share permissions that are not 
                   'Everyone Full control'
        2016/03/29 Improved error handling messages and statuses, added 'Unknown error' this allows us to have
                    the correct errors in the site specific log file instead of in the body of the general e-mail.
        2016/03/30 Fixed 'Test-AclEqualHC' to use 'AccessToString' instead of 'Access'
        2016/04/11 Improved 'Add-AclHC', 'Set-AclEqualHC', 'Set-AclOwnerHC' to use '-LiteralPath' instead of '-Path'
        2016/04/11 Improved 'Test-AclEqualHC'
        2016/06/17 Rewrote the whole script to use Workflows for load balancing
        2016/07/22 Improved 'Test-AclEqualHC' to not check for folder existence in the parameters
        2016/08/01 Fixed a race condition with the creation of the temp folder '$TempSourceFolder' for each matrix.
                   We now use a unique code instead of a combination of 'SiteCode' and 'Path' ending.
        2016/09/23 Added Windows Event Logging
        2016/09/29 Added 'ShareABECorrected' to avoid users seeing files and folders where they don't have access on
                   We enable the option Access Based Enumeration on Shares to avoid this
        2017/04/10 Enhanced Windows Event Logging
        2017/05/19 Fixed 'SrcFolder' by ireplace which is case insensitive
        2017/06/27 Improved Test-AclEqualHC to be faster by using SDDL
        2017/07/03 Changed Test-AclEqualHC again because files were reported incorrect which was not the case
        2017/07/24 Changed Test-AclEqualHC again, because incorrect permissions were reported 
                   when this wasn't the case
        2017/09/12 Changed Test-AclEqualHC again, in case of duplicate admin accounts we don't report this as incorrect
        2018/08/07 Changed parameters to Path, Action, Matrix
                   Rewrote functions for speed, simplicity and testability with Pester
                   Redesigned script to run as a job for a single matrix
        2018/08/10 Added DetailedLog
                   Removed unused functions
        2019/05/07 Added more detailed error handing in Get-FolderContentHC
        2020/01/23 Fixed an issue where old ACE's were not removed from the ACL for folders
                   Speed-up the applying of an ACL for files and folders
                   Added verbose messages
        2020/01/28 Moved object creation on inherited only ACL's outside the loop for better speed
                   Converted function Test-AclIsInheritedOnlyHC to code for speed
                   Added a workaround for a bug in .NET where the non inherited permissions are not being removed
                   https://social.microsoft.com/Forums/en-US/5770f0bd-fddd-442b-b917-daf88ff28b10/removing-modifysync-ntfs-permissions-using-removeaccessruleall-in-powershell?forum=Offtopic&prof=required
        2020/02/28 Fixed ignored folders not always handled correctly
        2020.02.03 Fixed a bug where the AD object name was not returned in the error message when the ACL creation failed
                   Inherited folders are now filtered with a '.where' clause instead of an an 'if'
        2020/03/05 Added better Pester tests
                   Changed the way that inherited folders and files are checked:
                   Previously we only checked if AccessRulesAreProtected was true, this was not sufficient
                   now we check each entry in the ACL to make sure that the folder/file is inheriting the correct permissions
        2020/08/06 Replaced SetAccessControl with Set-Acl as repetitive changes for correcting intherited folders were not correctly applied

        AUTHOR Brecht.Gijbels@heidelbergcement.com #>

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
    [Boolean]$DetailedLog
)

Begin {
    $scannedInheritedFolders = @{ }
    $testedInheritedFilesAndFolders = @{ }

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

        .NOTES
	        CHANGELOG
	        2018/08/07 Function born
            2019/03/22 Add verbose

	        AUTHOR Brecht.Gijbels@heidelbergcement.com #>

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

    Function Get-FolderContentHC {
        Param (
            [Parameter(Mandatory)]
            [String]$Path,
            [Parameter(Mandatory)]
            $FolderAcl,
            $FileAcl
        )

        # write-Verbose "Get folder content '$Path'"

        if ($scannedInheritedFolders.ContainsKey($Path)) { Return }
        $scannedInheritedFolders[$Path] = $true

        Try {
            $Members = (Get-ChildItem -LiteralPath $Path -EA Stop).Where( {
                    -not ($ignoredFolders.Contains($_.FullName))
                })
        }
        Catch {
            throw "Failed retrieving the folder content of '$Path': $_"
        }

        foreach ($M in $Members) {
            if ($NonInheritedFolders -notContains $M.FullName) {
                # Only for Pester testing:
                $testedInheritedFilesAndFolders[$M.FullName] = $true

                # Write-Verbose "Test inheritance only '$($M.FullName)'"

                Try {
                    $Acl = $M.GetAccessControl()
                }
                Catch {
                    throw "Failed retrieving the ACL of '$($M.FullName)': $_"
                }

                if ($Acl.AreAccessRulesProtected) {
                    & $IncorrectAclInheritedOnly
                }
                elseif (
                    (-not ($M.PSIsContainer)) -and 
                    (-not (Test-AclEqualHC -ReferenceAce $Acl -DifferenceAce $FileAcl))) {
                        & $IncorrectAclInheritedOnly
                }
                elseif (-not (Test-AclEqualHC -ReferenceAce $Acl -DifferenceAce $FolderAcl)) {
                    & $IncorrectAclInheritedOnly
                }
            }

            if ($M.PSIsContainer) {
                Get-FolderContentHC -Path $M.FullName -FolderAcl $FolderAcl -FileAcl $FileAcl
            }
        }

        <# Fix when $Acl = $M.GetAccessControl() fails:
        
        $error.Clear()

        #$File = 'C:\Users\SrvBatch\Downloads\Text.xml'
        #$Script = 'C:\Users\SrvBatch\Downloads\Permission matrix\Set permissions.ps1'
        #$Params = Import-CliXml -Path $File
        #& $Script -Path 'E:\DEPARTMENTS\RMC\IB\04-SITE\01-North\Genk' -Action 'Fix' -Matrix $Params.ArgumentList[2] -DetailedLog $true


        # Error file where 'GetAccessControl' does not work
        $File = '\\?\E:\DEPARTMENTS\RMC\IB\04-SITE\01-North\Genk\08-Technology\S-Genk-T-04-Grondst-MatPrem\03-Zevingen-Granulo\3738CC08.tmp'

        $FileItem = Get-Item -LiteralPath $File
        $FileItem.GetAccessControl()

        # Take ownership
        $user = $env:username
        $Account = New-Object System.Security.Principal.NTAccount($user)
        $FileSecurity = new-object System.Security.AccessControl.FileSecurity
        $FileSecurity.SetOwner($Account)
        [System.IO.File]::SetAccessControl($file, $FileSecurity)

        # Problem fixed
        $FileItem = Get-Item -LiteralPath $File
        $FileItem.GetAccessControl()
                #>
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

        .NOTES
            2018/08/06 Function born
            2018/08/09 Rewritten to only compare ACE's and not ACL's
            2019/01/22 Renamed to Test-AclEqualHC
                       Add verbose

            AUTHOR Brecht.Gijbels@heidelbergcement.com
	    #>

        [OutputType([Boolean])]
        Param (
            [Parameter(Mandatory)]
            [System.Object[]]$ReferenceAce,
            [System.Object[]]$DifferenceAce
        )

        Try {
            $Matches = @()

            foreach ($D in $DifferenceAce) {
                $Match = @($ReferenceAce).Where( {
                        ($D.FileSystemRights -eq $_.FileSystemRights) -and
                        ($D.AccessControlType -eq $_.AccessControlType) -and
                        ($D.IdentityReference -eq $_.IdentityReference) -and
                        ($D.InheritanceFlags -eq $_.InheritanceFlags) -and
                        ($D.PropagationFlags -eq $_.PropagationFlags)
                    })

                if ($Match) {
                    $Matches += $Match
                }
                else {
                    # Write-Verbose "ACL equal 'false'"
                    Return $False
                }
            }

            if ($Matches.Count -ne $ReferenceAce.Count) {
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

    $IncorrectAclInheritedOnly = {
        Write-Warning "Incorrect ACL '$($M.FullName)'"
        #region Log
        if ($DetailedLog) {
            $IncorrectAclInheritedFolders.($M.FullName.TrimStart('\\?\')) = $Acl.AccessToString
        }
        else {
            $IncorrectAclInheritedFolders.Add($M.FullName.TrimStart('\\?\'))
        }
        #endregion

        #region Set permissions
        if ($Action -eq 'Fix') {
            Write-Verbose 'Set ACL to inherited only'

            if ($M.PSIsContainer) {
                # This is a workaround for non inherited permissions
                # that do not get properly removed
                $Acl.Access | ForEach-Object {
                    $Acl.RemoveAccessRuleSpecific($_)
                }
                $M.SetAccessControl($Acl)
                # for one reason or another the below does not work repetitively
                # so we use Set-Acl instead
                # $M.SetAccessControl($InheritedDirAcl)
                Set-Acl -Path $M.FullName -AclObject $InheritedDirAcl
            }
            else {
                $Acl.Access | ForEach-Object {
                    $Acl.RemoveAccessRuleSpecific($_)
                }
                $M.SetAccessControl($Acl)
                # for one reason or another the below does not work repetitively
                # so we use Set-Acl instead
                # $M.SetAccessControl($InheritedFileAcl)
                Set-Acl -Path $M.FullName -AclObject $inheritedFileAcl
            }
        }
        #endregion
    }

    $AdjustTokenPrivileges = @"
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
        # $WarningPreference = 'Continue'

        #region Variables
        Write-Verbose 'Set variables'
        $BuiltinAdmin = [System.Security.Principal.NTAccount]'Builtin\Administrators'
        $MissingFolders = [System.Collections.Generic.List[String]]::New()
        $InaccessibleData = [System.Collections.Generic.List[String]]::New()

        if ($DetailedLog) {
            $incorrectAclNonInheritedFolders = @{ }
            $incorrectAclInheritedFolders = @{ }
        }
        else {
            $incorrectAclNonInheritedFolders = [System.Collections.Generic.List[String]]::New()
            $incorrectAclInheritedFolders = [System.Collections.Generic.List[String]]::New()
        }
        #endregion

        #region Get super powers
        Try {
            Write-Verbose 'Get super powers'
            Add-Type $AdjustTokenPrivileges
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
            if ($Action -eq 'New') {
                Try {
                    $missingFolders.Add((New-Item -Path $Path -ItemType Directory -EA Stop).FullName)
                }
                Catch {
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
        $IgnoredFolders, $Matrix = $Matrix.Where( { $_.Ignore }, 'Split')

        if ($IgnoredFolders) {
            $IgnoredFolders = @($IgnoredFolders.Path)
            $IgnoredFolders.ForEach( { Write-Verbose "Ignored folder '$_'" })
            
            [PSCustomObject]@{
                Type        = 'Information'
                Name        = 'Ignored folder'
                Description = "All rows in the worksheet 'Permissions' that have the character 'i' defined are ignored. These folders are not checked for incorrect permissions."
                Value       = $IgnoredFolders.TrimStart('\\?\')
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
            $AdminFullControlFolderAce = New-Object System.Security.AccessControl.FileSystemAccessRule(
                [System.Security.Principal.NTAccount]'BUILTIN\Administrators',
                [System.Security.AccessControl.FileSystemRights]::FullControl,
                [System.Security.AccessControl.InheritanceFlags]'ContainerInherit,ObjectInherit',
                [System.Security.AccessControl.PropagationFlags]::None,
                [System.Security.AccessControl.AccessControlType]::Allow
            )
            $AdminFullControlIFileAce = New-Object System.Security.AccessControl.FileSystemAccessRule(
                [System.Security.Principal.NTAccount]'BUILTIN\Administrators',
                [System.Security.AccessControl.FileSystemRights]::FullControl,
                [System.Security.AccessControl.AccessControlType]::Allow
            )

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
                
                $folderAcl = New-Object System.Security.AccessControl.DirectorySecurity
                $folderAcl.SetAccessRuleProtection($true, $false)
                $folderAcl.SetOwner($BuiltinAdmin)

                $inheritedFolderAcl = New-Object System.Security.AccessControl.DirectorySecurity
                $inheritedFolderAcl.SetAccessRuleProtection($false, $false)
                $inheritedFolderAcl.SetOwner($BuiltinAdmin)

                $inheritedFileAcl = New-Object System.Security.AccessControl.FileSecurity
                $inheritedFileAcl.SetAccessRuleProtection($false, $false)
                $inheritedFileAcl.SetOwner($BuiltinAdmin)

                $M.ACL.GetEnumerator().Foreach( {
                        Try {
                            $ID = "$($_.Key)@$($_.Value)"

                            if (-not $aceCache.ContainsKey($ID)) {
                                $aceCache[$ID] = @{
                                    Folder          = @(New-AceHC -Access $_.Value -Name $_.Key -Type 'Folder')
                                    InheritedFolder = @(New-AceHC -Access $_.Value -Name $_.Key -Type 'InheritedFolder')
                                    InheritedFile   = @(New-AceHC -Access $_.Value -Name $_.Key -Type 'InheritedFile')
                                }
                            }

                            $aceCache[$ID]['Folder'].ForEach( { $folderAcl.AddAccessRule($_) })
                            $aceCache[$ID]['InheritedFolder'].ForEach( { $inheritedFolderAcl.AddAccessRule($_) })
                            $aceCache[$ID]['InheritedFile'].ForEach( { $inheritedFileAcl.AddAccessRule($_) })
                        }
                        Catch {
                            throw "AD object '$($ID.split('@')[0])' with permission character '$($ID.split('@')[1])' probably doesn't exist in AD: $_"
                        }
                    })

                $folderAcl.AddAccessRule($AdminFullControlFolderAce)
                $inheritedFolderAcl.AddAccessRule($AdminFullControlFolderAce)
                $inheritedFileAcl.AddAccessRule($AdminFullControlIFileAce)

                $M.FolderAcl = $folderAcl
                $M.inheritedFolderAcl = $inheritedFolderAcl
                $M.inheritedFileAcl = $inheritedFileAcl
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
        $testedNonInheritedFolders = @{ }

        Try {
            Write-Verbose 'Folders with ACL in the matrix that are not ignored'

            $FoldersWithAcl = $Matrix.Where( { ($_.FolderAcl) -and (-not $_.ignore) })

            foreach ( $F in $FoldersWithAcl) {
                Write-Verbose "Folder '$($F.Path)'"
                $FolderItem = Get-Item -Path $F.Path -EA Stop

                # Only for Pester testing:
                $testedNonInheritedFolders[$F.Path] = $true

                $Acl = $FolderItem.GetAccessControl()

                $ReferenceAce = ($F.FolderAcl).Access
                $DifferenceAce = ($Acl).Access

                if ((-not $Acl.AreAccessRulesProtected) -or
                    (-not (Test-AclEqualHC -ReferenceAce $ReferenceAce -DifferenceAce $DifferenceAce))) {
                    Write-Warning "Incorrect folder ACL '$($F.Path)'"
                    #region Log
                    if ($Action -ne 'New') {
                        if ($DetailedLog) {
                            $incorrectAclNonInheritedFolders.($F.Path.TrimStart('\\?\')) = @{
                                'Old' = $Acl.AccessToString
                                'New' = ($F.FolderAcl).AccessToString
                            }
                        }
                        else {
                            $incorrectAclNonInheritedFolders.Add($F.Path.TrimStart('\\?\'))
                        }
                    }
                    #endregion

                    #region Set permissions
                    if ($Action -ne 'Check') {
                        Write-Verbose 'Set correct ACL'
                        
                        # This is a workaround for non inherited permissions
                        # that do not get properly removed
                        $Acl.Access | ForEach-Object {
                            $Acl.RemoveAccessRuleSpecific($_)
                        }
                        $FolderItem.SetAccessControl($Acl)
                        $FolderItem.SetAccessControl($F.FolderAcl)

                        Write-Verbose 'ACL corrected'
                    }
                    #endregion
                }
            }

            <#     
            $NewAcl = New-Object System.Security.AccessControl.DirectorySecurity
            $NewAcl.SetOwner($BuiltinAdmin)
            $NewAcl.SetAccessRuleProtection($true,$false)
            $ReferenceAce.ForEach({$NewAcl.AddAccessRule($_)})
            $FolderItem.SetAccessControl($NewAcl)
            
            Write-Verbose 'Set SetAccessRuleProtection'
            $Acl.SetAccessRuleProtection($True, $False)
            $FolderItem.SetAccessControl($Acl)
            
            Write-Verbose 'Set owner'
            $Acl = $FolderItem.GetAccessControl()
            $Acl.SetOwner($BuiltinAdmin)
            $FolderItem.SetAccessControl($Acl)
            
            Write-Verbose 'Remove ACEs from ACL'
            $Acl = $FolderItem.GetAccessControl()
            $Acl.Access.ForEach({$null = $Acl.RemoveAccessRule($_)})
            
            Write-Verbose 'Add new ACL'
            $ReferenceAce.ForEach({$Acl.AddAccessRule($_)})
            
            Write-Verbose 'Set correct ACL'
            $FolderItem.SetAccessControl($Acl)
            #>

            if ($incorrectAclNonInheritedFolders.Count -ne 0) {
                [PSCustomObject]@{
                    Type        = 'Warning'
                    Name        = 'Non inherited folder incorrect permissions'
                    Description = "The folders that have permissions defined in the worksheet 'Permissions' are not matching with the permissions found on the folders of the remote machine."
                    Value       = if ($DetailedLog) { $incorrectAclNonInheritedFolders }
                    else { $incorrectAclNonInheritedFolders.ToArray() }
                }
            }
        }
        Catch {
            throw "Failed checking/setting the permissions on non inherited folders: $_"
        }
        #endregion

        #region Inherited folder and file permissions
        Try {
            Write-Verbose 'Inherited permissions'
            if ($Action -ne 'New') {
                $InheritedDirAcl = New-Object System.Security.AccessControl.DirectorySecurity
                $InheritedDirAcl.SetOwner($BuiltinAdmin)
                $InheritedDirAcl.SetAccessRuleProtection($false, $false)

                $InheritedFileAcl = New-Object System.Security.AccessControl.FileSecurity
                $InheritedFileAcl.SetOwner($BuiltinAdmin)
                $InheritedFileAcl.SetAccessRuleProtection($false, $false)

                $NonInheritedFolders = @($FoldersWithAcl.Path | Sort-Object)

                $FoldersWithAcl.ForEach( {
                        Get-FolderContentHC -Path $_.Path -FolderAcl $_.InheritedFolderAcl -FileAcl $_.InheritedFileAcl
                    })

                if ($incorrectAclInheritedFolders.Count -ne 0) {
                    [PSCustomObject]@{
                        Type        = 'Warning'
                        Name        = 'Inherited permissions incorrect'
                        Description = "All folders that don't have permissions assigned to them in the worksheet 'Permissions' are supposed to inherit their permissions from the parent folder. Files can only inherit permissions from the parent folder and are not allowed to have explicit permissions."
                        Value       = if ($DetailedLog) { $incorrectAclInheritedFolders }
                        else { $incorrectAclInheritedFolders.ToArray() }
                    }
                }

                if ($InaccessibleData.Count -ne 0) {
                    [PSCustomObject]@{
                        Type        = 'Warning'
                        Name        = 'Inaccessible data'
                        Description = "Files and folders that are found in folders where only list permissions are granted. When no one has read or write permissions, the files/folders become inaccessible."
                        Value       = $InaccessibleData.ToArray().TrimStart('\\?\')
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