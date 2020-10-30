<#
    Color
    ---------------------------------
    |    Error    |     Warning     |
    ---------------------------------

OVERVIEW
    FileName                 E  W  GoTo
    --------                 -  -  ----
    BEL MTX-CEM-CBR Antoing  0  0  Settings
    BEL MTX-STAFF-Legal BE   0  0  Settings


MAIL
    ---------------------BEL MTX-CEM-CBR Antoing ------------------
    - FatalError
        'Unknown error'
        'Matrix disabled'
        'Excel file incorrect'

    - Warning
        'Archiving failed'


    ----------- Settings
     ComputerName    Path                                         Action  E  W  Duration
     ------------    ----                                         ------  -  -  --------
     DEUSFFRAN0010   E:\DEPARTMENTS\RMC\IB\03-DISTRICT\05-All     Check   0  0  00:00:00
     DEUSFFRAN0010   E:\DEPARTMENTS\RMC\IB\03-DISTRICT\01-North   Check   0  0  00:00:00
     DEUSFFRAN0010   E:\DEPARTMENTS\RMC\IB\03-DISTRICT\03-South   Fix     0  0  00:00:00
     DEUSFFRAN0010   E:\DEPARTMENTS\RMC\IB\03-DISTRICT\04-West    New     0  0  00:00:00


HTML PAGE 1 (File & Permissions)
    File: BEL MTX-CEM-CBR Antoing.xlsx


HTML PAGE 2 (Settings)
    File: BEL MTX-CEM-CBR Antoing.xlsx
    Settings: 
    
    ComputerName    Path                                         Action  E  W  Duration
    ------------    ----                                         ------  -  -  --------
    DEUSFFRAN0010   E:\DEPARTMENTS\RMC\IB\03-DISTRICT\05-All     Check

    Error
    ------------------------------------------------
    Administrator privileges   Administrator privileges are required to be able to apply permissions.
                               "SamAccountName '$env:USERNAME'"

    PowerShell version         PowerShell version 5.1 or higher is required to be able to use advanced methods.
                               PowerShell 3.0

    Warning
    ------------------------------------------------
    Share permissions          Share permissions should always be set to 'FullControl' for the group 'Everyone'. 
                               The effective permissions are managed on the NTFS level.
                               C\Share1
                               C\Share2
                               C\Share3

    Information
    ------------------------------------------------
    Conflichting AD Objects   AD Objects defined in the matrix are duplicate with the ones defined in the default 
                              permissions. In such cases the AD objects in the matrix win over those in the default 
                              permissions. This to ensure a folder can be made completely private to those defined 
                              in the matrix. This can be desired for departments like 'Legal' or 'HR' where data 
                              might contian sensitive information that should not be visible to IT admins defined 
                              in the default permissions.
                              Bob
                              Mike
                              Jake

    Excel file details
    FullName
    Modified
    LastModifiedBy
#>

#region Excel File
[PSCustomObject]@{
	Type = 'FatalError'
	Name = 'Unknown error'
	Description = 'While checking the input and generating the matrix an error was reported.'
	Value = $_
}

[PSCustomObject]@{
	Type = 'Warning'
	Name = 'Archiving failed'
	Description = "When the '-Archive' switch is used the file is moved to the archive folder.In case a file is still in use, the move operation might fail."
	Value = @($_)
}

[PSCustomObject]@{
	Type = 'FatalError'
	Name = 'Excel file incorrect'
	Description = "The worksheets 'Settings' and 'Permissions' are mandatory."
	Value = $M
}

[PSCustomObject]@{
	Type = 'FatalError'
	Name = 'Matrix disabled'
	Description = 'Every Excel file needs at least one enabled matrix.'
	Value = "The worksheet 'Settings' does not contain a row with 'Status' set to 'Enabled'."
}
#endregion

#region Worksheet Permissions
[PSCustomObject]@{
    Type = 'FatalError'
    Name = 'Missing rows'
    Description = 'At least 5 rows are required: 3 header rows. 1 row for the parent folder and at least 1 row for defining permissions on a sub folder.'
    Value = "$(@($Permissions).Count) rows"
}

[PSCustomObject]@{
    Type = 'FatalError'
    Name = 'Missing columns'
    Description = 'At least 2 columns are required: 1 for the folder names and 1 where the permissions are defined.'
    Value = "$(@($Props).Count) column"
}

[PSCustomObject]@{
    Type = 'FatalError'
    Name = 'AD Object not unique'
    Description = "All objects defined in the matrix need to be unique. Duplicate AD Objects can also be generated fromt he 'Settings' worksheet combined with the header rows in the 'Permissions' worksheet."
    Value = $NotUniqueADObjects.Name
}

[PSCustomObject]@{
    Type = 'FatalError'
    Name = 'AD Object name missing'
    Description = "Every column in the worksheet 'Permissions' needs to have an AD object name in the header row. The AD object name can not be blank."
    Value = $null
}

[PSCustomObject]@{
    Type = 'FatalError'
    Name = 'Permission character unknown'
    Description = "The only supported characters, to define permissions on a file or folder, are 'F' (FullControl), 'W' (Write/Modify), 'R' (Read) or 'L' (List)."
    Value = $UnknownPermChar
}

[PSCustomObject]@{
    Type = 'FatalError'
    Name = 'Permissions missing on parent folder'
    Description = 'Missing permissions on the parent folder. At least one permission character needs to be set.'
    Value = $null
}

[PSCustomObject]@{
    Type = 'FatalError'
    Name = 'Permissions missing on parent folder'
    Description = "The permission ignore 'i' cannot be used on the parent folder."
    Value = $null
}

[PSCustomObject]@{
    Type = 'FatalError'
    Name = 'Folder name missing'
    Description = 'Missing folder name in the first column. A folder name is required to be able to set permissions on it.'
    Value = $null
}

[PSCustomObject]@{
    Type = 'FatalError'
    Name = 'Folder name not unique'
    Description = 'Every folder name in the first column needs to be unique. This is required to be able to set the correct permissions.'
    Value = $NotUniqueFolder.Name
}

[PSCustomObject]@{
    Type = 'Warning'
    Name = 'Matrix design flaw'
    Description = "All folders need to be accessible by the end user. Please define at least (R)ead or (W)rite permissions on the deepest folder or use the permnission (I)gnore."
    Value = $inAccessibleFolder
}
#endregion

#region Worksheet Settings: Matrix specific
[PSCustomObject]@{
	Type = 'FatalError'
	Name = 'Administrator privileges'
	Description = "Administrator privileges are required to be able to apply permissions."
	Value = "SamAccountName '$env:USERNAME'"
}

[PSCustomObject]@{
	Type = 'FatalError'
	Name = 'PowerShell version'
	Description = "PowerShell version 5.1 or higher is required to be able to use advanced methods."
	Value = "PowerShell $($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor)"
}

[PSCustomObject]@{
	Type = 'Warning'
	Name = 'Access Based Enumeration'
	Description = "Access Based Enumeration should be set to '$true'. This will hide files and folders where the users don't have access to."
	Value = @($AbeCorrected)
}

[PSCustomObject]@{
	Type = 'Warning'
	Name = 'Share permissions'
	Description = "Share permissions should always be set to 'FullControl' for the group 'Everyone'. The effective permissions are managed on the NTFS level."
	Value = @($SharePermCorrected)
}

[PSCustomObject]@{
    Type = 'FatalError'
    Name = 'Unknown AD object'
    Description = "Every AD object defined in the header row needs to exist before the matrix can be correctly executed."
    Value = $result
}

[PSCustomObject]@{
    Type = 'Warning'
    Name = 'Empty groups'
    Description = 'Every AD Group defined in the header row needs to have at least one user account as a member, excluding the place holder account. Otherwise folders would be inaccessible.'
    Value = $result
}

[PSCustomObject]@{
    Type = 'Warning'
    Name = 'No folder access'
    Description = "Every folder defined in the first column needs to have at least one user account that is able to access it. Group membership is checked to verify if groups granting access to the folder have at least one user account as a member that is not a place holder account."
    Value = $result
}

[PSCustomObject]@{
    Type = 'Information'
    Name = 'Conflichting AD Objects'
    Description = "AD Objects defined in the matrix are duplicate with the ones defined in the default permissions. In such cases the AD objects in the matrix win over those in the default permissions. This to ensure a folder can be made completely private to those defined in the matrix. This can be desired for departments like 'Legal' or 'HR' where data might contian sensitive information that should not be visible to IT admins defined in the default permissions."
    Value = $duplicateADobject
}

[PSCustomObject]@{
	Type = 'FatalError'
	Name = 'Parent folder exists already'
	Description = "The folder defined as 'Path' in the worksheet 'Settings' cannot be present on the remote machine when 'Action=New' is used. Please use 'Action' with value 'Check' or 'Fix' instead."
	Value = $Path
}

[PSCustomObject]@{
	Type = 'FatalError'
	Name = 'Parent folder missing'
	Description = "The folder defined as 'Path' in the worksheet 'Settings' needs to be available on the remote machine. In case the folder structure needs to be created, plase use 'Action=New' instead."
	Value = $Path
}

[PSCustomObject]@{
	Type = 'Information'
	Name = 'Ignored folder'
	Description = "All rows in the worksheet 'Permissions' that have the character 'i' defined are ignored. These folders are not checked for incorrect permissions."
	Value = $IgnoredFolders
}

[PSCustomObject]@{
	Type = 'Warning'
	Name = 'Child folder created'
	Description = "All folders defined in the worksheet 'Permissions' have been created with the correct permissions underneath the parent folder defined in the worksheet 'Settings'."
	Value = $missingFolders.ToArray()
}

[PSCustomObject]@{
	Type = 'Warning'
	Name = 'Non inherited folder incorrect permissions'
	Description = "The folders that have permissions defined in the worksheet 'Permissions' are not matching with the permissions found on the folders of the remote machine."
	Value = if ($DetailedLog) {$incorrectAclNonInheritedFolders} 
			else {$incorrectAclNonInheritedFolders.ToArray()}
}

[PSCustomObject]@{
	Type = 'Warning'
	Name = 'Inherited permissions incorrect'
	Description = "All folders that don't have permissions assigned to them in the worsheet 'Permissions' are supposed to inherit their permissions from the parent folder. Files can only inherit permissions from the parent folder and are not allowed to have explicit permissions."
	Value = if ($DetailedLog) {$incorrectAclInheritedFolders} 
			else {$incorrectAclInheritedFolders.ToArray()}
}

[PSCustomObject]@{
	Type = 'Warning'
	Name = 'Inaccessible data'
	Description = "Files and folders that are found in folders where only list permissions are granted. When no one has read or write permissions, the files/folders become inaccessible."
	Value = $InaccessibleData.ToArray()
}
# Test-MatrixSettingHC:
[PSCustomObject]@{
    Type = 'FatalError'
    Name = 'Missing column header'
    Description = "The column headers 'ComputerName', Path' and 'Action' are mandatory."
    Value = $MissingProperty
}

[PSCustomObject]@{
    Type = 'FatalError'
    Name = 'Missing value'
    Description = "Values for 'ComputerName', Path' and 'Action' are mandatory."
    Value = $BlankProperty
}

[PSCustomObject]@{
    Type = 'FatalError'
    Name = 'Action value incorrect'
    Description = "Only the values 'New', 'Fix' or 'Check' are supported in the field 'Action'."
    Value = $Setting.Action
}

[PSCustomObject]@{
    Type = 'FatalError'
    Name = 'Path value incorrect'
    Description = "The 'Path' needs to be defined as a local folder (Ex. 'E:\Department\Finance')."
    Value = $Setting.Path
}

[PSCustomObject]@{
	Type = 'FatalError'
	Name = 'Duplicate ComputerName/Path combination'
	Description = "Every 'ComputerName' combined with a 'Path' needs to be unique over all the 'Settings' worksheets found in all the active matrix files."
	Value = @{$_.Import.ComputerName = $_.Import.Path}
}
#endregion