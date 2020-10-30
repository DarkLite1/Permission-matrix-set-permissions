#Requires -Version 5.1
#Requires -Modules Pester, Assert

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'

#region Preference
$VerbosePreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'
# $WarningPreference = 'Continue'
# $VerbosePreference = 'Continue'
#endregion

$testUser = 'dverhuls'
$testUser2 = 'wpeeter'

$ExpectedIncorrectAclNonInheritedFolders = [PSCustomObject]@{
    Type        = 'Warning'
    Name        = 'Non inherited folder incorrect permissions'
    Description = "The folders that have permissions defined in the worksheet 'Permissions' are not matching with the permissions found on the folders of the remote machine."
    Value       = $null
}

$ExpectedIncorrectAclInheritedFolders = [PSCustomObject]@{
    Type        = 'Warning'
    Name        = 'Inherited permissions incorrect'
    Description = "All folders that don't have permissions assigned to them in the worksheet 'Permissions' are supposed to inherit their permissions from the parent folder. Files can only inherit permissions from the parent folder and are not allowed to have explicit permissions."
    Value       = $null
}

$ExpectedInaccessibleData = [PSCustomObject]@{
    Type        = 'Warning'
    Name        = 'Inaccessible data'
    Description = "Files and folders that are found in folders where only list permissions are granted. When no one has read or write permissions, the files/folders become inaccessible."
    Value       = $null
}
Function New-TestAceHC {
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
                    # [System.Security.AccessControl.InheritanceFlags]::None, # Required to be absent to set a file acl
                    # [System.Security.AccessControl.PropagationFlags]::None,
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
                    # [System.Security.AccessControl.InheritanceFlags]::None, # Required to be absent to set a file acl
                    # [System.Security.AccessControl.PropagationFlags]::None,
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
                    # [System.Security.AccessControl.InheritanceFlags]::None, # Required to be absent to set a file acl
                    # [System.Security.AccessControl.PropagationFlags]::None, 
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
                    # [System.Security.AccessControl.InheritanceFlags]::None, # Required to be absent to set a file acl
                    # [System.Security.AccessControl.PropagationFlags]::None, 
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

$BuiltinAdmin = [System.Security.Principal.NTAccount]'Builtin\Administrators'


$Skip = $false

Describe $sut {
    Describe 'mandatory parameters' {
        $TestCases = @('Path', 'Action', 'Matrix').ForEach( { @{Name = $_ } })
        it '<Name>' -TestCases $TestCases {
            Param (
                [String]$Name
            )

            (Get-Command "$here\$sut").Parameters[$Name].Attributes.Mandatory | Should -BeTrue
        } -Skip:$Skip
    }
    in $TestDrive {
        Describe 'prepare the desired ACL' {
            BeforeEach {
                Set-Location $TestDrive
            }
            $Params = @{
                Path   = Join-Path $testDrive 'testFolder'
                Action = 'Check'
                Matrix = $null
            }
            New-Item -Path $Params.Path -ItemType Directory -EA Ignore

            it 'create a valid folder ACL' {
                $Params.Matrix = @(
                    [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true }
                )
                ."$here\$sut" @Params

                $Matrix.Where( { $_.FolderAcl}).ForEach( {
                        $_.FolderAcl | Should -BeOfType [System.Security.AccessControl.DirectorySecurity]
                    })
            } -Skip:$Skip
            it "convert multiple hashtables to valid ACL's" {
                $Params.Matrix = @(
                    [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true }
                    [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                )
                ."$here\$sut" @Params

                $Matrix.Where( { $_.Path -eq $Params.Matrix[0].Path }).ForEach( {
                        $_.FolderAcl | Should -BeOfType [System.Security.AccessControl.DirectorySecurity]
                    })

                $Matrix.Where( { $_.Path -eq $Params.Matrix[1].Path }).ForEach( {
                        $_.FolderAcl | Should -BeNullOrEmpty
                    })
            } -Skip:$Skip
            it "the group 'BUILTIN\Administrators' is added to every ACL" {
                $Params = @{
                    Path   = Join-Path $testDrive 'testFolder'
                    Action = 'Check'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                    )
                }

                ."$here\$sut" @Params

                $Matrix.Where( { $_.ACL }).Foreach( {
                        $_.FolderAcl.Access.IdentityReference | Should -Contain 'BUILTIN\Administrators'
                    })
                $Matrix.FolderAcl.Access.Where( { $_.IdentityReference -eq 'BUILTIN\Administrators' }).Foreach( {
                        $_.FileSystemRights | Should -Contain 'FullControl'
                    })
            } -Skip:$Skip
        }
        Describe 'ignored folders' {
            BeforeEach {
                Set-Location $TestDrive
                Remove-Item $TestDrive\* -Recurse -Force
            }
            It 'all files and folders are checked when no folder is ignored' {
                $Params = @{
                    Path   = Join-Path $testDrive 'testFolder'
                    Action = 'Fix'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'R' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                        [PSCustomObject]@{Path = 'FolderB'; ACL = @{$env:USERNAME = 'R' }}
                        [PSCustomObject]@{Path = 'FolderC'; ACL = @{$env:USERNAME = 'R' }}
                        [PSCustomObject]@{Path = 'FolderC\Level1\Level2'; ACL = @{ }}
                    )
                }

                #region Create all folders
                $testParentFolder = New-Item -Path $Params.Path -ItemType Directory

                $Params.Matrix | Select-Object -Skip 1 | ForEach-Object {
                    New-Item -Path (Join-Path $testParentFolder $_.Path) -ItemType Directory -Force
                }
                #endregion

                #region Create all files
                @((Get-ChildItem $testParentFolder -Recurse -Directory) + $testParentFolder) | ForEach-Object {
                    New-Item -Path (Join-Path $_.FullName 'file') -ItemType File
                }
                #endregion

                ."$here\$sut" @Params

                #region Test if all non inherited folders are tested
                $expectedNonInheritanceTested = ($Params.Matrix.Where( { (-not $_.ignore) -and ($_.ACL) })).Path
                 
                foreach ($e in $expectedNonInheritanceTested) { $testedNonInheritedFolders.Keys | Should -Contain  $e }
                $testedNonInheritedFolders.Count | Should -BeExactly $expectedNonInheritanceTested.Count
                #endregion
                
                #region Test if all files and folders that should be inherited are tested
                $expectedInheritanceTested = @(
                    "\\?\$($Params.Path)\file",
                    "\\?\$($Params.Path)\FolderA",
                    "\\?\$($Params.Path)\FolderA\file",
                    "\\?\$($Params.Path)\FolderB\file",
                    "\\?\$($Params.Path)\FolderC\file",
                    "\\?\$($Params.Path)\FolderC\Level1",
                    "\\?\$($Params.Path)\FolderC\Level1\file",
                    "\\?\$($Params.Path)\FolderC\Level1\Level2",
                    "\\?\$($Params.Path)\FolderC\Level1\Level2\file"
                )

                foreach ($e in $expectedInheritanceTested) { $testedInheritedFilesAndFolders.Keys | Should -Contain $e }
                $testedInheritedFilesAndFolders.Count | Should -BeExactly $expectedInheritanceTested.Count
                #endregion
            } -Skip:$Skip
            It 'when a folder is ignored all its subfolders and files are not checked' {
                $Params = @{
                    Path   = Join-Path $testDrive 'testFolder'
                    Action = 'Fix'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'R' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                        [PSCustomObject]@{Path = 'FolderB'; ACL = @{$env:USERNAME = 'R' }; Ignore = $true }
                        [PSCustomObject]@{Path = 'FolderC'; ACL = @{$env:USERNAME = 'R' }}
                    )
                }
                #region Create all folders
                $testParentFolder = New-Item -Path $Params.Path -ItemType Directory

                $Params.Matrix | Select-Object -Skip 1 | ForEach-Object {
                    New-Item -Path (Join-Path $testParentFolder $_.Path) -ItemType Directory -Force
                }
                #endregion

                #region Create all files
                @((Get-ChildItem $testParentFolder -Recurse -Directory) + $testParentFolder) | ForEach-Object {
                    New-Item -Path (Join-Path $_.FullName 'file') -ItemType File
                }
                #endregion

                ."$here\$sut" @Params

                #region Test if all non inherited folders are tested
                $expectedNonInheritanceTested = ($Params.Matrix.Where( { (-not $_.ignore) -and ($_.ACL) })).Path
                 
                foreach ($e in $expectedNonInheritanceTested) { $testedNonInheritedFolders.Keys | Should -Contain  $e }
                $testedNonInheritedFolders.Count | Should -BeExactly $expectedNonInheritanceTested.Count
                #endregion
                
                #region Test if all files and folders that should be inherited are tested
                $expectedInheritanceTested = @(
                    "\\?\$($Params.Path)\file",
                    "\\?\$($Params.Path)\FolderA",
                    "\\?\$($Params.Path)\FolderA\file",
                    "\\?\$($Params.Path)\FolderC\file"
                )
                
                foreach ($e in $expectedInheritanceTested) { $testedInheritedFilesAndFolders.Keys | Should -Contain $e }
                $testedInheritedFilesAndFolders.Count | Should -BeExactly $expectedInheritanceTested.Count
                #endregion
            } -Skip:$Skip
            It 'Sub folders of ignored folders are not checked (1)' {
                $Params = @{
                    Path   = Join-Path $testDrive 'testFolder'
                    Action = 'Fix'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'R' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                        [PSCustomObject]@{Path = 'FolderB\SubFolder'; ACL = @{$env:USERNAME = 'R' }; Ignore = $true }
                        [PSCustomObject]@{Path = 'FolderB\SubFolder\Reports'; ACL = @{$env:USERNAME = 'R' }}
                        [PSCustomObject]@{Path = 'FolderB\SubFolder\Accounting'; ACL = @{ }} # ignored because SubFolder is ignored
                    )
                }
                #region Create all folders
                $testParentFolder = New-Item -Path $Params.Path -ItemType Directory

                $Params.Matrix | Select-Object -Skip 1 | ForEach-Object {
                    New-Item -Path (Join-Path $testParentFolder $_.Path) -ItemType Directory -Force
                }
                #endregion

                #region Create all files
                @((Get-ChildItem $testParentFolder -Recurse -Directory) + $testParentFolder) | ForEach-Object {
                    New-Item -Path (Join-Path $_.FullName 'file') -ItemType File
                }
                #endregion

                ."$here\$sut" @Params

                #region Test if all non inherited folders are tested
                $expectedNonInheritanceTested = ($Params.Matrix.Where( { (-not $_.ignore) -and ($_.ACL) })).Path
                 
                foreach ($e in $expectedNonInheritanceTested) { $testedNonInheritedFolders.Keys | Should -Contain  $e }
                $testedNonInheritedFolders.Count | Should -BeExactly $expectedNonInheritanceTested.Count
                #endregion
                
                #region Test if all files and folders that should be inherited are tested
                $expectedInheritanceTested = @(
                    "\\?\$($Params.Path)\file",
                    "\\?\$($Params.Path)\FolderA",
                    "\\?\$($Params.Path)\FolderA\file",
                    "\\?\$($Params.Path)\FolderB",
                    "\\?\$($Params.Path)\FolderB\File",
                    "\\?\$($Params.Path)\FolderB\SubFolder\Reports\File"
                )
                
                foreach ($e in $expectedInheritanceTested) { $testedInheritedFilesAndFolders.Keys | Should -Contain $e }
                $testedInheritedFilesAndFolders.Count | Should -BeExactly $expectedInheritanceTested.Count
                #endregion
            } -Skip:$Skip
            It 'Sub folders of ignored folders are not checked (2)' {
                $Params = @{
                    Path   = Join-Path $testDrive 'testFolder'
                    Action = 'Fix'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'R' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                        [PSCustomObject]@{Path = 'FolderB\SubFolder'; ACL = @{$env:USERNAME = 'R' }; Ignore = $true }
                        [PSCustomObject]@{Path = 'FolderB\SubFolder\Reports'; ACL = @{$env:USERNAME = 'R' }}
                        [PSCustomObject]@{Path = 'FolderB\SubFolder\Reports\Year\2020'; ACL = @{ }} # tested because it falls under the acl tree Reports
                        [PSCustomObject]@{Path = 'FolderB\SubFolder\Accounting'; ACL = @{ }} # ignored because SubFolder is ignored
                        [PSCustomObject]@{Path = 'FolderB\SubFolder\Sales'; ACL = @{ }} # ignored because SubFolder is ignored
                        [PSCustomObject]@{Path = 'FolderC'; ACL = @{$env:USERNAME = 'R' }}
                    )
                }
                #region Create all folders
                $testParentFolder = New-Item -Path $Params.Path -ItemType Directory

                $Params.Matrix | Select-Object -Skip 1 | ForEach-Object {
                    New-Item -Path (Join-Path $testParentFolder $_.Path) -ItemType Directory -Force
                }
                #endregion

                #region Create all files
                @((Get-ChildItem $testParentFolder -Recurse -Directory) + $testParentFolder) | ForEach-Object {
                    New-Item -Path (Join-Path $_.FullName 'file') -ItemType File
                }
                #endregion

                . "$here\$sut" @Params

                #region Test if all non inherited folders are tested
                $expectedNonInheritanceTested = ($Params.Matrix.Where( { (-not $_.ignore) -and ($_.ACL) })).Path
                 
                foreach ($e in $expectedNonInheritanceTested) { $testedNonInheritedFolders.Keys | Should -Contain  $e }
                $testedNonInheritedFolders.Count | Should -BeExactly $expectedNonInheritanceTested.Count
                #endregion
                
                #region Test if all files and folders that should be inherited are tested
                $expectedInheritanceTested = @(
                    "\\?\$($Params.Path)\file",
                    "\\?\$($Params.Path)\FolderA",
                    "\\?\$($Params.Path)\FolderA\file",
                    "\\?\$($Params.Path)\FolderB",
                    "\\?\$($Params.Path)\FolderB\File",
                    "\\?\$($Params.Path)\FolderB\SubFolder\Reports\File",
                    "\\?\$($Params.Path)\FolderB\SubFolder\Reports\Year",
                    "\\?\$($Params.Path)\FolderB\SubFolder\Reports\Year\File",
                    "\\?\$($Params.Path)\FolderB\SubFolder\Reports\Year\2020",
                    "\\?\$($Params.Path)\FolderB\SubFolder\Reports\Year\2020\File",
                    "\\?\$($Params.Path)\FolderC\File"
                )
                
                foreach ($e in $expectedInheritanceTested) { 
                    $testedInheritedFilesAndFolders.Keys | Should -Contain $e 
                }
                $testedInheritedFilesAndFolders.Count | Should -BeExactly $expectedInheritanceTested.Count
                #endregion
            } -Skip:$Skip
            It 'Sub folders of ignored folders are not checked (3)' {
                $Params = @{
                    Path   = Join-Path $testDrive 'testFolder'
                    Action = 'Fix'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'R' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                        [PSCustomObject]@{Path = 'FolderB\SubFolder'; ACL = @{$env:USERNAME = 'R' }; Ignore = $true }
                        [PSCustomObject]@{Path = 'FolderB\SubFolder\Reports'; ACL = @{$env:USERNAME = 'R' }}
                        [PSCustomObject]@{Path = 'FolderB\SubFolder\Reports\Year\2020'; ACL = @{ }} # tested because it falls under the acl tree Reports
                        [PSCustomObject]@{Path = 'FolderB\SubFolder\Reports\Year\2020\CEM'; ACL = @{ }; Ignore = $true } 
                        [PSCustomObject]@{Path = 'FolderB\SubFolder\Reports\Year\2020\CEM\Profit'; ACL = @{ }} 
                        [PSCustomObject]@{Path = 'FolderB\SubFolder\Reports\Year\2020\CEM\Loss'; ACL = @{ }} 
                        [PSCustomObject]@{Path = 'FolderB\SubFolder\Reports\Year\2020\CEM\Loss\HR'; ACL = @{ $env:USERNAME = 'R' }} 
                        [PSCustomObject]@{Path = 'FolderB\SubFolder\Accounting'; ACL = @{ }} # ignored because SubFolder is ignored
                        [PSCustomObject]@{Path = 'FolderB\SubFolder\Sales'; ACL = @{ }} # ignored because SubFolder is ignored
                        [PSCustomObject]@{Path = 'FolderC'; ACL = @{$env:USERNAME = 'R' }}
                    )
                }
                #region Create all folders
                $testParentFolder = New-Item -Path $Params.Path -ItemType Directory

                $Params.Matrix | Select-Object -Skip 1 | ForEach-Object {
                    New-Item -Path (Join-Path $testParentFolder $_.Path) -ItemType Directory -Force
                }
                #endregion

                #region Create all files
                @((Get-ChildItem $testParentFolder -Recurse -Directory) + $testParentFolder) | ForEach-Object {
                    New-Item -Path (Join-Path $_.FullName 'file') -ItemType File
                }
                #endregion

                ."$here\$sut" @Params

                #region Test if all non inherited folders are tested
                $expectedNonInheritanceTested = ($Params.Matrix.Where( { (-not $_.ignore) -and ($_.ACL) })).Path
                 
                foreach ($e in $expectedNonInheritanceTested) { $testedNonInheritedFolders.Keys | Should -Contain  $e }
                $testedNonInheritedFolders.Count | Should -BeExactly $expectedNonInheritanceTested.Count
                #endregion
                
                #region Test if all files and folders that should be inherited are tested
                $expectedInheritanceTested = @(
                    "\\?\$($Params.Path)\file",
                    "\\?\$($Params.Path)\FolderA",
                    "\\?\$($Params.Path)\FolderA\file",
                    "\\?\$($Params.Path)\FolderB",
                    "\\?\$($Params.Path)\FolderB\File",
                    "\\?\$($Params.Path)\FolderB\SubFolder\Reports\File",
                    "\\?\$($Params.Path)\FolderB\SubFolder\Reports\Year",
                    "\\?\$($Params.Path)\FolderB\SubFolder\Reports\Year\File",
                    "\\?\$($Params.Path)\FolderB\SubFolder\Reports\Year\2020",
                    "\\?\$($Params.Path)\FolderB\SubFolder\Reports\Year\2020\File",
                    "\\?\$($Params.Path)\FolderB\SubFolder\Reports\Year\2020\CEM\Loss\HR\File",
                    "\\?\$($Params.Path)\FolderC\File"
                )
                
                foreach ($e in $expectedInheritanceTested) { $testedInheritedFilesAndFolders.Keys | Should -Contain $e }
                $testedInheritedFilesAndFolders.Count | Should -BeExactly $expectedInheritanceTested.Count
                #endregion
            } -Skip:$Skip
            It 'ignored folders are stored in an information object' {
                $Params = @{
                    Path   = Join-Path $testDrive 'testFolder'
                    Action = 'Fix'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'R' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                        [PSCustomObject]@{Path = 'FolderB\SubFolder'; ACL = @{$env:USERNAME = 'R' }; Ignore = $true }
                        [PSCustomObject]@{Path = 'FolderB\SubFolder\Reports'; ACL = @{$env:USERNAME = 'R' }}
                        [PSCustomObject]@{Path = 'FolderB\SubFolder\Reports\Year\2020'; ACL = @{ }} # tested because it falls under the acl tree Reports
                        [PSCustomObject]@{Path = 'FolderB\SubFolder\Reports\Year\2020\CEM'; ACL = @{ }; Ignore = $true } 
                        [PSCustomObject]@{Path = 'FolderB\SubFolder\Reports\Year\2020\CEM\Profit'; ACL = @{ }} 
                        [PSCustomObject]@{Path = 'FolderB\SubFolder\Reports\Year\2020\CEM\Loss'; ACL = @{ }} 
                        [PSCustomObject]@{Path = 'FolderB\SubFolder\Reports\Year\2020\CEM\Loss\HR'; ACL = @{ $env:USERNAME = 'R' }} 
                        [PSCustomObject]@{Path = 'FolderB\SubFolder\Accounting'; ACL = @{ }} # ignored because SubFolder is ignored
                        [PSCustomObject]@{Path = 'FolderB\SubFolder\Sales'; ACL = @{ }} # ignored because SubFolder is ignored
                        [PSCustomObject]@{Path = 'FolderC'; ACL = @{$env:USERNAME = 'R' }}
                    )
                }
                
                $null = New-Item -Path $Params.Path -ItemType Directory

                #region Test if in an information object is created
                $Expected = [PSCustomObject]@{
                    Type  = 'Information'
                    Name  = 'Ignored folder'
                    Value = @(
                        "$($Params.Path)\FolderB\SubFolder",
                        "$($Params.Path)\FolderB\SubFolder\Reports\Year\2020\CEM"
                    )
                }

                $Actual = ."$here\$sut" @Params | Where-Object { ($_.Type -eq $Expected.Type) -and
                    ($_.Name -eq $Expected.Name) }

                $Actual.Type | Should -Be $Expected.Type
                $Actual.Name | Should -Be $Expected.Name
                $Actual.Value | Should -Be $Expected.Value
                #endregion
            } -Skip:$Skip
        }
        Describe 'Permissions' {
            BeforeEach {
                Set-Location $TestDrive
                Remove-Item $TestDrive\* -Recurse -Force
            }
            Context 'are not corrected when they are correct for' {
                It 'List, Write, Read on the parent folder' {
                    $Params = @{
                        Path   = Join-Path $testDrive 'testFolder'
                        Action = 'Fix'
                        Matrix = @(
                            [PSCustomObject]@{Path = 'Path'; ACL = @{
                                    $env:USERNAME = 'L' ; $testUser = 'W'; $testUser2 = 'R';
                                }; Parent = $true 
                            }
                            [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                        )
                    }

                    #region Create all folders
                    $Params.Matrix | Select-Object -Skip 1 | ForEach-Object {
                        New-Item -Path (Join-Path $Params.Path $_.Path) -ItemType Directory -Force
                    }
                    #endregion

                    #region Create all files
                    @(, (Get-ChildItem $Params.Path -Recurse -Directory).FullName + $Params.Path) | 
                    ForEach-Object {
                        New-Item -Path (Join-Path $_ 'file') -ItemType File
                    }
                    #endregion

                    #region Set correct permissions on parent folder
                    $testItem = Get-Item $Params.Path

                    $acl = New-Object System.Security.AccessControl.DirectorySecurity
                    $acl.SetAccessRuleProtection($true, $false)
                    $acl.SetOwner($BuiltinAdmin)

                    $aceList = @($AdminFullControlFolderAce)
                    $aceList += New-TestAceHC -Type 'Folder' -Access 'L' -Name $env:USERNAME
                    $aceList += New-TestAceHC -Type 'Folder' -Access 'W' -Name $testUser
                    $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser2
                    $aceList.foreach( { $acl.AddAccessRule($_) })

                    $testItem.SetAccessControl($acl)
                    #endregion

                    $Actual = ."$here\$sut" @Params | Where-Object {
                        ($_Name -eq 'Non inherited folder incorrect permissions') -or
                        ($_Name -eq 'Inherited permissions incorrect')
                    }

                    $Actual | Should -BeNullOrEmpty
                } -Skip:$Skip
                It 'List only on the parent folder' {
                    $Params = @{
                        Path   = Join-Path $testDrive 'testFolder'
                        Action = 'Fix'
                        Matrix = @(
                            [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L'; $testUser = 'L' }; Parent = $true }
                            [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                        )
                    }

                    #region Create all folders
                    $Params.Matrix | Select-Object -Skip 1 | ForEach-Object {
                        New-Item -Path (Join-Path $Params.Path $_.Path) -ItemType Directory -Force
                    }
                    #endregion

                    #region Create all files
                    @(, (Get-ChildItem $Params.Path -Recurse -Directory).FullName + $Params.Path) | 
                    ForEach-Object {
                        New-Item -Path (Join-Path $_ 'file') -ItemType File
                    }
                    #endregion

                    #region Set correct permissions on parent folder
                    $testItem = Get-Item $Params.Path

                    $acl = New-Object System.Security.AccessControl.DirectorySecurity
                    $acl.SetAccessRuleProtection($true, $false)
                    $acl.SetOwner($BuiltinAdmin)

                    $aceList = @($AdminFullControlFolderAce)
                    $aceList += New-TestAceHC -Type 'Folder' -Access 'L' -Name $env:USERNAME
                    $aceList += New-TestAceHC -Type 'Folder' -Access 'L' -Name $testUser
                    $aceList.foreach( { $acl.AddAccessRule($_) })

                    $testItem.SetAccessControl($acl)
                    #endregion

                    $Actual = ."$here\$sut" @Params | Where-Object {
                        ($_Name -eq 'Non inherited folder incorrect permissions') -or
                        ($_Name -eq 'Inherited permissions incorrect')
                    }

                    $Actual | Should -BeNullOrEmpty
                } -Skip:$Skip
                It 'List only on the parent folder and Read on a subfolder' {
                    $Params = @{
                        Path   = Join-Path $testDrive 'testFolder'
                        Action = 'Fix'
                        Matrix = @(
                            [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L'; $testUser = 'L' }; Parent = $true }
                            [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser = 'R' } }
                            [PSCustomObject]@{Path = 'FolderB'; ACL = @{ } }
                        )
                    }

                    #region Create all folders
                    $Params.Matrix | Select-Object -Skip 1 | ForEach-Object {
                        New-Item -Path (Join-Path $Params.Path $_.Path) -ItemType Directory -Force
                    }
                    #endregion

                    #region Create all files
                    @(, (Get-ChildItem $Params.Path -Recurse -Directory).FullName + $Params.Path) | 
                    ForEach-Object {
                        New-Item -Path (Join-Path $_ 'file') -ItemType File
                    }
                    #endregion

                    #region Set correct permissions on parent folder
                    $testItem = Get-Item $Params.Path

                    $acl = New-Object System.Security.AccessControl.DirectorySecurity
                    $acl.SetAccessRuleProtection($true, $false)
                    $acl.SetOwner($BuiltinAdmin)

                    $aceList = @($AdminFullControlFolderAce)
                    $aceList += New-TestAceHC -Type 'Folder' -Access 'L' -Name $env:USERNAME
                    $aceList += New-TestAceHC -Type 'Folder' -Access 'L' -Name $testUser
                    $aceList.foreach( { $acl.AddAccessRule($_) })

                    $testItem.SetAccessControl($acl)
                    #endregion

                    #region Set correct permissions on a sub folder
                    $testItem = Get-Item "$($Params.Path)\FolderA"

                    $acl = New-Object System.Security.AccessControl.DirectorySecurity
                    $acl.SetAccessRuleProtection($true, $false)
                    $acl.SetOwner($BuiltinAdmin)

                    $aceList = @($AdminFullControlFolderAce)
                    $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser
                    $aceList.foreach( { $acl.AddAccessRule($_) })

                    $testItem.SetAccessControl($acl)
                    #endregion

                    $Actual = ."$here\$sut" @Params | Where-Object {
                        ($_Name -eq 'Non inherited folder incorrect permissions') -or
                        ($_Name -eq 'Inherited permissions incorrect')
                    }

                    $Actual | Should -BeNullOrEmpty
                } -Skip:$Skip
                It 'List only on the parent folder and different permissions on subfolders' {
                    $Params = @{
                        Path   = Join-Path $testDrive 'testFolder'
                        Action = 'Fix'
                        Matrix = @(
                            [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L'; $testUser = 'L'; $testUser2 = 'L' }; Parent = $true }
                            [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser = 'R' } }
                            [PSCustomObject]@{Path = 'FolderB'; ACL = @{ } }
                            [PSCustomObject]@{Path = 'FolderB\SubFolderB'; ACL = @{$testUser2 = 'W' } }
                            [PSCustomObject]@{Path = 'FolderB\SubFolderC'; ACL = @{$testUser2 = 'R' } }
                        )
                    }

                    #region Create all folders
                    $Params.Matrix | Select-Object -Skip 1 | ForEach-Object {
                        New-Item -Path (Join-Path $Params.Path $_.Path) -ItemType Directory -Force
                    }
                    #endregion

                    #region Create all files
                    @(, (Get-ChildItem $Params.Path -Recurse -Directory).FullName + $Params.Path) | 
                    ForEach-Object {
                        New-Item -Path (Join-Path $_ 'file') -ItemType File
                    }
                    #endregion

                    #region Set correct permissions on parent folder
                    $testItem = Get-Item $Params.Path

                    $acl = New-Object System.Security.AccessControl.DirectorySecurity
                    $acl.SetAccessRuleProtection($true, $false)
                    $acl.SetOwner($BuiltinAdmin)

                    $aceList = @($AdminFullControlFolderAce)
                    $aceList += New-TestAceHC -Type 'Folder' -Access 'L' -Name $env:USERNAME
                    $aceList += New-TestAceHC -Type 'Folder' -Access 'L' -Name $testUser
                    $aceList += New-TestAceHC -Type 'Folder' -Access 'L' -Name $testUser2
                    $aceList.foreach( { $acl.AddAccessRule($_) })

                    $testItem.SetAccessControl($acl)
                    #endregion

                    #region Set correct permissions on a sub folder
                    $testItem = Get-Item "$($Params.Path)\FolderA"

                    $acl = New-Object System.Security.AccessControl.DirectorySecurity
                    $acl.SetAccessRuleProtection($true, $false)
                    $acl.SetOwner($BuiltinAdmin)

                    $aceList = @($AdminFullControlFolderAce)
                    $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser
                    $aceList.foreach( { $acl.AddAccessRule($_) })

                    $testItem.SetAccessControl($acl)
                    #endregion

                    #region Set correct permissions on a sub folder
                    $testItem = Get-Item "$($Params.Path)\FolderB\SubFolderB"

                    $acl = New-Object System.Security.AccessControl.DirectorySecurity
                    $acl.SetAccessRuleProtection($true, $false)
                    $acl.SetOwner($BuiltinAdmin)

                    $aceList = @($AdminFullControlFolderAce)
                    $aceList += New-TestAceHC -Type 'Folder' -Access 'W' -Name $testUser2
                    $aceList.foreach( { $acl.AddAccessRule($_) })

                    $testItem.SetAccessControl($acl)
                    #endregion

                    #region Set correct permissions on a sub folder
                    $testItem = Get-Item "$($Params.Path)\FolderB\SubFolderC"

                    $acl = New-Object System.Security.AccessControl.DirectorySecurity
                    $acl.SetAccessRuleProtection($true, $false)
                    $acl.SetOwner($BuiltinAdmin)

                    $aceList = @($AdminFullControlFolderAce)
                    $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser2
                    $aceList.foreach( { $acl.AddAccessRule($_) })

                    $testItem.SetAccessControl($acl)
                    #endregion

                    $Actual = ."$here\$sut" @Params | Where-Object {
                        ($_Name -eq 'Non inherited folder incorrect permissions') -or
                        ($_Name -eq 'Inherited permissions incorrect')
                    }

                    $Actual | Should -BeNullOrEmpty
                } -Skip:$Skip
                It 'folders that are not in the matrix as they should be inherited' {
                    $Params = @{
                        Path   = Join-Path $testDrive 'testFolder'
                        Action = 'Fix'
                        Matrix = @(
                            [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L'; $testUser = 'L'; $testUser2 = 'L' }; Parent = $true }
                            [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser = 'R' } }
                            [PSCustomObject]@{Path = 'FolderB'; ACL = @{ } }
                            [PSCustomObject]@{Path = 'FolderB\SubFolderB'; ACL = @{$testUser2 = 'W' } }
                            [PSCustomObject]@{Path = 'FolderB\SubFolderC'; ACL = @{$testUser2 = 'R' } }
                        )
                    }

                    #region Create all folders
                    $Params.Matrix | Select-Object -Skip 1 | ForEach-Object {
                        New-Item -Path (Join-Path $Params.Path $_.Path) -ItemType Directory -Force
                    }
                    #endregion

                    #region Create all files
                    @(, (Get-ChildItem $Params.Path -Recurse -Directory).FullName + $Params.Path) | 
                    ForEach-Object {
                        New-Item -Path (Join-Path $_ 'file') -ItemType File
                    }
                    #endregion

                    #region Set correct permissions on parent folder
                    $testItem = Get-Item $Params.Path

                    $acl = New-Object System.Security.AccessControl.DirectorySecurity
                    $acl.SetAccessRuleProtection($true, $false)
                    $acl.SetOwner($BuiltinAdmin)

                    $aceList = @($AdminFullControlFolderAce)
                    $aceList += New-TestAceHC -Type 'Folder' -Access 'L' -Name $env:USERNAME
                    $aceList += New-TestAceHC -Type 'Folder' -Access 'L' -Name $testUser
                    $aceList += New-TestAceHC -Type 'Folder' -Access 'L' -Name $testUser2
                    $aceList.foreach( { $acl.AddAccessRule($_) })

                    $testItem.SetAccessControl($acl)
                    #endregion

                    #region Set correct permissions on a sub folder
                    $testItem = Get-Item "$($Params.Path)\FolderA"

                    $acl = New-Object System.Security.AccessControl.DirectorySecurity
                    $acl.SetAccessRuleProtection($true, $false)
                    $acl.SetOwner($BuiltinAdmin)

                    $aceList = @($AdminFullControlFolderAce)
                    $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser
                    $aceList.foreach( { $acl.AddAccessRule($_) })

                    $testItem.SetAccessControl($acl)
                    #endregion

                    #region Set correct permissions on a sub folder
                    $testItem = Get-Item "$($Params.Path)\FolderB\SubFolderB"

                    $acl = New-Object System.Security.AccessControl.DirectorySecurity
                    $acl.SetAccessRuleProtection($true, $false)
                    $acl.SetOwner($BuiltinAdmin)

                    $aceList = @($AdminFullControlFolderAce)
                    $aceList += New-TestAceHC -Type 'Folder' -Access 'W' -Name $testUser2
                    $aceList.foreach( { $acl.AddAccessRule($_) })

                    $testItem.SetAccessControl($acl)
                    #endregion

                    #region Set correct permissions on a sub folder
                    $testItem = Get-Item "$($Params.Path)\FolderB\SubFolderC"

                    $acl = New-Object System.Security.AccessControl.DirectorySecurity
                    $acl.SetAccessRuleProtection($true, $false)
                    $acl.SetOwner($BuiltinAdmin)

                    $aceList = @($AdminFullControlFolderAce)
                    $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser2
                    $aceList.foreach( { $acl.AddAccessRule($_) })

                    $testItem.SetAccessControl($acl)
                    #endregion

                    #region Create extra folders not defined in the matrix
                    $null = New-Item -Path "$($Params.Path)\FolderB\OtherFolder" -ItemType Directory -Force
                    $null = New-Item -Path "$($Params.Path)\FolderC" -ItemType Directory -Force
                    $null = New-Item -Path "$($Params.Path)\FolderC\File" -ItemType File -Force
                    $null = New-Item -Path "$($Params.Path)\FolderD\Fruits\Kiwi\Green" -ItemType Directory -Force
                    $null = New-Item -Path "$($Params.Path)\FolderD\Fruits\Kiwi\Green\File" -ItemType File -Force
                    #endregion

                    $Actual = ."$here\$sut" @Params | Where-Object {
                        ($_Name -eq 'Non inherited folder incorrect permissions') -or
                        ($_Name -eq 'Inherited permissions incorrect')
                    }

                    $Actual | Should -BeNullOrEmpty
                } -Skip:$Skip
            }
            Context 'are corrected when they are incorrect when' {
                Context 'a folder that should have explicit permissions has' {
                    It 'incorrect explicit permissions' {
                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'Fix'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L'; $testUser2 = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser = 'R' } }
                                [PSCustomObject]@{Path = 'FolderB'; ACL = @{ } }
                                [PSCustomObject]@{Path = 'FolderB\SubFolderB'; ACL = @{$testUser2 = 'W' } }
                                [PSCustomObject]@{Path = 'FolderB\SubFolderC'; ACL = @{$testUser2 = 'R' } }
                            )
                        }

                        #region Create all folders
                        $Params.Matrix | Select-Object -Skip 1 | ForEach-Object {
                            New-Item -Path (Join-Path $Params.Path $_.Path) -ItemType Directory -Force
                        }
                        #endregion

                        #region Create all files
                        @(, (Get-ChildItem $Params.Path -Recurse -Directory).FullName + $Params.Path) | 
                        ForEach-Object {
                            New-Item -Path (Join-Path $_ 'file') -ItemType File
                        }
                        #endregion

                        #region Set incorrect permissions on parent folder
                        $testItem = Get-Item $Params.Path

                        $acl = New-Object System.Security.AccessControl.DirectorySecurity
                        $acl.SetAccessRuleProtection($true, $false)
                        $acl.SetOwner($BuiltinAdmin)

                        $aceList = @($AdminFullControlFolderAce)
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser # Incorrect
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'L' -Name $testUser2
                        $aceList.foreach( { $acl.AddAccessRule($_) })

                        $testItem.SetAccessControl($acl)
                        #endregion

                        #region Set correct permissions on a sub folder
                        $testItem = Get-Item "$($Params.Path)\FolderA"

                        $acl = New-Object System.Security.AccessControl.DirectorySecurity
                        $acl.SetAccessRuleProtection($true, $false)
                        $acl.SetOwner($BuiltinAdmin)

                        $aceList = @($AdminFullControlFolderAce)
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser
                        $aceList.foreach( { $acl.AddAccessRule($_) })

                        $testItem.SetAccessControl($acl)
                        #endregion

                        #region Set correct permissions on a sub folder
                        $testItem = Get-Item "$($Params.Path)\FolderB\SubFolderB"

                        $acl = New-Object System.Security.AccessControl.DirectorySecurity
                        $acl.SetAccessRuleProtection($true, $false)
                        $acl.SetOwner($BuiltinAdmin)

                        $aceList = @($AdminFullControlFolderAce)
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'W' -Name $testUser2
                        $aceList.foreach( { $acl.AddAccessRule($_) })

                        $testItem.SetAccessControl($acl)
                        #endregion

                        #region Set correct permissions on a sub folder
                        $testItem = Get-Item "$($Params.Path)\FolderB\SubFolderC"

                        $acl = New-Object System.Security.AccessControl.DirectorySecurity
                        $acl.SetAccessRuleProtection($true, $false)
                        $acl.SetOwner($BuiltinAdmin)

                        $aceList = @($AdminFullControlFolderAce)
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser2
                        $aceList.foreach( { $acl.AddAccessRule($_) })

                        $testItem.SetAccessControl($acl)
                        #endregion

                        $Actual = ."$here\$sut" @Params | Where-Object Name -eq 'Non inherited folder incorrect permissions'
                    
                        $Actual.Value | Should -Be $Params.Path
                    } -Skip:$Skip
                    It 'the correct explicit permissions but one ACE too much' {
                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'Fix'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L'; $testUser2 = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser = 'R' } }
                                [PSCustomObject]@{Path = 'FolderB'; ACL = @{ } }
                                [PSCustomObject]@{Path = 'FolderB\SubFolderB'; ACL = @{$testUser2 = 'W' } }
                                [PSCustomObject]@{Path = 'FolderB\SubFolderC'; ACL = @{$testUser2 = 'R' } }
                            )
                        }

                        #region Create all folders
                        $Params.Matrix | Select-Object -Skip 1 | ForEach-Object {
                            New-Item -Path (Join-Path $Params.Path $_.Path) -ItemType Directory -Force
                        }
                        #endregion

                        #region Create all files
                        @(, (Get-ChildItem $Params.Path -Recurse -Directory).FullName + $Params.Path) | 
                        ForEach-Object {
                            New-Item -Path (Join-Path $_ 'file') -ItemType File
                        }
                        #endregion

                        #region Set incorrect permissions on parent folder
                        $testItem = Get-Item $Params.Path

                        $acl = New-Object System.Security.AccessControl.DirectorySecurity
                        $acl.SetAccessRuleProtection($true, $false)
                        $acl.SetOwner($BuiltinAdmin)

                        $aceList = @($AdminFullControlFolderAce)
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $env:USERNAME # Incorrect
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'L' -Name $testUser
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'L' -Name $testUser2
                        $aceList.foreach( { $acl.AddAccessRule($_) })

                        $testItem.SetAccessControl($acl)
                        #endregion

                        #region Set correct permissions on a sub folder
                        $testItem = Get-Item "$($Params.Path)\FolderA"

                        $acl = New-Object System.Security.AccessControl.DirectorySecurity
                        $acl.SetAccessRuleProtection($true, $false)
                        $acl.SetOwner($BuiltinAdmin)

                        $aceList = @($AdminFullControlFolderAce)
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser
                        $aceList.foreach( { $acl.AddAccessRule($_) })

                        $testItem.SetAccessControl($acl)
                        #endregion

                        #region Set correct permissions on a sub folder
                        $testItem = Get-Item "$($Params.Path)\FolderB\SubFolderB"

                        $acl = New-Object System.Security.AccessControl.DirectorySecurity
                        $acl.SetAccessRuleProtection($true, $false)
                        $acl.SetOwner($BuiltinAdmin)

                        $aceList = @($AdminFullControlFolderAce)
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'W' -Name $testUser2
                        $aceList.foreach( { $acl.AddAccessRule($_) })

                        $testItem.SetAccessControl($acl)
                        #endregion

                        #region Set correct permissions on a sub folder
                        $testItem = Get-Item "$($Params.Path)\FolderB\SubFolderC"

                        $acl = New-Object System.Security.AccessControl.DirectorySecurity
                        $acl.SetAccessRuleProtection($true, $false)
                        $acl.SetOwner($BuiltinAdmin)

                        $aceList = @($AdminFullControlFolderAce)
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser2
                        $aceList.foreach( { $acl.AddAccessRule($_) })

                        $testItem.SetAccessControl($acl)
                        #endregion

                        $Actual = ."$here\$sut" @Params | Where-Object Name -eq 'Non inherited folder incorrect permissions'
                    
                        $Actual.Value | Should -Be $Params.Path
                    } -Skip:$Skip
                    It 'inherited permissions' {
                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'Fix'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L'; $testUser2 = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser = 'R' } }
                                [PSCustomObject]@{Path = 'FolderB'; ACL = @{ } }
                                [PSCustomObject]@{Path = 'FolderB\SubFolderB'; ACL = @{$testUser2 = 'W' } }
                                [PSCustomObject]@{Path = 'FolderB\SubFolderC'; ACL = @{$testUser2 = 'R' } }
                            )
                        }

                        #region Create all folders
                        $Params.Matrix | Select-Object -Skip 1 | ForEach-Object {
                            New-Item -Path (Join-Path $Params.Path $_.Path) -ItemType Directory -Force
                        }
                        #endregion

                        #region Create all files
                        @(, (Get-ChildItem $Params.Path -Recurse -Directory).FullName + $Params.Path) | 
                        ForEach-Object {
                            New-Item -Path (Join-Path $_ 'file') -ItemType File
                        }
                        #endregion

                        #region Set incorrect permissions on parent folder
                        # $Params.Path is inherited
                        #endregion

                        #region Set correct permissions on a sub folder
                        $testItem = Get-Item "$($Params.Path)\FolderA"

                        $acl = New-Object System.Security.AccessControl.DirectorySecurity
                        $acl.SetAccessRuleProtection($true, $false)
                        $acl.SetOwner($BuiltinAdmin)

                        $aceList = @($AdminFullControlFolderAce)
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser
                        $aceList.foreach( { $acl.AddAccessRule($_) })

                        $testItem.SetAccessControl($acl)
                        #endregion

                        #region Set correct permissions on a sub folder
                        $testItem = Get-Item "$($Params.Path)\FolderB\SubFolderB"

                        $acl = New-Object System.Security.AccessControl.DirectorySecurity
                        $acl.SetAccessRuleProtection($true, $false)
                        $acl.SetOwner($BuiltinAdmin)

                        $aceList = @($AdminFullControlFolderAce)
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'W' -Name $testUser2
                        $aceList.foreach( { $acl.AddAccessRule($_) })

                        $testItem.SetAccessControl($acl)
                        #endregion

                        #region Set correct permissions on a sub folder
                        $testItem = Get-Item "$($Params.Path)\FolderB\SubFolderC"

                        $acl = New-Object System.Security.AccessControl.DirectorySecurity
                        $acl.SetAccessRuleProtection($true, $false)
                        $acl.SetOwner($BuiltinAdmin)

                        $aceList = @($AdminFullControlFolderAce)
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser2
                        $aceList.foreach( { $acl.AddAccessRule($_) })

                        $testItem.SetAccessControl($acl)
                        #endregion

                        $Actual = ."$here\$sut" @Params | Where-Object Name -eq 'Non inherited folder incorrect permissions'
                    
                        $Actual.Value | Should -Be $Params.Path
                    } -Skip:$Skip
                }
                Context 'a file has' {
                    It 'explicit permissions' {
                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'Fix'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L'; $testUser = 'L'; $testUser2 = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser = 'R' } }
                            )
                        }

                        #region Create all folders
                        $Params.Matrix | Select-Object -Skip 1 | ForEach-Object {
                            New-Item -Path (Join-Path $Params.Path $_.Path) -ItemType Directory -Force
                        }
                        #endregion

                        #region Create all files
                        @(, (Get-ChildItem $Params.Path -Recurse -Directory).FullName + $Params.Path) | 
                        ForEach-Object {
                            New-Item -Path (Join-Path $_ 'file') -ItemType File
                        }
                        #endregion

                        #region Set correct permissions on parent folder
                        $testItem = Get-Item $Params.Path

                        $acl = New-Object System.Security.AccessControl.DirectorySecurity
                        $acl.SetAccessRuleProtection($true, $false)
                        $acl.SetOwner($BuiltinAdmin)

                        $aceList = @($AdminFullControlFolderAce)
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'L' -Name $env:USERNAME
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'L' -Name $testUser
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'L' -Name $testUser2
                        $aceList.foreach( { $acl.AddAccessRule($_) })

                        $testItem.SetAccessControl($acl)
                        #endregion

                        #region Set correct permissions on a sub folder
                        $testItem = Get-Item "$($Params.Path)\FolderA"

                        $acl = New-Object System.Security.AccessControl.DirectorySecurity
                        $acl.SetAccessRuleProtection($true, $false)
                        $acl.SetOwner($BuiltinAdmin)

                        $aceList = @($AdminFullControlFolderAce)
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser
                        $aceList.foreach( { $acl.AddAccessRule($_) })

                        $testItem.SetAccessControl($acl)
                        #endregion

                        #region Set incorrect permissions on a file
                        $testItem = Get-Item "$($Params.Path)\FolderA\File"

                        $acl = New-Object System.Security.AccessControl.FileSecurity
                        $acl.SetAccessRuleProtection($true, $false)
                        $acl.SetOwner($BuiltinAdmin)

                        $aceList = @($AdminFullControlIFileAce)
                        $aceList += New-TestAceHC -Type 'InheritedFile' -Access 'W' -Name $testUser2
                        $aceList.foreach( { $acl.AddAccessRule($_) })

                        $testItem.SetAccessControl($acl)
                        #endregion


                        $Actual = ."$here\$sut" @Params

                        $Actual | Where-Object Name -eq 'Inherited permissions incorrect'
                    
                        $Actual.Value | Should -Be "$($Params.Path)\FolderA\File"
                    } -Skip:$Skip
                }
                Context 'a folder that should have inherited permissions' {
                    It 'in the matrix has explicit permissions' {
                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'Fix'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L'; $testUser = 'L'; $testUser2 = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser = 'R' } }
                                [PSCustomObject]@{Path = 'FolderB\SubFolderB'; ACL = @{$testUser2 = 'W' } }
                                [PSCustomObject]@{Path = 'FolderB\SubFolderC'; ACL = @{$testUser2 = 'R' } }
                            )
                        }

                        #region Create all folders
                        $Params.Matrix | Select-Object -Skip 1 | ForEach-Object {
                            New-Item -Path (Join-Path $Params.Path $_.Path) -ItemType Directory -Force
                        }
                        #endregion

                        #region Create all files
                        @(, (Get-ChildItem $Params.Path -Recurse -Directory).FullName + $Params.Path) | 
                        ForEach-Object {
                            New-Item -Path (Join-Path $_ 'file') -ItemType File
                        }
                        #endregion

                        #region Set correct permissions on parent folder
                        $testItem = Get-Item $Params.Path

                        $acl = New-Object System.Security.AccessControl.DirectorySecurity
                        $acl.SetAccessRuleProtection($true, $false)
                        $acl.SetOwner($BuiltinAdmin)

                        $aceList = @($AdminFullControlFolderAce)
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'L' -Name $env:USERNAME
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'L' -Name $testUser
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'L' -Name $testUser2
                        $aceList.foreach( { $acl.AddAccessRule($_) })

                        $testItem.SetAccessControl($acl)
                        #endregion

                        #region Set correct permissions on a sub folder
                        $testItem = Get-Item "$($Params.Path)\FolderA"

                        $acl = New-Object System.Security.AccessControl.DirectorySecurity
                        $acl.SetAccessRuleProtection($true, $false)
                        $acl.SetOwner($BuiltinAdmin)

                        $aceList = @($AdminFullControlFolderAce)
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser
                        $aceList.foreach( { $acl.AddAccessRule($_) })

                        $testItem.SetAccessControl($acl)
                        #endregion

                        #region Set correct permissions on a sub folder
                        $testItem = Get-Item "$($Params.Path)\FolderB\SubFolderB"

                        $acl = New-Object System.Security.AccessControl.DirectorySecurity
                        $acl.SetAccessRuleProtection($true, $false)
                        $acl.SetOwner($BuiltinAdmin)

                        $aceList = @($AdminFullControlFolderAce)
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'W' -Name $testUser2
                        $aceList.foreach( { $acl.AddAccessRule($_) })

                        $testItem.SetAccessControl($acl)
                        #endregion

                        #region Set correct permissions on a sub folder
                        $testItem = Get-Item "$($Params.Path)\FolderB\SubFolderC"

                        $acl = New-Object System.Security.AccessControl.DirectorySecurity
                        $acl.SetAccessRuleProtection($true, $false)
                        $acl.SetOwner($BuiltinAdmin)

                        $aceList = @($AdminFullControlFolderAce)
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser2
                        $aceList.foreach( { $acl.AddAccessRule($_) })

                        $testItem.SetAccessControl($acl)
                        #endregion

                        #region Set incorrect permissions on an inherited folder
                        $testItem = Get-Item "$($Params.Path)\FolderB"

                        $acl = New-Object System.Security.AccessControl.DirectorySecurity
                        $acl.SetAccessRuleProtection($true, $false)
                        $acl.SetOwner($BuiltinAdmin)

                        $aceList = @($AdminFullControlFolderAce)
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser2
                        $aceList.foreach( { $acl.AddAccessRule($_) })

                        $testItem.SetAccessControl($acl)
                        #endregion

                        $Actual = ."$here\$sut" @Params
                    
                        $Actual | Where-Object Name -eq 'Inherited permissions incorrect'
                    
                        $Actual.Value | Should -Be "$($Params.Path)\FolderB"
                    } -Skip:$Skip
                    It 'not defined in the matrix has explicit permissions' {
                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'Fix'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L'; $testUser = 'L'; $testUser2 = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser = 'R' } }
                                [PSCustomObject]@{Path = 'FolderB\SubFolderB'; ACL = @{$testUser2 = 'W' } }
                                [PSCustomObject]@{Path = 'FolderB\SubFolderC'; ACL = @{$testUser2 = 'R' } }
                            )
                        }

                        #region Create all folders
                        $Params.Matrix | Select-Object -Skip 1 | ForEach-Object {
                            New-Item -Path (Join-Path $Params.Path $_.Path) -ItemType Directory -Force
                        }
                        #endregion

                        #region Create all files
                        @(, (Get-ChildItem $Params.Path -Recurse -Directory).FullName + $Params.Path) | 
                        ForEach-Object {
                            New-Item -Path (Join-Path $_ 'file') -ItemType File
                        }
                        #endregion

                        #region Set correct permissions on parent folder
                        $testItem = Get-Item $Params.Path

                        $acl = New-Object System.Security.AccessControl.DirectorySecurity
                        $acl.SetAccessRuleProtection($true, $false)
                        $acl.SetOwner($BuiltinAdmin)

                        $aceList = @($AdminFullControlFolderAce)
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'L' -Name $env:USERNAME
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'L' -Name $testUser
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'L' -Name $testUser2
                        $aceList.foreach( { $acl.AddAccessRule($_) })

                        $testItem.SetAccessControl($acl)
                        #endregion

                        #region Set correct permissions on a sub folder
                        $testItem = Get-Item "$($Params.Path)\FolderA"

                        $acl = New-Object System.Security.AccessControl.DirectorySecurity
                        $acl.SetAccessRuleProtection($true, $false)
                        $acl.SetOwner($BuiltinAdmin)

                        $aceList = @($AdminFullControlFolderAce)
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser
                        $aceList.foreach( { $acl.AddAccessRule($_) })

                        $testItem.SetAccessControl($acl)
                        #endregion

                        #region Set correct permissions on a sub folder
                        $testItem = Get-Item "$($Params.Path)\FolderB\SubFolderB"

                        $acl = New-Object System.Security.AccessControl.DirectorySecurity
                        $acl.SetAccessRuleProtection($true, $false)
                        $acl.SetOwner($BuiltinAdmin)

                        $aceList = @($AdminFullControlFolderAce)
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'W' -Name $testUser2
                        $aceList.foreach( { $acl.AddAccessRule($_) })

                        $testItem.SetAccessControl($acl)
                        #endregion

                        #region Set correct permissions on a sub folder
                        $testItem = Get-Item "$($Params.Path)\FolderB\SubFolderC"

                        $acl = New-Object System.Security.AccessControl.DirectorySecurity
                        $acl.SetAccessRuleProtection($true, $false)
                        $acl.SetOwner($BuiltinAdmin)

                        $aceList = @($AdminFullControlFolderAce)
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser2
                        $aceList.foreach( { $acl.AddAccessRule($_) })

                        $testItem.SetAccessControl($acl)
                        #endregion

                        #region Set incorrect permissions on an inherited folder not in the matrix file

                        $testItem = New-Item -Path "$($Params.Path)\FolderC" -ItemType Directory -Force

                        $acl = New-Object System.Security.AccessControl.DirectorySecurity
                        $acl.SetAccessRuleProtection($true, $false)
                        $acl.SetOwner($BuiltinAdmin)

                        $aceList = @($AdminFullControlFolderAce)
                        $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser2
                        $aceList.foreach( { $acl.AddAccessRule($_) })

                        $testItem.SetAccessControl($acl)
                        #endregion

                        $Actual = ."$here\$sut" @Params
                    
                        $Actual | Where-Object Name -eq 'Inherited permissions incorrect'
                    
                        $Actual.Value | Should -Be "$($Params.Path)\FolderC"
                    } -Skip:$Skip
                }
            }
        }
        Describe 'when Action is' {
            BeforeEach {
                Set-Location $TestDrive
                Remove-Item $TestDrive\* -Recurse -Force
            }
            Context 'New' {
                it "create the parent folder 'Path'" {
                    $Params = @{
                        Path   = Join-Path $testDrive 'testFolder'
                        Action = 'New'
                        Matrix = @([PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true })
                    }

                    ."$here\$sut" @Params

                    $Params.Path | Should -Exist
                } -Skip:$Skip
                it 'create a FatalError object when the parent folder is already present' {
                    $Params = @{
                        Path   = Join-Path $testDrive 'testFolder'
                        Action = 'New'
                        Matrix = [PSCustomObject]@{Name = 'test' }
                    }

                    $null = New-Item -Path $Params.Path -ItemType Directory -EA Ignore

                    $Actual = ."$here\$sut" @Params

                    $Expected = [PSCustomObject]@{
                        Type        = 'FatalError'
                        Name        = 'Parent folder exists already'
                        Description = "The folder defined as 'Path' in the worksheet 'Settings' cannot be present on the remote machine when 'Action=New' is used. Please use 'Action' with value 'Check' or 'Fix' instead."
                        Value       = $Params.Path
                    }

                    $Actual.Type | Should -Be $Expected.Type
                    $Actual.Name | Should -Be $Expected.Name
                    $Actual.Description | Should -Be $Expected.Description
                    $Actual.Value | Should -Be $Expected.Value
                } -Skip:$Skip
                Context 'folders in the matrix that need to be created' {
                    it 'are created' {
                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'New'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                [PSCustomObject]@{Path = 'FolderB\FolderC'; ACL = @{ }}
                            )
                        }

                        ."$here\$sut" @Params

                        $Params.Path | Should -Exist
                        $Params.Path + '\FolderA' | Should -Exist
                        $Params.Path + '\FolderB\FolderC' | Should -Exist
                    } -Skip:$Skip
                    it 'are registered in a Warning object' {
                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'New'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                [PSCustomObject]@{Path = 'FolderB\FolderC'; ACL = @{ }}
                            )
                        }

                        $Actual = ."$here\$sut" @Params | Where-Object Name -Like "*child folder*"

                        $Actual.Type | Should -Be 'Warning'
                        $Actual.Name | Should -Be 'Child folder created'

                        @(
                            "$($Params.Path)",
                            "$($Params.Path)\FolderA",
                            "$($Params.Path)\FolderB\FolderC"
                        ).ForEach( {
                                $Actual.Value | Should -Contain $_
                            })
                        $actual.Value.Count | Should -BeExactly 3
                    } -Skip:$Skip
                    it 'are not created when Path is set to Ignore' {
                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'New'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }; Ignore = $true }
                                [PSCustomObject]@{Path = 'FolderB\FolderC'; ACL = @{ }}
                            )
                        }

                        ."$here\$sut" @Params

                        "$($Params.Path)\FolderA" | Should -Not -Exist
                        "$($Params.Path)\FolderB\FolderC" | Should -Exist
                    } -Skip:$Skip
                }
                Context 'set permissions' {
                    it 'on the parent folder' {
                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'New'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                            )
                        }

                        ."$here\$sut" @Params

                        $Actual = (Get-Acl -Path $Params.Path).Access

                        $Actual.Count | Should -BeExactly 2 -Because "ACL is 'BUILTIN\Administrators' and '$testUser'."
                        $Actual[0].IdentityReference | Should -Be 'BUILTIN\Administrators'
                        $Actual[1].IdentityReference | Should -Be "$env:USERDOMAIN\$testUser"
                    } -Skip:$Skip
                    it 'on the child folders' {
                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'New'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' }}
                            )
                        }

                        ."$here\$sut" @Params

                        $Actual = (Get-Acl -Path $Params.Path).Access
                        $Actual.Count | Should -BeExactly 2
                        $Actual[0].IdentityReference | Should -Be 'BUILTIN\Administrators'
                        $Actual[1].IdentityReference | Should -Be "$env:USERDOMAIN\$testUser"

                        $Actual = (Get-Acl -Path "$($Params.Path)\FolderB").Access
                        $Actual.Count | Should -BeExactly 2
                        $Actual[0].IdentityReference | Should -Be 'BUILTIN\Administrators'
                        $Actual[1].IdentityReference | Should -Be "$env:USERDOMAIN\$testUser2"
                    } -Skip:$Skip
                    it 'a Warning object for incorrect permissions is not created' {
                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'New'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                [PSCustomObject]@{Path = 'FolderB\FolderC'; ACL = @{ }}
                            )
                        }

                        ."$here\$sut" @Params | Where-Object Name -EQ 'Non inherited folder incorrect permissions' |
                        Should -BeNullOrEmpty
                    } -Skip:$Skip
                }
            }
            Context 'Fix' {
                it "create a FatalError object when the parent folder doesn't exist" {
                    $Params = @{
                        Path   = 'NotExistingTestFolder'
                        Action = 'Fix'
                        Matrix = [PSCustomObject]@{Name = 'test' }
                    }

                    $Actual = ."$here\$sut" @Params

                    $Expected = [PSCustomObject]@{
                        Type        = 'FatalError'
                        Name        = 'Parent folder missing'
                        Description = "The folder defined as 'Path' in the worksheet 'Settings' needs to be available on the remote machine. In case the folder structure needs to be created, please use 'Action=New' instead."
                        Value       = $Params.Path
                    }

                    $Actual.Type | Should -Be $Expected.Type
                    $Actual.Name | Should -Be $Expected.Name
                    $Actual.Description | Should -Be $Expected.Description
                    $Actual.Value | Should -Be $Expected.Value
                } -Skip:$Skip
                Context 'folders in the matrix that are missing' {
                    it 'are created' {
                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'Fix'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                [PSCustomObject]@{Path = 'FolderB\FolderC'; ACL = @{ }}
                            )
                        }
                        $null = New-Item -Path $Params.Path -ItemType Directory

                        ."$here\$sut" @Params

                        "$($Params.Path)\FolderA" | Should -Exist
                        "$($Params.Path)\FolderB\FolderC" | Should -Exist
                    } -Skip:$Skip
                    it 'are registered in a Warning object' {
                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'Fix'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                [PSCustomObject]@{Path = 'FolderB\FolderC'; ACL = @{ }}
                            )
                        }
                        $null = New-Item -Path $Params.Path -ItemType Directory

                        $Actual = ."$here\$sut" @Params | Where-Object Name -Like "*child folder*"

                        $Actual.Type | Should -Be 'Warning'
                        $Actual.Name | Should -Be 'Child folder created'
                        
                        $Actual.Value[0] | Should -Be "$($Params.Path)\FolderA"
                        $Actual.Value[1] | Should -Be "$($Params.Path)\FolderB\FolderC"
                    } -Skip:$Skip
                    it 'are not created when Path is set to Ignore' {
                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'Fix'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }; Ignore = $true }
                                [PSCustomObject]@{Path = 'FolderB\FolderC'; ACL = @{ }}
                            )
                        }
                        New-Item -Path $Params.Path -ItemType Directory

                        ."$here\$sut" @Params
                        
                        "$($Params.Path)\FolderA" | Should -Not -Exist
                        "$($Params.Path)\FolderB\FolderC" | Should -Exist
                    } -Skip:$Skip
                }
                Context 'incorrect folder permissions' {
                    Context 'on non inherited folders' {
                        it 'are corrected' {
                            $Params = @{
                                Path   = Join-Path $testDrive 'testFolder'
                                Action = 'New'
                                Matrix = @(
                                    [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                    [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser2 = 'R' }}
                                )
                            }

                            ."$here\$sut" @Params

                            $Params = @{
                                Path   = Join-Path $testDrive 'testFolder'
                                Action = 'Fix'
                                Matrix = @(
                                    [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser2 = 'R' }; Parent = $true }
                                    [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser = 'L' }}
                                )
                            }
                            Set-Location $TestDrive
                            ."$here\$sut" @Params

                            $Actual = (Get-Acl -Path $Params.Path).Access
                            $Actual.Count | Should -BeExactly 2
                            $Actual[0].IdentityReference | Should -Be 'BUILTIN\Administrators'
                            $Actual[1].IdentityReference | Should -Be "$env:USERDOMAIN\$testUser2"

                            $Actual = (Get-Acl -Path ($Params.Path + '\FolderA')).Access
                            $Actual.Count | Should -BeExactly 2
                            $Actual[0].IdentityReference | Should -Be 'BUILTIN\Administrators'
                            $Actual[1].IdentityReference | Should -Be "$env:USERDOMAIN\$testUser"
                        } -Skip:$Skip
                        Context 'are registered in a Warning object when' {
                            it 'DetailedLog is False only the folder name is saved' {
                                $Params = @{
                                    Path   = Join-Path $testDrive 'testFolder'
                                    Action = 'Fix'
                                    Matrix = @(
                                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser2 = 'R' }}
                                    )
                                }

                                New-Item -Path 'testFolder\FolderA' -ItemType Directory -Force

                                $Actual = ."$here\$sut" @Params |
                                Where-Object Name -EQ $ExpectedIncorrectAclNonInheritedFolders.Name

                                $Actual.Type | Should -Be $ExpectedIncorrectAclNonInheritedFolders.Type
                                @(
                                    $Params.Path,
                                    "$($Params.Path)\FolderA"
                                ).ForEach( {
                                        $Actual.Value | Should -Contain $_
                                    })
                            } -Skip:$Skip
                            it 'DetailedLog is True the folder name, the old ACL and the new ACL are saved' {
                                $Params = @{
                                    Path        = Join-Path $testDrive 'testFolder'
                                    Action      = 'Fix'
                                    Matrix      = @(
                                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser2 = 'R' }}
                                    )
                                    DetailedLog = $true
                                }

                                New-Item -Path 'testFolder\FolderA' -ItemType Directory -Force

                                $Actual = ."$here\$sut" @Params |
                                Where-Object Name -EQ $ExpectedIncorrectAclNonInheritedFolders.Name

                                $Actual.Type | Should -Be $ExpectedIncorrectAclNonInheritedFolders.Type
                                $Actual.Value.Count | Should -BeExactly 2 -Because 'two folders have an incorrect ACL'

                                @(
                                    $Params.Path,
                                    "$($Params.Path)\FolderA"
                                ).ForEach( {
                                        $Actual.Value.Keys | Should -Contain $_ -Because 'the folder FullName is expected'
                                    })

                                $Actual.Value.GetEnumerator().ForEach( {
                                        foreach ($v in @('old', 'new')) {
                                            $_.Value.$v | Should -Not -BeNullOrEmpty -Because 'an ACL is expected'
                                        }
                                    })
                            } -Skip:$Skip
                        }
                    }
                    Context 'on inherited folders' {
                        it 'are corrected' {
                            $Params = @{
                                Path   = Join-Path $testDrive 'testFolder'
                                Action = 'New'
                                Matrix = @(
                                    [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                    [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser2 = 'R' }}
                                )
                            }

                            ."$here\$sut" @Params

                            $Params = @{
                                Path   = Join-Path $testDrive 'testFolder'
                                Action = 'Fix'
                                Matrix = @(
                                    [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                    [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                )
                            }
                            Set-Location $TestDrive
                            ."$here\$sut" @Params

                            $Actual = (Get-Acl -Path "$($Params.Path)\FolderA").Access
                            $Actual.IsInherited | Should -Not -Contain $false -Because 'IsInedited needs to be True on all Ace'
                        } -Skip:$Skip
                        Context 'are registered in a Warning object when' {
                            it 'DetailedLog is False only the folder name is saved' {
                                $Params = @{
                                    Path   = Join-Path $testDrive 'testFolder'
                                    Action = 'New'
                                    Matrix = @(
                                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser2 = 'R' }}
                                        [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' }}
                                    )
                                }

                                ."$here\$sut" @Params

                                $Params = @{
                                    Path   = Join-Path $testDrive 'testFolder'
                                    Action = 'Fix'
                                    Matrix = @(
                                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                    )
                                }
                                Set-Location $TestDrive
                                $Actual = ."$here\$sut" @Params |
                                Where-Object Name -EQ $ExpectedIncorrectAclInheritedFolders.Name

                                $Actual.Type | Should -Be $ExpectedIncorrectAclInheritedFolders.Type
                                @(
                                    ($TestDrive.FullName + '\testFolder\FolderA')
                                    ($TestDrive.FullName + '\testFolder\FolderB')
                                ).ForEach( {
                                        $Actual.Value | Should -Contain $_
                                    })

                            } -Skip:$Skip
                            it 'DetailedLog is True the folder name, the old ACL and the new ACL are saved' {
                                $Params = @{
                                    Path   = Join-Path $testDrive 'testFolder'
                                    Action = 'New'
                                    Matrix = @(
                                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser2 = 'R' }}
                                        [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' }}
                                    )
                                }

                                ."$here\$sut" @Params

                                $Params = @{
                                    Path        = Join-Path $testDrive 'testFolder'
                                    Action      = 'Fix'
                                    Matrix      = @(
                                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                    )
                                    DetailedLog = $true
                                }
                                Set-Location $TestDrive
                                $Actual = ."$here\$sut" @Params |
                                Where-Object Name -EQ $ExpectedIncorrectAclInheritedFolders.Name

                                $Actual.Type | Should -Be $ExpectedIncorrectAclInheritedFolders.Type
                                @(
                                    ($TestDrive.FullName + '\testFolder\FolderA')
                                    ($TestDrive.FullName + '\testFolder\FolderB')
                                ).ForEach( {
                                        $Actual.Value.Keys | Should -Contain $_
                                    })

                                $Actual.Value.GetEnumerator().ForEach( {
                                        $_.Value | Should -Not -BeNullOrEmpty -Because 'an ACL is expected'
                                    })
                            } -Skip:$Skip
                        }
                    }
                    Context "set the owner to 'BUILTIN\Administrators' when" {
                        it 'the admin has access to all folders' {
                            $Params = @{
                                Path   = Join-Path $testDrive 'testFolder'
                                Action = 'Fix'
                                Matrix = @(
                                    [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                    [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                    [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' }}
                                )
                            }

                            New-Item -Path 'testFolder\FolderA' -ItemType Directory -Force
                            $testFolderPath = 'testFolder\FolderB'
                            $testFolder = New-Item -Path $testFolderPath -ItemType Directory -Force

                            #region Add ourselves as owner
                            $testOwner = [System.Security.Principal.NTAccount]"$env:USERDOMAIN\$env:USERNAME"
                            $testAcl = $testFolder.GetAccessControl()
                            $testAcl.SetOwner($testOwner)
                            $testFolder.SetAccessControl($testAcl)
                            #endregion

                            (Get-Acl -Path $testFolderPath).Owner | Should -Be "$env:USERDOMAIN\$env:USERNAME"

                            ."$here\$sut" @Params
                            Set-Location -Path $TestDrive

                            (Get-Acl -Path $testFolderPath).Owner | Should -Be 'BUILTIN\Administrators'
                        } -Skip:$Skip
                        it 'the admin has no access to the folder' {
                            $Params = @{
                                Path   = Join-Path $testDrive 'testFolder'
                                Action = 'Fix'
                                Matrix = @(
                                    [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                    [PSCustomObject]@{Path = 'Reports'; ACL = @{ }}
                                    [PSCustomObject]@{Path = 'Reports\Fruits'; ACL = @{ }}
                                    [PSCustomObject]@{Path = 'Reports\Fruits\Kiwi'; ACL = @{$testUser = 'R' }}
                                )
                            }
                            $testFolderPath = 'testFolder\Reports\Fruits\Kiwi'
                            $testFolder = New-Item -Path $testFolderPath -ItemType Directory -Force

                            #region Remove access
                            $testOwner = [System.Security.Principal.NTAccount]"$env:USERDOMAIN\$testUser"
                            $testAcl = $testFolder.GetAccessControl()
                            $testAcl.SetAccessRuleProtection($True, $False)
                            $testAcl.SetOwner($testOwner)
                            $testAcl.Access.ForEach( { $null = $testAcl.RemoveAccessRule($_) })
                            $testFolder.SetAccessControl($testAcl)
                            #endregion

                            (Get-Acl -Path $testFolderPath).Owner | Should -Be "$env:USERDOMAIN\$testUser"
                            (Get-Acl -Path $testFolderPath).Access | Should -BeNullOrEmpty

                            ."$here\$sut" @Params
                            Set-Location -Path $TestDrive

                            (Get-Acl -Path $testFolderPath).Owner | Should -Be 'BUILTIN\Administrators'
                            (Get-Acl -Path $testFolderPath).Access | Should -Not -BeNullOrEmpty
                        } -Skip:$Skip
                        it 'the admin has no access to the parent folder' {
                            $Params = @{
                                Path   = Join-Path $testDrive 'testFolder'
                                Action = 'Fix'
                                Matrix = @(
                                    [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                    [PSCustomObject]@{Path = 'Reports'; ACL = @{ }}
                                    [PSCustomObject]@{Path = 'Reports\Fruits'; ACL = @{ }}
                                    [PSCustomObject]@{Path = 'Reports\Fruits\Kiwi'; ACL = @{$testUser = 'R' }}
                                )
                            }
                            New-Item -Path 'testFolder\Reports\Fruits\Kiwi' -ItemType Directory -Force

                            $testFolderPath = 'testFolder\Reports\Fruits'
                            $testFolder = New-Item -Path $testFolderPath -ItemType Directory -Force

                            #region Remove access and set owner
                            $testOwner = [System.Security.Principal.NTAccount]"$env:USERDOMAIN\$testUser"
                            $testAcl = $testFolder.GetAccessControl()
                            $testAcl.SetAccessRuleProtection($True, $False)
                            $testAcl.SetOwner($testOwner)
                            $testAcl.Access.ForEach( { $null = $testAcl.RemoveAccessRule($_) })
                            $testFolder.SetAccessControl($testAcl)
                            #endregion

                            (Get-Acl -Path $testFolderPath).Owner | Should -Be "$env:USERDOMAIN\$testUser"
                            (Get-Acl -Path $testFolderPath).Access | Should -BeNullOrEmpty
                            (Get-Acl -Path 'testFolder\Reports\Fruits\Kiwi').Access.IdentityReference |
                            Should -Not -Contain "$env:USERDOMAIN\$testUser"

                            ."$here\$sut" @Params
                            Set-Location -Path $TestDrive

                            (Get-Acl -Path $testFolderPath).Owner | Should -Be 'BUILTIN\Administrators'
                            (Get-Acl -Path 'testFolder\Reports\Fruits\Kiwi').Owner | Should -Be 'BUILTIN\Administrators'
                            (Get-Acl -Path 'testFolder\Reports\Fruits\Kiwi').Access.IdentityReference |
                            Should -Contain "$env:USERDOMAIN\$testUser"
                        } -Skip:$Skip
                    }
                }
                Context 'when the script is run again after Action Fix/New' {
                    it 'the permissions are unchanged' {
                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'New'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' }}
                            )
                        }

                        $testPermissions = {
                            $Actual = (Get-Acl -Path $Params.Path).Access
                            $Actual.Count | Should -BeExactly 2
                            $Actual[0].IdentityReference | Should -Be 'BUILTIN\Administrators'
                            $Actual[1].IdentityReference | Should -Be "$env:USERDOMAIN\$testUser"

                            $Actual = (Get-Acl -Path "$($Params.Path)\FolderB").Access
                            $Actual.Count | Should -BeExactly 2
                            $Actual[0].IdentityReference | Should -Be 'BUILTIN\Administrators'
                            $Actual[1].IdentityReference | Should -Be "$env:USERDOMAIN\$testUser2"
                        }

                        ."$here\$sut" @Params

                        & $testPermissions

                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'Fix'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' }}
                            )
                        }
                        Set-Location $TestDrive
                        ."$here\$sut" @Params

                        & $testPermissions
                    } -Skip:$Skip
                    it 'nothing is reported as being incorrect' {
                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'New'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' }}
                            )
                        }

                        ."$here\$sut" @Params

                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'Fix'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' }}
                            )
                        }
                        Set-Location $TestDrive
                        ."$here\$sut" @Params | Where-Object { $_.Type -notmatch 'Information|Warning' } | Should -BeNullOrEmpty
                    } -Skip:$Skip
                }
            }
            Context 'Check' {
                it "create a FatalError object when the parent folder doesn't exist" {
                    $Params = @{
                        Path   = 'NotExistingTestFolder'
                        Action = 'Check'
                        Matrix = [PSCustomObject]@{Name = 'test' }
                    }

                    $Actual = ."$here\$sut" @Params

                    $Expected = [PSCustomObject]@{
                        Type        = 'FatalError'
                        Name        = 'Parent folder missing'
                        Description = "The folder defined as 'Path' in the worksheet 'Settings' needs to be available on the remote machine. In case the folder structure needs to be created, please use 'Action=New' instead."
                        Value       = $Params.Path
                    }

                    $Actual.Type | Should -Be $Expected.Type
                    $Actual.Name | Should -Be $Expected.Name
                    $Actual.Description | Should -Be $Expected.Description
                    $Actual.Value | Should -Be $Expected.Value
                } -Skip:$Skip
                Context 'folders in the matrix that are missing' {
                    it 'are not created' {
                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'Check'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                [PSCustomObject]@{Path = 'FolderB\FolderC'; ACL = @{ }}
                            )
                        }
                        $null = New-Item -Path $Params.Path -ItemType Directory

                        ."$here\$sut" @Params

                        $Params.Path + '\' + $Params.Matrix[1].Path | Should -Not -Exist
                        $Params.Path + '\' + $Params.Matrix[2].Path | Should -Not -Exist
                    } -Skip:$Skip
                    it 'are registered in a Warning object' {
                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'Check'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                [PSCustomObject]@{Path = 'FolderB\FolderC'; ACL = @{ }}
                            )
                        }
                        New-Item -Path $Params.Path -ItemType Directory

                        $Actual = ."$here\$sut" @Params | Where-Object Name -Like "*child folder*"

                        $Actual.Type | Should -Be 'Warning'
                        $Actual.Name | Should -Be 'Child folder missing'
                        $Actual.Value[0] | Should -Be "$($Params.Path)\FolderA"
                        $Actual.Value[1] | Should -Be "$($Params.Path)\FolderB\FolderC"
                    } -Skip:$Skip
                    it 'are not checked when they are set to ignore' {
                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'Check'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }; Ignore = $true }
                                [PSCustomObject]@{Path = 'FolderB\FolderC'; ACL = @{ }}
                            )
                        }
                        New-Item -Path $Params.Path -ItemType Directory

                        ."$here\$sut" @Params | Where-Object { ($_.Name -Like "*child folder*") -and
                            ($_.Value -contains ($TestDrive.FullName + '\testFolder\FolderA')) } |
                        Should -BeNullOrEmpty

                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'Check'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                [PSCustomObject]@{Path = 'FolderB\FolderC'; ACL = @{ }}
                            )
                        }
                        Set-Location -Path $TestDrive
                        ."$here\$sut" @Params | Where-Object { ($_.Name -Like "*child folder*") -and
                            ($_.Value -contains ($TestDrive.FullName + '\testFolder\FolderA')) } |
                        Should -Not -BeNullOrEmpty

                    } -Skip:$Skip
                }
                Context 'incorrect folder permissions' {
                    Context 'on non inherited folders' {
                        it 'are not corrected' {
                            $Params = @{
                                Path   = Join-Path $testDrive 'testFolder'
                                Action = 'Check'
                                Matrix = @(
                                    [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                    [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                    [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' }}
                                )
                            }

                            New-Item -Path 'testFolder\FolderA' -ItemType Directory -Force
                            New-Item -Path 'testFolder\FolderB' -ItemType Directory -Force

                            $Expected = @(
                                (Get-Acl -Path $Params.Path)
                                (Get-Acl -Path "$($Params.Path)\FolderA")
                                (Get-Acl -Path "$($Params.Path)\FolderB")
                            )

                            ."$here\$sut" @Params

                            $Actual = @(
                                (Get-Acl -Path $Params.Path)
                                (Get-Acl -Path "$($Params.Path)\FolderA")
                                (Get-Acl -Path "$($Params.Path)\FolderB")
                            )

                            for ($i = 0; $i -lt $Expected.Count; $i++) {
                                $AssertParams = @{
                                    Actual   = $Actual[$i].Access | Sort-Object IdentityReference
                                    Expected = $Expected[$i].Access | Sort-Object IdentityReference
                                }
                                Assert-Equivalent @AssertParams

                                $AssertParams = @{
                                    Actual   = $Actual[$i].Owner | Sort-Object IdentityReference
                                    Expected = $Expected[$i].Owner | Sort-Object IdentityReference
                                }
                                Assert-Equivalent @AssertParams
                            }
                        } -Skip:$Skip
                        Context 'are registered in a Warning object when' {
                            it 'DetailedLog is False only the folder name is saved' {
                                $Params = @{
                                    Path   = Join-Path $testDrive 'testFolder'
                                    Action = 'Check'
                                    Matrix = @(
                                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser2 = 'R' }}
                                    )
                                }

                                New-Item -Path 'testFolder\FolderA' -ItemType Directory -Force

                                $Actual = ."$here\$sut" @Params |
                                Where-Object Name -EQ $ExpectedIncorrectAclNonInheritedFolders.Name

                                $Actual.Type | Should -Be $ExpectedIncorrectAclNonInheritedFolders.Type
                                @(
                                    ($TestDrive.FullName + '\testFolder'),
                                    ($TestDrive.FullName + '\testFolder\FolderA')
                                ).ForEach( {
                                        $Actual.Value | Should -Contain $_
                                    })

                            } -Skip:$Skip
                            it 'DetailedLog is True the folder name, the old ACL and the new ACL are saved' {
                                $Params = @{
                                    Path        = Join-Path $testDrive 'testFolder'
                                    Action      = 'Check'
                                    Matrix      = @(
                                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser2 = 'R' }}
                                    )
                                    DetailedLog = $true
                                }

                                New-Item -Path 'testFolder\FolderA' -ItemType Directory -Force

                                $Actual = ."$here\$sut" @Params |
                                Where-Object Name -EQ $ExpectedIncorrectAclNonInheritedFolders.Name

                                $Actual.Type | Should -Be $ExpectedIncorrectAclNonInheritedFolders.Type
                                $Actual.Value.Count | Should -BeExactly 2 -Because 'two folders have an incorrect ACL'

                                @(
                                    ($TestDrive.FullName + '\testFolder'),
                                    ($TestDrive.FullName + '\testFolder\FolderA')
                                ).ForEach( {
                                        $Actual.Value.Keys | Should -Contain $_ -Because 'the folder FullName is expected'
                                    })

                                $Actual.Value.GetEnumerator().ForEach( {
                                        foreach ($v in @('old', 'new')) {
                                            $_.Value.$v | Should -Not -BeNullOrEmpty -Because 'an ACL is expected'
                                        }
                                    })

                            } -Skip:$Skip
                        }
                    }
                    Context 'on inherited folders' {
                        it 'are not corrected' {
                            $Params = @{
                                Path   = Join-Path $testDrive 'testFolder'
                                Action = 'New'
                                Matrix = @(
                                    [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                    [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser2 = 'R' }}
                                )
                            }

                            ."$here\$sut" @Params

                            $Expected = @(
                                (Get-Acl -Path ($TestDrive.FullName + '\testFolder'))
                                (Get-Acl -Path ($TestDrive.FullName + '\testFolder\FolderA'))
                            )

                            $Params = @{
                                Path   = Join-Path $testDrive 'testFolder'
                                Action = 'Check'
                                Matrix = @(
                                    [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                    [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                )
                            }
                            Set-Location $TestDrive
                            ."$here\$sut" @Params

                            $Actual = @(
                                (Get-Acl -Path ($TestDrive.FullName + '\testFolder'))
                                (Get-Acl -Path ($TestDrive.FullName + '\testFolder\FolderA'))
                            )

                            for ($i = 0; $i -lt $Expected.Count; $i++) {
                                $AssertParams = @{
                                    Actual   = $Actual[$i].Access | Sort-Object IdentityReference
                                    Expected = $Expected[$i].Access | Sort-Object IdentityReference
                                }
                                Assert-Equivalent @AssertParams

                                $AssertParams = @{
                                    Actual   = $Actual[$i].Owner | Sort-Object IdentityReference
                                    Expected = $Expected[$i].Owner | Sort-Object IdentityReference
                                }
                                Assert-Equivalent @AssertParams
                            }
                        } -Skip:$Skip
                        Context 'are registered in a Warning object when' {
                            it 'DetailedLog is False only the folder name is saved' {
                                $Params = @{
                                    Path   = Join-Path $testDrive 'testFolder'
                                    Action = 'New'
                                    Matrix = @(
                                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser2 = 'R' }}
                                        [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' }}
                                    )
                                }

                                ."$here\$sut" @Params

                                $Params = @{
                                    Path   = Join-Path $testDrive 'testFolder'
                                    Action = 'Check'
                                    Matrix = @(
                                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                    )
                                }
                                Set-Location $TestDrive
                                $Actual = ."$here\$sut" @Params |
                                Where-Object Name -EQ $ExpectedIncorrectAclInheritedFolders.Name

                                $Actual.Type | Should -Be $ExpectedIncorrectAclInheritedFolders.Type
                                @(
                                    ($TestDrive.FullName + '\testFolder\FolderA')
                                    ($TestDrive.FullName + '\testFolder\FolderB')
                                ).ForEach( {
                                        $Actual.Value | Should -Contain $_
                                    })
                            } -Skip:$Skip
                            it 'DetailedLog is True the folder name, the old ACL and the new ACL are saved' {
                                $Params = @{
                                    Path   = Join-Path $testDrive 'testFolder'
                                    Action = 'New'
                                    Matrix = @(
                                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser2 = 'R' }}
                                        [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' }}
                                    )
                                }

                                ."$here\$sut" @Params

                                $Params = @{
                                    Path        = Join-Path $testDrive 'testFolder'
                                    Action      = 'Check'
                                    Matrix      = @(
                                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                    )
                                    DetailedLog = $true
                                }
                                Set-Location $TestDrive
                                $Actual = ."$here\$sut" @Params |
                                Where-Object Name -EQ $ExpectedIncorrectAclInheritedFolders.Name

                                $Actual.Type | Should -Be $ExpectedIncorrectAclInheritedFolders.Type
                                @(
                                    ($TestDrive.FullName + '\testFolder\FolderA')
                                    ($TestDrive.FullName + '\testFolder\FolderB')
                                ).ForEach( {
                                        $Actual.Value.Keys | Should -Contain $_
                                    })

                                $Actual.Value.GetEnumerator().ForEach( {
                                        $_.Value | Should -Not -BeNullOrEmpty -Because 'an ACL is expected'
                                    })
                            } -Skip:$Skip
                        }
                    }
                    it "don't report missing folders as having incorrect permissions" {
                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'Check'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' }}
                                [PSCustomObject]@{Path = 'FolderC'; ACL = @{ }}
                            )
                        }

                        New-Item -Path ($Params.Path + '\FolderA') -ItemTyp Directory -Force

                        $actual = ."$here\$sut" @Params | Where-Object {
                            ($_.Name -EQ $ExpectedIncorrectAclNonInheritedFolders.Name) -or 
                            ($_.Name -EQ $ExpectedIncorrectAclInheritedFolders.Name)
                        }

                        @(
                            ($Params.Path + '\FolderB')
                            ($Params.Path + '\FolderC')
                        ).ForEach( {
                                $actual.Value | Should -Not -Contain $_
                            })
                    } -Skip:$Skip
                }
                Context 'incorrect file permissions' {
                    it 'are not corrected' {
                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'Check'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' }}
                            )
                        }

                        $testFile = New-Item -Path ($Params.Path + '\FolderB\File.txt') -ItemTyp File -Force
                        $testFileItem = Get-Item $testFile

                        #region Remove access and set owner
                        $testOwner = [System.Security.Principal.NTAccount]"$env:USERDOMAIN\$testUser"
                        $testAcl = $testFileItem.GetAccessControl()
                        $testAcl.SetAccessRuleProtection($True, $False)
                        $testAcl.SetOwner($testOwner)
                        $testAcl.Access.ForEach( { $null = $testAcl.RemoveAccessRule($_) })
                        $testFileItem.SetAccessControl($testAcl)
                        #endregion

                        $Expected = @(
                            Get-Acl -Path $testFile.FullName
                        )

                        ."$here\$sut" @Params

                        $Actual = @(
                            Get-Acl -Path $testFile.FullName
                        )

                        for ($i = 0; $i -lt $Expected.Count; $i++) {
                            $AssertParams = @{
                                Actual   = $Actual[$i].Access | Sort-Object IdentityReference
                                Expected = $Expected[$i].Access | Sort-Object IdentityReference
                            }
                            Assert-Equivalent @AssertParams

                            $AssertParams = @{
                                Actual   = $Actual[$i].Owner | Sort-Object IdentityReference
                                Expected = $Expected[$i].Owner | Sort-Object IdentityReference
                            }
                            Assert-Equivalent @AssertParams
                        }
                    } -Skip:$Skip
                    Context 'are registered in a Warning object when' {
                        $Expected = [PSCustomObject]@{
                            Type = 'Warning'
                            Name = 'Inherited permissions incorrect'
                        }
                        it 'DetailedLog is False only the file name is saved' {
                            $Params = @{
                                Path   = Join-Path $testDrive 'testFolder'
                                Action = 'Check'
                                Matrix = @(
                                    [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                    [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                    [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' }}
                                )
                            }

                            $testFile = New-Item -Path ($TestDrive.FullName + '\testFolder\FolderB\File.txt') -ItemTyp File -Force
                            $testFileItem = Get-Item $testFile

                            #region Remove access and set owner
                            $testOwner = [System.Security.Principal.NTAccount]"$env:USERDOMAIN\$testUser"
                            $testAcl = $testFileItem.GetAccessControl()
                            $testAcl.SetAccessRuleProtection($True, $False)
                            $testAcl.SetOwner($testOwner)
                            $testAcl.Access.ForEach( { $null = $testAcl.RemoveAccessRule($_) })
                            $testFileItem.SetAccessControl($testAcl)
                            #endregion

                            $Actual = ."$here\$sut" @Params | Where-Object Name -EQ $Expected.Name

                            @(
                                $testFile.FullName
                            ).ForEach( {
                                    $Actual.Value | Should -Contain $_
                                })
                        } -Skip:$Skip
                        it 'DetailedLog is True the file name and the the old ACL are saved' {
                            $Params = @{
                                Path        = Join-Path $testDrive 'testFolder'
                                Action      = 'Check'
                                Matrix      = @(
                                    [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                    [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                    [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' }}
                                )
                                DetailedLog = $true
                            }

                            $testFile = New-Item -Path ($Params.Path + '\FolderB\File.txt') -ItemTyp File -Force
                            $testFileItem = Get-Item $testFile

                            #region Add explicit access
                            $testAcl = $testFileItem.GetAccessControl()
                            $testAcl.SetAccessRuleProtection($True, $False)
                            $testAcl.Access.ForEach( { $null = $testAcl.RemoveAccessRule($_) })

                            $testAce = New-Object System.Security.AccessControl.FileSystemAccessRule(
                                "$env:USERDOMAIN\$testUser",
                                [System.Security.AccessControl.FileSystemRights]'DeleteSubdirectoriesAndFiles, Modify, Synchronize',
                                [System.Security.AccessControl.InheritanceFlags]::None,
                                [System.Security.AccessControl.PropagationFlags]::None,
                                [System.Security.AccessControl.AccessControlType]::Allow
                            )
                            $testAcl.AddAccessRule($testAce)
                            $testFileItem.SetAccessControl($testAcl)
                            #endregion

                            $Actual = ."$here\$sut" @Params | Where-Object Name -EQ $Expected.Name

                            $Actual.Type | Should -Be $Expected.Type
                            $Actual.Value.Count | Should -BeExactly 1 -Because 'one file is not having inheritance set'

                            @(
                                $testFile.FullName
                            ).ForEach( {
                                    $Actual.Value.Keys | Should -Contain $_ -Because 'the file FullName is expected'
                                })

                            $Actual.Value.GetEnumerator().ForEach( {
                                    $_.Value | Should -Not -BeNullOrEmpty -Because 'an ACL is expected'
                                })
                        } -Skip:$Skip
                    }
                }
                Context 'when the script is run again after Action Fix/New' {
                    it 'the permissions are unchanged' {
                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'New'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' }}
                            )
                        }

                        $testPermissions = {
                            $Actual = (Get-Acl -Path $Params.Path).Access
                            $Actual.Count | Should -BeExactly 2
                            $Actual[0].IdentityReference | Should -Be 'BUILTIN\Administrators'
                            $Actual[1].IdentityReference | Should -Be "$env:USERDOMAIN\$testUser"

                            $Actual = (Get-Acl -Path "$($Params.Path)\FolderB").Access
                            $Actual.Count | Should -BeExactly 2
                            $Actual[0].IdentityReference | Should -Be 'BUILTIN\Administrators'
                            $Actual[1].IdentityReference | Should -Be "$env:USERDOMAIN\$testUser2"
                        }

                        ."$here\$sut" @Params

                        & $testPermissions

                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'Fix'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' }}
                            )
                        }
                        
                        ."$here\$sut" @Params

                        & $testPermissions

                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'Check'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' }}
                            )
                        }
                        
                        ."$here\$sut" @Params

                        & $testPermissions
                    } -Skip:$Skip
                    it 'nothing is reported as being incorrect' {
                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'New'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' }}
                            )
                        }

                        ."$here\$sut" @Params

                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'Fix'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' }}
                            )
                        }
                        Set-Location $TestDrive
                        ."$here\$sut" @Params | Where-Object { $_.Type -notmatch 'Information|Warning' } | Should -BeNullOrEmpty

                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'Check'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' }}
                            )
                        }
                        Set-Location $TestDrive
                        ."$here\$sut" @Params | Where-Object { $_.Type -notmatch 'Information|Warning' } | Should -BeNullOrEmpty
                    } -Skip:$Skip
                }
                Context 'create a Warning object for inaccessible data when' {
                    it 'files are found in the deepest folder of a matrix list only path' {
                        $Params = @{
                            Path   = Join-Path $testDrive 'testFolder'
                            Action = 'Check'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }}
                                [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' }}
                                [PSCustomObject]@{Path = 'FolderB\SubfolderB1'; ACL = @{$testUser2 = 'L' }}
                                [PSCustomObject]@{Path = 'FolderB\SubfolderB2'; ACL = @{$testUser2 = 'W' }}
                            )
                        }

                        $TestFile = New-Item -Path ($Params.Path + '\FolderB\SubfolderB1\File.txt') -ItemTyp File -Force

                        $Actual = ."$here\$sut" @Params | Where-Object Name -EQ $ExpectedInaccessibleData.Name

                        @(
                            $TestFile.FullName
                        ).ForEach( {
                                $Actual.Value | Should -Contain $_ -Because 'the deepest folder has only list permissions'
                            })
                    } -Skip
                }
            }
        }
    }
}