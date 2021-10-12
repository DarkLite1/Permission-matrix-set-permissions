#Requires -Modules Pester, Assert
#Requires -Version 5.1

BeforeAll {
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

    $testParentFolder = (New-Item 'TestDrive:\testFolder' -ItemType Directory -Force).FullName

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')

    Mock Write-Warning
}

Describe 'the mandatory parameters are' {
    It "<_>" -TestCases @('Path', 'Action', 'Matrix') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory | Should -BeTrue
    } 
}
Describe 'create a Matrix object' {
    BeforeAll {
        $testParent = New-Item 'TestDrive:\parent' -ItemType Directory -Force
        $testFolder = New-Item 'TestDrive:\parent\folder' -ItemType Directory -Force

        $testParams = @{
            Path   = $testParent.FullName
            Action = 'Check'
            Matrix = @(
                [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true }
                [PSCustomObject]@{Path = 'folder'; ACL = @{$env:USERNAME = 'R' } }
            )
        }
        .$testScript @testParams
    }
    It 'one Matrix object for each folder' {
        $Matrix | Should -HaveCount 2
    }
    Context 'the property FolderAcl' {
        It 'is added for each folder' {
            $Matrix.FolderAcl | Should -HaveCount 2
        }
        It 'is of type DirectorySecurity' {
            $Matrix.FolderAcl | ForEach-Object {
                $_ | Should -BeOfType [System.Security.AccessControl.DirectorySecurity]
            }
        }
        It "has 'BUILTIN\Administrators' added with 'FullControl'" {
            foreach ($testFolderAcl in $Matrix.FolderAcl) {
                $testFolderAcl.Access[0].IdentityReference | 
                Should -Be 'BUILTIN\Administrators'
                $testFolderAcl.Access[0].FileSystemRights | 
                Should -Be 'FullControl'
            }
        }
        It "has 'BUILTIN\Administrators' set as owner" {
            foreach ($testFolderAcl in $Matrix.FolderAcl) {
                $testFolderAcl.Owner | 
                Should -Be 'BUILTIN\Administrators'
            }
        }
        It 'contains all other requested permissions' {
            foreach ($testFolderAcl in $Matrix.FolderAcl) {
                $testFolderAcl.Access[1] | 
                Should -Not -BeNullOrEmpty
            }
        }
    }
    Context 'the property Path' {
        It 'is converted to the folder FullName' {
            $Matrix[0].Path | Should -Be "\\?\$($testParent.FullName)"
            $Matrix[1].Path | Should -Be "\\?\$($testFolder.FullName)"
        }
    }
    Context 'the property Parent' {
        It 'is only marked as true for the parent folder' {
            $Matrix[0].Parent | Should -BeTrue
            $Matrix[1].Parent | Should -BeFalse
        }
    }
    Context 'extra property added for later comparison' {
        It 'InheritedFolderAcl' {
            foreach ($testMatrix in $Matrix) {
                $testMatrix.InheritedFolderAcl.GetType().Name | 
                Should -Be 'DirectorySecurity'    
            }
        }
        It 'InheritedFileAcl' {
            foreach ($testMatrix in $Matrix) {
                $testMatrix.InheritedFileAcl.GetType().Name | 
                Should -Be 'FileSecurity'    
            }
        }
    }
}
$testCases = @(
    @{
        name       = 'with no ignored folders'
        state      = @{
            before = @{
                folders = @(
                    '{0}\FolderA',
                    '{0}\FolderB',
                    '{0}\FolderC\Level1\Level2'
                )
            }
        }
        testMatrix = @(
            [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'R' }; Parent = $true }
            [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
            [PSCustomObject]@{Path = 'FolderB'; ACL = @{$env:USERNAME = 'R' } }
            [PSCustomObject]@{Path = 'FolderC'; ACL = @{$env:USERNAME = 'R' } }
            [PSCustomObject]@{Path = 'FolderC\Level1\Level2'; ACL = @{ } }
        )
        expected   = @{
            nonInheritanceTested = @(
                '\\?\{0}',
                '\\?\{0}\FolderB',
                '\\?\{0}\FolderC'
            )
            inheritanceTested    = @(
                '\\?\{0}\file',
                '\\?\{0}\FolderA',
                '\\?\{0}\FolderA\file',
                '\\?\{0}\FolderB\file',
                '\\?\{0}\FolderC\file',
                '\\?\{0}\FolderC\Level1',
                '\\?\{0}\FolderC\Level1\file',
                '\\?\{0}\FolderC\Level1\Level2',
                '\\?\{0}\FolderC\Level1\Level2\file'
            )
        }
    }
    @{
        name       = 'with an ignored folder all its subfolders and files are not checked'
        state      = @{
            before = @{
                folders = @(
                    '{0}\FolderA',
                    '{0}\FolderB',
                    '{0}\FolderC'
                )
            }
        }
        testMatrix = @(
            [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'R' }; Parent = $true }
            [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
            [PSCustomObject]@{Path = 'FolderB'; ACL = @{$env:USERNAME = 'R' } }
            [PSCustomObject]@{Path = 'FolderC'; ACL = @{$env:USERNAME = 'R' }; Ignore = $true }
        )
        expected   = @{
            nonInheritanceTested = @(
                '\\?\{0}',
                '\\?\{0}\FolderB'
            )
            inheritanceTested    = @(
                '\\?\{0}\file',
                '\\?\{0}\FolderA',
                '\\?\{0}\FolderA\file',
                '\\?\{0}\FolderB\file'
            )
        }
    }
    @{
        name       = 'with an ignored folder, its subfolders are checked only when they have permissions defined in the matrix'
        state      = @{
            before = @{
                folders = @(
                    '{0}\FolderA',
                    '{0}\FolderB\SubFolder\Reports',
                    '{0}\FolderB\SubFolder\Accounting'
                )
            }
        }
        testMatrix = @(
            [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'R' }; Parent = $true }
            [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
            [PSCustomObject]@{Path = 'FolderB\SubFolder'; ACL = @{$env:USERNAME = 'R' }; Ignore = $true }
            [PSCustomObject]@{Path = 'FolderB\SubFolder\Reports'; ACL = @{$env:USERNAME = 'R' } }
            [PSCustomObject]@{Path = 'FolderB\SubFolder\Accounting'; ACL = @{ } } # ignored because SubFolder is ignored
        )
        expected   = @{
            nonInheritanceTested = @(
                '\\?\{0}',
                '\\?\{0}\FolderB\SubFolder\Reports'
            )
            inheritanceTested    = @(
                '\\?\{0}\file',
                '\\?\{0}\FolderA',
                '\\?\{0}\FolderA\file',
                '\\?\{0}\FolderB',
                '\\?\{0}\FolderB\File',
                '\\?\{0}\FolderB\SubFolder\Reports\File'
            )
        }
    }
    @{
        name       = 'with an ignored folder, all folders below are not checked unless they have permissions set, in that case permission checking resumes for all its subfolders'
        state      = @{
            before = @{
                folders = @(
                    '{0}\FolderA',
                    '{0}\FolderB\SubFolder\Reports\Year\2020',
                    '{0}\FolderB\SubFolder\Accounting',
                    '{0}\FolderB\SubFolder\Sales',
                    '{0}\FolderC'
                )
            }
        }
        testMatrix = @(
            [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'R' }; Parent = $true }
            [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
            [PSCustomObject]@{Path = 'FolderB\SubFolder'; ACL = @{$env:USERNAME = 'R' }; Ignore = $true }
            [PSCustomObject]@{Path = 'FolderB\SubFolder\Reports'; ACL = @{$env:USERNAME = 'R' } }
            [PSCustomObject]@{Path = 'FolderB\SubFolder\Reports\Year\2020'; ACL = @{ } } # tested because it falls under the Reports folder
            [PSCustomObject]@{Path = 'FolderB\SubFolder\Accounting'; ACL = @{ } } # ignored because SubFolder is ignored
            [PSCustomObject]@{Path = 'FolderB\SubFolder\Sales'; ACL = @{ } } # ignored because SubFolder is ignored
            [PSCustomObject]@{Path = 'FolderC'; ACL = @{$env:USERNAME = 'R' } }
        )
        expected   = @{
            nonInheritanceTested = @(
                '\\?\{0}',
                '\\?\{0}\FolderB\SubFolder\Reports',
                '\\?\{0}\FolderC'
            )
            inheritanceTested    = @(
                '\\?\{0}\file',
                '\\?\{0}\FolderA',
                '\\?\{0}\FolderA\file',
                '\\?\{0}\FolderB',
                '\\?\{0}\FolderB\File',
                '\\?\{0}\FolderB\SubFolder\Reports\File',
                '\\?\{0}\FolderB\SubFolder\Reports\Year',
                '\\?\{0}\FolderB\SubFolder\Reports\Year\File',
                '\\?\{0}\FolderB\SubFolder\Reports\Year\2020',
                '\\?\{0}\FolderB\SubFolder\Reports\Year\2020\File',
                '\\?\{0}\FolderC\File'
            )
        }
    }
    @{
        name       = 'sub folders of ignored folders are not checked'
        state      = @{
            before = @{
                folders = @(
                    '{0}\FolderA',
                    '{0}\FolderB\SubFolder\Reports\Year\2020\CEM\Profit',
                    '{0}\FolderB\SubFolder\Reports\Year\2020\CEM\Loss\HR',
                    '{0}\FolderB\SubFolder\Accounting',
                    '{0}\FolderB\SubFolder\Sales',
                    '{0}\FolderC'
                )
            }
        }
        testMatrix = @(
            [PSCustomObject]@{
                Path   = 'Path'; 
                ACL    = @{$env:USERNAME = 'R' }; 
                Parent = $true 
            }
            [PSCustomObject]@{
                Path   = 'FolderB\SubFolder'; 
                ACL    = @{$env:USERNAME = 'R' }; 
                Ignore = $true 
            }
            [PSCustomObject]@{
                Path = 'FolderB\SubFolder\Reports'; 
                ACL  = @{$env:USERNAME = 'R' } 
            }
            [PSCustomObject]@{
                Path   = 'FolderB\SubFolder\Reports\Year\2020\CEM';
                ACL    = @{ };
                Ignore = $true 
            } 
            [PSCustomObject]@{
                Path = 'FolderB\SubFolder\Reports\Year\2020\CEM\Loss\HR';
                ACL  = @{ $env:USERNAME = 'R' } 
            } 
            [PSCustomObject]@{
                Path = 'FolderC';
                ACL  = @{$env:USERNAME = 'R' } 
            }
        )
        expected   = @{
            nonInheritanceTested = @(
                '\\?\{0}',
                '\\?\{0}\FolderB\SubFolder\Reports',
                '\\?\{0}\FolderB\SubFolder\Reports\Year\2020\CEM\Loss\HR',
                '\\?\{0}\FolderC'
            )
            inheritanceTested    = @(
                '\\?\{0}\file',
                '\\?\{0}\FolderA',
                '\\?\{0}\FolderA\file',
                '\\?\{0}\FolderB',
                '\\?\{0}\FolderB\File',
                '\\?\{0}\FolderB\SubFolder\Reports\File',
                '\\?\{0}\FolderB\SubFolder\Reports\Year',
                '\\?\{0}\FolderB\SubFolder\Reports\Year\File',
                '\\?\{0}\FolderB\SubFolder\Reports\Year\2020',
                '\\?\{0}\FolderB\SubFolder\Reports\Year\2020\File',
                '\\?\{0}\FolderB\SubFolder\Reports\Year\2020\CEM\Loss\HR\File',
                '\\?\{0}\FolderC\File'
            )
        }
    }
    @{
        name       = 'when only Path is in the matrix and has permissions all files and folders need to be inherited'
        state      = @{
            before = @{
                folders = @(
                    '{0}\FolderA',
                    '{0}\FolderB'
                )
            }
        }
        testMatrix = @(
            [PSCustomObject]@{
                Path   = 'Path'; 
                ACL    = @{$env:USERNAME = 'R' }; 
                Parent = $true 
            }
        )
        expected   = @{
            nonInheritanceTested = @(
                '\\?\{0}'
            )
            inheritanceTested    = @(
                '\\?\{0}\FolderA',
                '\\?\{0}\file',
                '\\?\{0}\FolderA\file',
                '\\?\{0}\FolderB',
                '\\?\{0}\FolderB\File'
            )
        }
    }
    @{
        name       = 'when Path is ignored and a child folder has permissions, only the child folder is checked'
        state      = @{
            before = @{
                folders = @(
                    '{0}\FolderA',
                    '{0}\FolderB'
                )
            }
        }
        testMatrix = @(
            [PSCustomObject]@{
                Path   = 'Path'
                ACL    = @{}
                Parent = $true 
                Ignore = $true
            }
            [PSCustomObject]@{
                Path = 'FolderA'
                ACL  = @{$env:USERNAME = 'R' }
            }
        )
        expected   = @{
            nonInheritanceTested = @(
                '\\?\{0}\FolderA'
            )
            inheritanceTested    = @(
                '\\?\{0}\FolderA\file'
            )
        }
    }
)
Describe 'when the script runs for a matrix' {
    Context '<name>' -ForEach $testCases {
        BeforeAll {
            Remove-Item $testParentFolder -Recurse -Force
       
            if ($state.before.folders) {
                $state.before.folders | ForEach-Object {
                    $tmpTestFolder = ($_ -f $testParentFolder)
                    New-Item -Path $tmpTestFolder -ItemType Directory -Force
                }
                
                @(
                    (Get-ChildItem $testParentFolder -Recurse -Directory).FullName + $testParentFolder
                ) | ForEach-Object {
                    $tmpTestFile = Join-Path $_ 'file'
                    New-Item -Path $tmpTestFile -ItemType File
                }
            }

            $testParams = @{
                Path   = $testParentFolder
                Action = 'Fix'
                Matrix = Copy-ObjectHC $testMatrix
            }
            $testResult = .$testScript @testParams
        }
        It 'all non inherited folders are checked' {
            $testedNonInheritedFolders | Should -Not -BeNullOrEmpty -Because 'it is a production script variable'
                 
            $expected.nonInheritanceTested | ForEach-Object {
                $testedNonInheritedFolders.Keys | 
                Should -Contain ($_ -f $testParams.Path) 
                # Pester scoping issue: variables not available in TestCases
            }
            $testedNonInheritedFolders.Count | 
            Should -BeExactly $expected.nonInheritanceTested.Count
        }
        It 'all files and folders that should be inherited are checked' {
            $testedInheritedFilesAndFolders | Should -Not -BeNullOrEmpty -Because 'it is a production script variable'
    
            $expected.inheritanceTested | ForEach-Object {
                $testedInheritedFilesAndFolders.Keys | 
                Should -Contain ($_ -f $testParams.Path)
                # Pester scoping issue: variables not available in TestCases
            }
            $testedInheritedFilesAndFolders.Count | 
            Should -BeExactly $expected.inheritanceTested.Count
        }
        It 'output is generated for ignored folders in an information object' {
            if ($testIgnoredFolders = $testMatrix | Where-Object ignore) {
                $actual = $testResult | Where-Object { 
                        ($_.Type -eq 'Information') -and
                        ($_.Name -eq 'Ignored folder') }

                $testIgnoredFolders = $testIgnoredFolders | ForEach-Object {
                    if ($_.Parent) { $testParentFolder }
                    else { Join-Path $testParentFolder $_.Path }
                }

                $actual.Value | Should -Be $testIgnoredFolders    
            }
        }
    } 
} -Tag test
Describe 'Permissions' {
    BeforeEach {
        Remove-Item 'TestDrive:\*' -Recurse -Force
    }
    Context 'are not corrected when they are correct for' {
        It 'List, Write, Read on the parent folder' {
            $testParams = @{
                Path   = $testParentFolder
                Action = 'Fix'
                Matrix = @(
                    [PSCustomObject]@{Path = 'Path'; ACL = @{
                            $env:USERNAME = 'L' ; $testUser = 'W'; $testUser2 = 'R';
                        }; Parent = $true 
                    }
                    [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                )
            }

            #region Create all folders
            $testParams.Matrix | Select-Object -Skip 1 | ForEach-Object {
                New-Item -Path (Join-Path $testParams.Path $_.Path) -ItemType Directory -Force
            }
            #endregion

            #region Create all files
            @(, (Get-ChildItem $testParams.Path -Recurse -Directory).FullName + $testParams.Path) | 
            ForEach-Object {
                New-Item -Path (Join-Path $_ 'file') -ItemType File
            }
            #endregion

            #region Set correct permissions on parent folder
            $testItem = Get-Item $testParams.Path

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

            $Actual = .$testScript @testParams | Where-Object {
                ($_Name -eq 'Non inherited folder incorrect permissions') -or
                ($_Name -eq 'Inherited permissions incorrect')
            }

            $Actual | Should -BeNullOrEmpty
        } 
        It 'List only on the parent folder' {
            $testParams = @{
                Path   = $testParentFolder
                Action = 'Fix'
                Matrix = @(
                    [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L'; $testUser = 'L' }; Parent = $true }
                    [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                )
            }

            #region Create all folders
            $testParams.Matrix | Select-Object -Skip 1 | ForEach-Object {
                New-Item -Path (Join-Path $testParams.Path $_.Path) -ItemType Directory -Force
            }
            #endregion

            #region Create all files
            @(, (Get-ChildItem $testParams.Path -Recurse -Directory).FullName + $testParams.Path) | 
            ForEach-Object {
                New-Item -Path (Join-Path $_ 'file') -ItemType File
            }
            #endregion

            #region Set correct permissions on parent folder
            $testItem = Get-Item $testParams.Path

            $acl = New-Object System.Security.AccessControl.DirectorySecurity
            $acl.SetAccessRuleProtection($true, $false)
            $acl.SetOwner($BuiltinAdmin)

            $aceList = @($AdminFullControlFolderAce)
            $aceList += New-TestAceHC -Type 'Folder' -Access 'L' -Name $env:USERNAME
            $aceList += New-TestAceHC -Type 'Folder' -Access 'L' -Name $testUser
            $aceList.foreach( { $acl.AddAccessRule($_) })

            $testItem.SetAccessControl($acl)
            #endregion

            $Actual = .$testScript @testParams | Where-Object {
                ($_Name -eq 'Non inherited folder incorrect permissions') -or
                ($_Name -eq 'Inherited permissions incorrect')
            }

            $Actual | Should -BeNullOrEmpty
        } 
        It 'List only on the parent folder and Read on a subfolder' {
            $testParams = @{
                Path   = $testParentFolder
                Action = 'Fix'
                Matrix = @(
                    [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L'; $testUser = 'L' }; Parent = $true }
                    [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser = 'R' } }
                    [PSCustomObject]@{Path = 'FolderB'; ACL = @{ } }
                )
            }

            #region Create all folders
            $testParams.Matrix | Select-Object -Skip 1 | ForEach-Object {
                New-Item -Path (Join-Path $testParams.Path $_.Path) -ItemType Directory -Force
            }
            #endregion

            #region Create all files
            @(, (Get-ChildItem $testParams.Path -Recurse -Directory).FullName + $testParams.Path) | 
            ForEach-Object {
                New-Item -Path (Join-Path $_ 'file') -ItemType File
            }
            #endregion

            #region Set correct permissions on parent folder
            $testItem = Get-Item $testParams.Path

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
            $testItem = Get-Item "$($testParams.Path)\FolderA"

            $acl = New-Object System.Security.AccessControl.DirectorySecurity
            $acl.SetAccessRuleProtection($true, $false)
            $acl.SetOwner($BuiltinAdmin)

            $aceList = @($AdminFullControlFolderAce)
            $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser
            $aceList.foreach( { $acl.AddAccessRule($_) })

            $testItem.SetAccessControl($acl)
            #endregion

            $Actual = .$testScript @testParams | Where-Object {
                ($_Name -eq 'Non inherited folder incorrect permissions') -or
                ($_Name -eq 'Inherited permissions incorrect')
            }

            $Actual | Should -BeNullOrEmpty
        } 
        It 'List only on the parent folder and different permissions on subfolders' {
            $testParams = @{
                Path   = $testParentFolder
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
            $testParams.Matrix | Select-Object -Skip 1 | ForEach-Object {
                New-Item -Path (Join-Path $testParams.Path $_.Path) -ItemType Directory -Force
            }
            #endregion

            #region Create all files
            @(, (Get-ChildItem $testParams.Path -Recurse -Directory).FullName + $testParams.Path) | 
            ForEach-Object {
                New-Item -Path (Join-Path $_ 'file') -ItemType File
            }
            #endregion

            #region Set correct permissions on parent folder
            $testItem = Get-Item $testParams.Path

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
            $testItem = Get-Item "$($testParams.Path)\FolderA"

            $acl = New-Object System.Security.AccessControl.DirectorySecurity
            $acl.SetAccessRuleProtection($true, $false)
            $acl.SetOwner($BuiltinAdmin)

            $aceList = @($AdminFullControlFolderAce)
            $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser
            $aceList.foreach( { $acl.AddAccessRule($_) })

            $testItem.SetAccessControl($acl)
            #endregion

            #region Set correct permissions on a sub folder
            $testItem = Get-Item "$($testParams.Path)\FolderB\SubFolderB"

            $acl = New-Object System.Security.AccessControl.DirectorySecurity
            $acl.SetAccessRuleProtection($true, $false)
            $acl.SetOwner($BuiltinAdmin)

            $aceList = @($AdminFullControlFolderAce)
            $aceList += New-TestAceHC -Type 'Folder' -Access 'W' -Name $testUser2
            $aceList.foreach( { $acl.AddAccessRule($_) })

            $testItem.SetAccessControl($acl)
            #endregion

            #region Set correct permissions on a sub folder
            $testItem = Get-Item "$($testParams.Path)\FolderB\SubFolderC"

            $acl = New-Object System.Security.AccessControl.DirectorySecurity
            $acl.SetAccessRuleProtection($true, $false)
            $acl.SetOwner($BuiltinAdmin)

            $aceList = @($AdminFullControlFolderAce)
            $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser2
            $aceList.foreach( { $acl.AddAccessRule($_) })

            $testItem.SetAccessControl($acl)
            #endregion

            $Actual = .$testScript @testParams | Where-Object {
                ($_Name -eq 'Non inherited folder incorrect permissions') -or
                ($_Name -eq 'Inherited permissions incorrect')
            }

            $Actual | Should -BeNullOrEmpty
        } 
        It 'folders that are not in the matrix as they should be inherited' {
            $testParams = @{
                Path   = $testParentFolder
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
            $testParams.Matrix | Select-Object -Skip 1 | ForEach-Object {
                New-Item -Path (Join-Path $testParams.Path $_.Path) -ItemType Directory -Force
            }
            #endregion

            #region Create all files
            @(, (Get-ChildItem $testParams.Path -Recurse -Directory).FullName + $testParams.Path) | 
            ForEach-Object {
                New-Item -Path (Join-Path $_ 'file') -ItemType File
            }
            #endregion

            #region Set correct permissions on parent folder
            $testItem = Get-Item $testParams.Path

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
            $testItem = Get-Item "$($testParams.Path)\FolderA"

            $acl = New-Object System.Security.AccessControl.DirectorySecurity
            $acl.SetAccessRuleProtection($true, $false)
            $acl.SetOwner($BuiltinAdmin)

            $aceList = @($AdminFullControlFolderAce)
            $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser
            $aceList.foreach( { $acl.AddAccessRule($_) })

            $testItem.SetAccessControl($acl)
            #endregion

            #region Set correct permissions on a sub folder
            $testItem = Get-Item "$($testParams.Path)\FolderB\SubFolderB"

            $acl = New-Object System.Security.AccessControl.DirectorySecurity
            $acl.SetAccessRuleProtection($true, $false)
            $acl.SetOwner($BuiltinAdmin)

            $aceList = @($AdminFullControlFolderAce)
            $aceList += New-TestAceHC -Type 'Folder' -Access 'W' -Name $testUser2
            $aceList.foreach( { $acl.AddAccessRule($_) })

            $testItem.SetAccessControl($acl)
            #endregion

            #region Set correct permissions on a sub folder
            $testItem = Get-Item "$($testParams.Path)\FolderB\SubFolderC"

            $acl = New-Object System.Security.AccessControl.DirectorySecurity
            $acl.SetAccessRuleProtection($true, $false)
            $acl.SetOwner($BuiltinAdmin)

            $aceList = @($AdminFullControlFolderAce)
            $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser2
            $aceList.foreach( { $acl.AddAccessRule($_) })

            $testItem.SetAccessControl($acl)
            #endregion

            #region Create extra folders not defined in the matrix
            $null = New-Item -Path "$($testParams.Path)\FolderB\OtherFolder" -ItemType Directory -Force
            $null = New-Item -Path "$($testParams.Path)\FolderC" -ItemType Directory -Force
            $null = New-Item -Path "$($testParams.Path)\FolderC\File" -ItemType File -Force
            $null = New-Item -Path "$($testParams.Path)\FolderD\Fruits\Kiwi\Green" -ItemType Directory -Force
            $null = New-Item -Path "$($testParams.Path)\FolderD\Fruits\Kiwi\Green\File" -ItemType File -Force
            #endregion

            $Actual = .$testScript @testParams | Where-Object {
                ($_Name -eq 'Non inherited folder incorrect permissions') -or
                ($_Name -eq 'Inherited permissions incorrect')
            }

            $Actual | Should -BeNullOrEmpty
        } 
    }
    Context 'are corrected when they are incorrect when' {
        Context 'a folder that should have explicit permissions has' {
            It 'incorrect explicit permissions' {
                $testParams = @{
                    Path   = $testParentFolder
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
                $testParams.Matrix | Select-Object -Skip 1 | ForEach-Object {
                    New-Item -Path (Join-Path $testParams.Path $_.Path) -ItemType Directory -Force
                }
                #endregion

                #region Create all files
                @(, (Get-ChildItem $testParams.Path -Recurse -Directory).FullName + $testParams.Path) | 
                ForEach-Object {
                    New-Item -Path (Join-Path $_ 'file') -ItemType File
                }
                #endregion

                #region Set incorrect permissions on parent folder
                $testItem = Get-Item $testParams.Path

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
                $testItem = Get-Item "$($testParams.Path)\FolderA"

                $acl = New-Object System.Security.AccessControl.DirectorySecurity
                $acl.SetAccessRuleProtection($true, $false)
                $acl.SetOwner($BuiltinAdmin)

                $aceList = @($AdminFullControlFolderAce)
                $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser
                $aceList.foreach( { $acl.AddAccessRule($_) })

                $testItem.SetAccessControl($acl)
                #endregion

                #region Set correct permissions on a sub folder
                $testItem = Get-Item "$($testParams.Path)\FolderB\SubFolderB"

                $acl = New-Object System.Security.AccessControl.DirectorySecurity
                $acl.SetAccessRuleProtection($true, $false)
                $acl.SetOwner($BuiltinAdmin)

                $aceList = @($AdminFullControlFolderAce)
                $aceList += New-TestAceHC -Type 'Folder' -Access 'W' -Name $testUser2
                $aceList.foreach( { $acl.AddAccessRule($_) })

                $testItem.SetAccessControl($acl)
                #endregion

                #region Set correct permissions on a sub folder
                $testItem = Get-Item "$($testParams.Path)\FolderB\SubFolderC"

                $acl = New-Object System.Security.AccessControl.DirectorySecurity
                $acl.SetAccessRuleProtection($true, $false)
                $acl.SetOwner($BuiltinAdmin)

                $aceList = @($AdminFullControlFolderAce)
                $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser2
                $aceList.foreach( { $acl.AddAccessRule($_) })

                $testItem.SetAccessControl($acl)
                #endregion

                $Actual = .$testScript @testParams | Where-Object Name -EQ 'Non inherited folder incorrect permissions'
                    
                $Actual.Value | Should -Be $testParams.Path
            } 
            It 'the correct explicit permissions but one ACE too much' {
                $testParams = @{
                    Path   = $testParentFolder
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
                $testParams.Matrix | Select-Object -Skip 1 | ForEach-Object {
                    New-Item -Path (Join-Path $testParams.Path $_.Path) -ItemType Directory -Force
                }
                #endregion

                #region Create all files
                @(, (Get-ChildItem $testParams.Path -Recurse -Directory).FullName + $testParams.Path) | 
                ForEach-Object {
                    New-Item -Path (Join-Path $_ 'file') -ItemType File
                }
                #endregion

                #region Set incorrect permissions on parent folder
                $testItem = Get-Item $testParams.Path

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
                $testItem = Get-Item "$($testParams.Path)\FolderA"

                $acl = New-Object System.Security.AccessControl.DirectorySecurity
                $acl.SetAccessRuleProtection($true, $false)
                $acl.SetOwner($BuiltinAdmin)

                $aceList = @($AdminFullControlFolderAce)
                $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser
                $aceList.foreach( { $acl.AddAccessRule($_) })

                $testItem.SetAccessControl($acl)
                #endregion

                #region Set correct permissions on a sub folder
                $testItem = Get-Item "$($testParams.Path)\FolderB\SubFolderB"

                $acl = New-Object System.Security.AccessControl.DirectorySecurity
                $acl.SetAccessRuleProtection($true, $false)
                $acl.SetOwner($BuiltinAdmin)

                $aceList = @($AdminFullControlFolderAce)
                $aceList += New-TestAceHC -Type 'Folder' -Access 'W' -Name $testUser2
                $aceList.foreach( { $acl.AddAccessRule($_) })

                $testItem.SetAccessControl($acl)
                #endregion

                #region Set correct permissions on a sub folder
                $testItem = Get-Item "$($testParams.Path)\FolderB\SubFolderC"

                $acl = New-Object System.Security.AccessControl.DirectorySecurity
                $acl.SetAccessRuleProtection($true, $false)
                $acl.SetOwner($BuiltinAdmin)

                $aceList = @($AdminFullControlFolderAce)
                $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser2
                $aceList.foreach( { $acl.AddAccessRule($_) })

                $testItem.SetAccessControl($acl)
                #endregion

                $Actual = .$testScript @testParams | Where-Object Name -EQ 'Non inherited folder incorrect permissions'
                    
                $Actual.Value | Should -Be $testParams.Path
            } 
            It 'inherited permissions' {
                $testParams = @{
                    Path   = $testParentFolder
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
                $testParams.Matrix | Select-Object -Skip 1 | ForEach-Object {
                    New-Item -Path (Join-Path $testParams.Path $_.Path) -ItemType Directory -Force
                }
                #endregion

                #region Create all files
                @(, (Get-ChildItem $testParams.Path -Recurse -Directory).FullName + $testParams.Path) | 
                ForEach-Object {
                    New-Item -Path (Join-Path $_ 'file') -ItemType File
                }
                #endregion

                #region Set incorrect permissions on parent folder
                # $testParams.Path is inherited
                #endregion

                #region Set correct permissions on a sub folder
                $testItem = Get-Item "$($testParams.Path)\FolderA"

                $acl = New-Object System.Security.AccessControl.DirectorySecurity
                $acl.SetAccessRuleProtection($true, $false)
                $acl.SetOwner($BuiltinAdmin)

                $aceList = @($AdminFullControlFolderAce)
                $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser
                $aceList.foreach( { $acl.AddAccessRule($_) })

                $testItem.SetAccessControl($acl)
                #endregion

                #region Set correct permissions on a sub folder
                $testItem = Get-Item "$($testParams.Path)\FolderB\SubFolderB"

                $acl = New-Object System.Security.AccessControl.DirectorySecurity
                $acl.SetAccessRuleProtection($true, $false)
                $acl.SetOwner($BuiltinAdmin)

                $aceList = @($AdminFullControlFolderAce)
                $aceList += New-TestAceHC -Type 'Folder' -Access 'W' -Name $testUser2
                $aceList.foreach( { $acl.AddAccessRule($_) })

                $testItem.SetAccessControl($acl)
                #endregion

                #region Set correct permissions on a sub folder
                $testItem = Get-Item "$($testParams.Path)\FolderB\SubFolderC"

                $acl = New-Object System.Security.AccessControl.DirectorySecurity
                $acl.SetAccessRuleProtection($true, $false)
                $acl.SetOwner($BuiltinAdmin)

                $aceList = @($AdminFullControlFolderAce)
                $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser2
                $aceList.foreach( { $acl.AddAccessRule($_) })

                $testItem.SetAccessControl($acl)
                #endregion

                $Actual = .$testScript @testParams | Where-Object Name -EQ 'Non inherited folder incorrect permissions'
                    
                $Actual.Value | Should -Be $testParams.Path
            } 
        }
        Context 'a file has' {
            It 'explicit permissions' {
                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'Fix'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L'; $testUser = 'L'; $testUser2 = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser = 'R' } }
                    )
                }

                #region Create all folders
                $testParams.Matrix | Select-Object -Skip 1 | ForEach-Object {
                    New-Item -Path (Join-Path $testParams.Path $_.Path) -ItemType Directory -Force
                }
                #endregion

                #region Create all files
                @(, (Get-ChildItem $testParams.Path -Recurse -Directory).FullName + $testParams.Path) | 
                ForEach-Object {
                    New-Item -Path (Join-Path $_ 'file') -ItemType File
                }
                #endregion

                #region Set correct permissions on parent folder
                $testItem = Get-Item $testParams.Path

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
                $testItem = Get-Item "$($testParams.Path)\FolderA"

                $acl = New-Object System.Security.AccessControl.DirectorySecurity
                $acl.SetAccessRuleProtection($true, $false)
                $acl.SetOwner($BuiltinAdmin)

                $aceList = @($AdminFullControlFolderAce)
                $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser
                $aceList.foreach( { $acl.AddAccessRule($_) })

                $testItem.SetAccessControl($acl)
                #endregion

                #region Set incorrect permissions on a file
                $testItem = Get-Item "$($testParams.Path)\FolderA\File"

                $acl = New-Object System.Security.AccessControl.FileSecurity
                $acl.SetAccessRuleProtection($true, $false)
                $acl.SetOwner($BuiltinAdmin)

                $aceList = @($AdminFullControlIFileAce)
                $aceList += New-TestAceHC -Type 'InheritedFile' -Access 'W' -Name $testUser2
                $aceList.foreach( { $acl.AddAccessRule($_) })

                $testItem.SetAccessControl($acl)
                #endregion


                $Actual = .$testScript @testParams

                $Actual | Where-Object Name -EQ 'Inherited permissions incorrect'
                    
                $Actual.Value | Should -Be "$($testParams.Path)\FolderA\File"
            } 
        }
        Context 'a folder that should have inherited permissions' {
            It 'in the matrix has explicit permissions' {
                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'Fix'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L'; $testUser = 'L'; $testUser2 = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser = 'R' } }
                        [PSCustomObject]@{Path = 'FolderB\SubFolderB'; ACL = @{$testUser2 = 'W' } }
                        [PSCustomObject]@{Path = 'FolderB\SubFolderC'; ACL = @{$testUser2 = 'R' } }
                    )
                }

                #region Create all folders
                $testParams.Matrix | Select-Object -Skip 1 | ForEach-Object {
                    New-Item -Path (Join-Path $testParams.Path $_.Path) -ItemType Directory -Force
                }
                #endregion

                #region Create all files
                @(, (Get-ChildItem $testParams.Path -Recurse -Directory).FullName + $testParams.Path) | 
                ForEach-Object {
                    New-Item -Path (Join-Path $_ 'file') -ItemType File
                }
                #endregion

                #region Set correct permissions on parent folder
                $testItem = Get-Item $testParams.Path

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
                $testItem = Get-Item "$($testParams.Path)\FolderA"

                $acl = New-Object System.Security.AccessControl.DirectorySecurity
                $acl.SetAccessRuleProtection($true, $false)
                $acl.SetOwner($BuiltinAdmin)

                $aceList = @($AdminFullControlFolderAce)
                $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser
                $aceList.foreach( { $acl.AddAccessRule($_) })

                $testItem.SetAccessControl($acl)
                #endregion

                #region Set correct permissions on a sub folder
                $testItem = Get-Item "$($testParams.Path)\FolderB\SubFolderB"

                $acl = New-Object System.Security.AccessControl.DirectorySecurity
                $acl.SetAccessRuleProtection($true, $false)
                $acl.SetOwner($BuiltinAdmin)

                $aceList = @($AdminFullControlFolderAce)
                $aceList += New-TestAceHC -Type 'Folder' -Access 'W' -Name $testUser2
                $aceList.foreach( { $acl.AddAccessRule($_) })

                $testItem.SetAccessControl($acl)
                #endregion

                #region Set correct permissions on a sub folder
                $testItem = Get-Item "$($testParams.Path)\FolderB\SubFolderC"

                $acl = New-Object System.Security.AccessControl.DirectorySecurity
                $acl.SetAccessRuleProtection($true, $false)
                $acl.SetOwner($BuiltinAdmin)

                $aceList = @($AdminFullControlFolderAce)
                $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser2
                $aceList.foreach( { $acl.AddAccessRule($_) })

                $testItem.SetAccessControl($acl)
                #endregion

                #region Set incorrect permissions on an inherited folder
                $testItem = Get-Item "$($testParams.Path)\FolderB"

                $acl = New-Object System.Security.AccessControl.DirectorySecurity
                $acl.SetAccessRuleProtection($true, $false)
                $acl.SetOwner($BuiltinAdmin)

                $aceList = @($AdminFullControlFolderAce)
                $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser2
                $aceList.foreach( { $acl.AddAccessRule($_) })

                $testItem.SetAccessControl($acl)
                #endregion

                $Actual = .$testScript @testParams
                    
                $Actual | Where-Object Name -EQ 'Inherited permissions incorrect'
                    
                $Actual.Value | Should -Be "$($testParams.Path)\FolderB"
            } 
            It 'not defined in the matrix has explicit permissions' {
                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'Fix'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L'; $testUser = 'L'; $testUser2 = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser = 'R' } }
                        [PSCustomObject]@{Path = 'FolderB\SubFolderB'; ACL = @{$testUser2 = 'W' } }
                        [PSCustomObject]@{Path = 'FolderB\SubFolderC'; ACL = @{$testUser2 = 'R' } }
                    )
                }

                #region Create all folders
                $testParams.Matrix | Select-Object -Skip 1 | ForEach-Object {
                    New-Item -Path (Join-Path $testParams.Path $_.Path) -ItemType Directory -Force
                }
                #endregion

                #region Create all files
                @(, (Get-ChildItem $testParams.Path -Recurse -Directory).FullName + $testParams.Path) | 
                ForEach-Object {
                    New-Item -Path (Join-Path $_ 'file') -ItemType File
                }
                #endregion

                #region Set correct permissions on parent folder
                $testItem = Get-Item $testParams.Path

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
                $testItem = Get-Item "$($testParams.Path)\FolderA"

                $acl = New-Object System.Security.AccessControl.DirectorySecurity
                $acl.SetAccessRuleProtection($true, $false)
                $acl.SetOwner($BuiltinAdmin)

                $aceList = @($AdminFullControlFolderAce)
                $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser
                $aceList.foreach( { $acl.AddAccessRule($_) })

                $testItem.SetAccessControl($acl)
                #endregion

                #region Set correct permissions on a sub folder
                $testItem = Get-Item "$($testParams.Path)\FolderB\SubFolderB"

                $acl = New-Object System.Security.AccessControl.DirectorySecurity
                $acl.SetAccessRuleProtection($true, $false)
                $acl.SetOwner($BuiltinAdmin)

                $aceList = @($AdminFullControlFolderAce)
                $aceList += New-TestAceHC -Type 'Folder' -Access 'W' -Name $testUser2
                $aceList.foreach( { $acl.AddAccessRule($_) })

                $testItem.SetAccessControl($acl)
                #endregion

                #region Set correct permissions on a sub folder
                $testItem = Get-Item "$($testParams.Path)\FolderB\SubFolderC"

                $acl = New-Object System.Security.AccessControl.DirectorySecurity
                $acl.SetAccessRuleProtection($true, $false)
                $acl.SetOwner($BuiltinAdmin)

                $aceList = @($AdminFullControlFolderAce)
                $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser2
                $aceList.foreach( { $acl.AddAccessRule($_) })

                $testItem.SetAccessControl($acl)
                #endregion

                #region Set incorrect permissions on an inherited folder not in the matrix file

                $testItem = New-Item -Path "$($testParams.Path)\FolderC" -ItemType Directory -Force

                $acl = New-Object System.Security.AccessControl.DirectorySecurity
                $acl.SetAccessRuleProtection($true, $false)
                $acl.SetOwner($BuiltinAdmin)

                $aceList = @($AdminFullControlFolderAce)
                $aceList += New-TestAceHC -Type 'Folder' -Access 'R' -Name $testUser2
                $aceList.foreach( { $acl.AddAccessRule($_) })

                $testItem.SetAccessControl($acl)
                #endregion

                $Actual = .$testScript @testParams
                    
                $Actual | Where-Object Name -EQ 'Inherited permissions incorrect'
                    
                $Actual.Value | Should -Be "$($testParams.Path)\FolderC"
            } 
        }
    }
}
Describe 'when Action is' {
    BeforeEach {
        Remove-Item $testParentFolder -Recurse -Force -EA ignore
    }
    Context 'New' {
        It "create the parent folder 'Path'" {
            $testParams = @{
                Path   = $testParentFolder
                Action = 'New'
                Matrix = @([PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true })
            }

            .$testScript @testParams

            $testParams.Path | Should -Exist
        } 
        It 'create a FatalError object when the parent folder is already present' {
            $testParams = @{
                Path   = $testParentFolder
                Action = 'New'
                Matrix = [PSCustomObject]@{Name = 'test' }
            }

            New-Item -Path $testParams.Path -ItemType Directory

            $Actual = .$testScript @testParams

            $Expected = [PSCustomObject]@{
                Type        = 'FatalError'
                Name        = 'Parent folder exists already'
                Description = "The folder defined as 'Path' in the worksheet 'Settings' cannot be present on the remote machine when 'Action=New' is used. Please use 'Action' with value 'Check' or 'Fix' instead."
                Value       = $testParams.Path
            }

            $Actual.Type | Should -Be $Expected.Type
            $Actual.Name | Should -Be $Expected.Name
            $Actual.Description | Should -Be $Expected.Description
            $Actual.Value | Should -Be $Expected.Value
        } 
        Context 'folders in the matrix that need to be created' {
            It 'are created' {
                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'New'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                        [PSCustomObject]@{Path = 'FolderB\FolderC'; ACL = @{ } }
                    )
                }

                .$testScript @testParams

                $testParams.Path | Should -Exist
                $testParams.Path + '\FolderA' | Should -Exist
                $testParams.Path + '\FolderB\FolderC' | Should -Exist
            } 
            It 'are registered in a Warning object' {
                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'New'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                        [PSCustomObject]@{Path = 'FolderB\FolderC'; ACL = @{ } }
                    )
                }

                $Actual = .$testScript @testParams | Where-Object Name -Like "*child folder*"

                $Actual.Type | Should -Be 'Warning'
                $Actual.Name | Should -Be 'Child folder created'

                @(
                    "$($testParams.Path)",
                    "$($testParams.Path)\FolderA",
                    "$($testParams.Path)\FolderB\FolderC"
                ).ForEach( {
                        $Actual.Value | Should -Contain $_
                    })
                $actual.Value.Count | Should -BeExactly 3
            } 
            It 'are not created when Path is set to Ignore' {
                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'New'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }; Ignore = $true }
                        [PSCustomObject]@{Path = 'FolderB\FolderC'; ACL = @{ } }
                    )
                }

                .$testScript @testParams

                "$($testParams.Path)\FolderA" | Should -Not -Exist
                "$($testParams.Path)\FolderB\FolderC" | Should -Exist
            } 
        }
        Context 'set permissions' {
            It 'on the parent folder' {
                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'New'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                    )
                }

                .$testScript @testParams

                $Actual = (Get-Acl -Path $testParams.Path).Access

                $Actual.Count | Should -BeExactly 2 -Because "ACL is 'BUILTIN\Administrators' and '$testUser'."
                $Actual[0].IdentityReference | Should -Be 'BUILTIN\Administrators'
                $Actual[1].IdentityReference | Should -Be "$env:USERDOMAIN\$testUser"
            } 
            It 'on the child folders' {
                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'New'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                        [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' } }
                    )
                }

                .$testScript @testParams

                $Actual = (Get-Acl -Path $testParams.Path).Access
                $Actual.Count | Should -BeExactly 2
                $Actual[0].IdentityReference | Should -Be 'BUILTIN\Administrators'
                $Actual[1].IdentityReference | Should -Be "$env:USERDOMAIN\$testUser"

                $Actual = (Get-Acl -Path "$($testParams.Path)\FolderB").Access
                $Actual.Count | Should -BeExactly 2
                $Actual[0].IdentityReference | Should -Be 'BUILTIN\Administrators'
                $Actual[1].IdentityReference | Should -Be "$env:USERDOMAIN\$testUser2"
            } 
            It 'a Warning object for incorrect permissions is not created' {
                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'New'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                        [PSCustomObject]@{Path = 'FolderB\FolderC'; ACL = @{ } }
                    )
                }

                .$testScript @testParams | Where-Object Name -EQ 'Non inherited folder incorrect permissions' |
                Should -BeNullOrEmpty
            } 
        }
    }
    Context 'Fix' {
        It "create a FatalError object when the parent folder doesn't exist" {
            $testParams = @{
                Path   = 'NotExistingTestFolder'
                Action = 'Fix'
                Matrix = [PSCustomObject]@{Name = 'test' }
            }

            $Actual = .$testScript @testParams

            $Expected = [PSCustomObject]@{
                Type        = 'FatalError'
                Name        = 'Parent folder missing'
                Description = "The folder defined as 'Path' in the worksheet 'Settings' needs to be available on the remote machine. In case the folder structure needs to be created, please use 'Action=New' instead."
                Value       = $testParams.Path
            }

            $Actual.Type | Should -Be $Expected.Type
            $Actual.Name | Should -Be $Expected.Name
            $Actual.Description | Should -Be $Expected.Description
            $Actual.Value | Should -Be $Expected.Value
        } 
        Context 'folders in the matrix that are missing' {
            It 'are created' {
                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'Fix'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                        [PSCustomObject]@{Path = 'FolderB\FolderC'; ACL = @{ } }
                    )
                }
                $null = New-Item -Path $testParams.Path -ItemType Directory

                .$testScript @testParams

                "$($testParams.Path)\FolderA" | Should -Exist
                "$($testParams.Path)\FolderB\FolderC" | Should -Exist
            } 
            It 'are registered in a Warning object' {
                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'Fix'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                        [PSCustomObject]@{Path = 'FolderB\FolderC'; ACL = @{ } }
                    )
                }
                $null = New-Item -Path $testParams.Path -ItemType Directory

                $Actual = .$testScript @testParams | Where-Object Name -Like "*child folder*"

                $Actual.Type | Should -Be 'Warning'
                $Actual.Name | Should -Be 'Child folder created'
                        
                $Actual.Value[0] | Should -Be "$($testParams.Path)\FolderA"
                $Actual.Value[1] | Should -Be "$($testParams.Path)\FolderB\FolderC"
            } 
            It 'are not created when Path is set to Ignore' {
                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'Fix'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }; Ignore = $true }
                        [PSCustomObject]@{Path = 'FolderB\FolderC'; ACL = @{ } }
                    )
                }
                New-Item -Path $testParams.Path -ItemType Directory

                .$testScript @testParams
                        
                "$($testParams.Path)\FolderA" | Should -Not -Exist
                "$($testParams.Path)\FolderB\FolderC" | Should -Exist
            } 
        }
        Context 'incorrect folder permissions' {
            Context 'on non inherited folders' {
                It 'are corrected' {
                    $testParams = @{
                        Path   = $testParentFolder
                        Action = 'New'
                        Matrix = @(
                            [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                            [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser2 = 'R' } }
                        )
                    }

                    .$testScript @testParams

                    $testParams = @{
                        Path   = $testParentFolder
                        Action = 'Fix'
                        Matrix = @(
                            [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser2 = 'R' }; Parent = $true }
                            [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser = 'L' } }
                        )
                    }
                    
                    .$testScript @testParams

                    $Actual = (Get-Acl -Path $testParams.Path).Access
                    $Actual.Count | Should -BeExactly 2
                    $Actual[0].IdentityReference | Should -Be 'BUILTIN\Administrators'
                    $Actual[1].IdentityReference | Should -Be "$env:USERDOMAIN\$testUser2"

                    $Actual = (Get-Acl -Path ($testParams.Path + '\FolderA')).Access
                    $Actual.Count | Should -BeExactly 2
                    $Actual[0].IdentityReference | Should -Be 'BUILTIN\Administrators'
                    $Actual[1].IdentityReference | Should -Be "$env:USERDOMAIN\$testUser"
                } 
                Context 'are registered in a Warning object when' {
                    It 'DetailedLog is False only the folder name is saved' {
                        $testParams = @{
                            Path   = $testParentFolder
                            Action = 'Fix'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser2 = 'R' } }
                            )
                        }

                        New-Item -Path (Join-Path $testParentFolder '\FolderA') -ItemType Directory -Force

                        $Actual = .$testScript @testParams |
                        Where-Object Name -EQ $ExpectedIncorrectAclNonInheritedFolders.Name

                        $Actual.Type | Should -Be $ExpectedIncorrectAclNonInheritedFolders.Type
                        @(
                            $testParams.Path,
                            "$($testParams.Path)\FolderA"
                        ).ForEach( {
                                $Actual.Value | Should -Contain $_
                            })
                    } 
                    It 'DetailedLog is True the folder name, the old ACL and the new ACL are saved' {
                        $testParams = @{
                            Path        = $testParentFolder
                            Action      = 'Fix'
                            Matrix      = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser2 = 'R' } }
                            )
                            DetailedLog = $true
                        }

                        New-Item -Path (Join-Path $testParentFolder '\FolderA') -ItemType Directory -Force

                        $Actual = .$testScript @testParams |
                        Where-Object Name -EQ $ExpectedIncorrectAclNonInheritedFolders.Name

                        $Actual.Type | Should -Be $ExpectedIncorrectAclNonInheritedFolders.Type
                        $Actual.Value.Count | Should -BeExactly 2 -Because 'two folders have an incorrect ACL'

                        @(
                            $testParams.Path,
                            "$($testParams.Path)\FolderA"
                        ).ForEach( {
                                $Actual.Value.Keys | Should -Contain $_ -Because 'the folder FullName is expected'
                            })

                        $Actual.Value.GetEnumerator().ForEach( {
                                foreach ($v in @('old', 'new')) {
                                    $_.Value.$v | Should -Not -BeNullOrEmpty -Because 'an ACL is expected'
                                }
                            })
                    } 
                }
            }
            Context 'on inherited folders' {
                It 'are corrected' {
                    $testParams = @{
                        Path   = $testParentFolder
                        Action = 'New'
                        Matrix = @(
                            [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                            [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser2 = 'R' } }
                        )
                    }

                    .$testScript @testParams

                    $testParams = @{
                        Path   = $testParentFolder
                        Action = 'Fix'
                        Matrix = @(
                            [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                            [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                        )
                    }
                    
                    .$testScript @testParams

                    $Actual = (Get-Acl -Path "$($testParams.Path)\FolderA").Access
                    $Actual.IsInherited | Should -Not -Contain $false -Because 'IsInedited needs to be True on all Ace'
                } 
                Context 'are registered in a Warning object when' {
                    It 'DetailedLog is False only the folder name is saved' {
                        $testParams = @{
                            Path   = $testParentFolder
                            Action = 'New'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser2 = 'R' } }
                                [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' } }
                            )
                        }

                        .$testScript @testParams

                        $testParams = @{
                            Path   = $testParentFolder
                            Action = 'Fix'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                            )
                        }
                        
                        $Actual = .$testScript @testParams |
                        Where-Object Name -EQ $ExpectedIncorrectAclInheritedFolders.Name

                        $Actual.Type | Should -Be $ExpectedIncorrectAclInheritedFolders.Type
                        @(
                            ($testParams.Path + '\FolderA')
                            ($testParams.Path + '\FolderB')
                        ).ForEach( {
                                $Actual.Value | Should -Contain $_
                            })

                    } 
                    It 'DetailedLog is True the folder name, the old ACL and the new ACL are saved' {
                        $testParams = @{
                            Path   = $testParentFolder
                            Action = 'New'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser2 = 'R' } }
                                [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' } }
                            )
                        }

                        .$testScript @testParams

                        $testParams = @{
                            Path        = $testParentFolder
                            Action      = 'Fix'
                            Matrix      = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                            )
                            DetailedLog = $true
                        }
                        
                        $Actual = .$testScript @testParams |
                        Where-Object Name -EQ $ExpectedIncorrectAclInheritedFolders.Name

                        $Actual.Type | Should -Be $ExpectedIncorrectAclInheritedFolders.Type
                        @(
                            ($testParams.Path + '\FolderA')
                            ($testParams.Path + '\FolderB')
                        ).ForEach( {
                                $Actual.Value.Keys | Should -Contain $_
                            })

                        $Actual.Value.GetEnumerator().ForEach( {
                                $_.Value | Should -Not -BeNullOrEmpty -Because 'an ACL is expected'
                            })
                    } 
                }
            }
            Context "set the owner to 'BUILTIN\Administrators' when" {
                It 'the admin has access to all folders' {
                    $testParams = @{
                        Path   = $testParentFolder
                        Action = 'Fix'
                        Matrix = @(
                            [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                            [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                            [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' } }
                        )
                    }

                    New-Item -Path (Join-Path $testParams.Path '\FolderA') -ItemType Directory -Force
                    $testFolderPath = "$($testParams.Path)\FolderB"
                    $testFolder = New-Item -Path $testFolderPath -ItemType Directory -Force

                    #region Add ourselves as owner
                    $testOwner = [System.Security.Principal.NTAccount]"$env:USERDOMAIN\$env:USERNAME"
                    $testAcl = $testFolder.GetAccessControl()
                    $testAcl.SetOwner($testOwner)
                    $testFolder.SetAccessControl($testAcl)
                    #endregion

                    (Get-Acl -Path $testFolderPath).Owner | Should -Be "$env:USERDOMAIN\$env:USERNAME"

                    .$testScript @testParams
                    #Set-Location -Path $TestDrive

                    (Get-Acl -Path $testFolderPath).Owner | Should -Be 'BUILTIN\Administrators'
                } 
                It 'the admin has no access to the folder' {
                    $testParams = @{
                        Path   = $testParentFolder
                        Action = 'Fix'
                        Matrix = @(
                            [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                            [PSCustomObject]@{Path = 'Reports'; ACL = @{ } }
                            [PSCustomObject]@{Path = 'Reports\Fruits'; ACL = @{ } }
                            [PSCustomObject]@{Path = 'Reports\Fruits\Kiwi'; ACL = @{$testUser = 'R' } }
                        )
                    }
                    $testFolderPath = "$($testParams.Path)\Reports\Fruits\Kiwi"
                    $testFolder = New-Item -Path $testFolderPath -ItemType Directory -Force

                    #region Remove access
                    $testOwner = [System.Security.Principal.NTAccount]"$env:USERDOMAIN\$testUser"
                    $testAcl = $testFolder.GetAccessControl()
                    $testAcl.SetAccessRuleProtection($True, $False)
                    $testAcl.SetOwner($testOwner)
                    $testAcl.Access.ForEach( { $null = $testAcl.RemoveAccessRule($_) })
                    $testFolder.SetAccessControl($testAcl)
                    #endregion

                    (Get-Acl -Path $testFolderPath).Owner | 
                    Should -Be "$env:USERDOMAIN\$testUser"
                    (Get-Acl -Path $testFolderPath).Access | 
                    Should -BeNullOrEmpty

                    .$testScript @testParams
                    #Set-Location -Path $TestDrive

                    (Get-Acl -Path $testFolderPath).Owner | Should -Be 'BUILTIN\Administrators'
                    (Get-Acl -Path $testFolderPath).Access | Should -Not -BeNullOrEmpty
                } 
                It 'the admin has no access to the parent folder' {
                    $testParams = @{
                        Path   = $testParentFolder
                        Action = 'Fix'
                        Matrix = @(
                            [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                            [PSCustomObject]@{Path = 'Reports'; ACL = @{ } }
                            [PSCustomObject]@{Path = 'Reports\Fruits'; ACL = @{ } }
                            [PSCustomObject]@{Path = 'Reports\Fruits\Kiwi'; ACL = @{$testUser = 'R' } }
                        )
                    }
                    New-Item -Path "$($testParams.Path)\Reports\Fruits\Kiwi" -ItemType Directory -Force

                    $testFolderPath = "$($testParams.Path)\Reports\Fruits"
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
                    (Get-Acl -Path "$($testParams.Path)\Reports\Fruits\Kiwi").Access.IdentityReference |
                    Should -Not -Contain "$env:USERDOMAIN\$testUser"

                    .$testScript @testParams
                    #Set-Location -Path $TestDrive

                    (Get-Acl -Path $testFolderPath).Owner | Should -Be 'BUILTIN\Administrators'
                    (Get-Acl -Path "$($testParams.Path)\Reports\Fruits\Kiwi").Owner | Should -Be 'BUILTIN\Administrators'
                    (Get-Acl -Path "$($testParams.Path)\Reports\Fruits\Kiwi").Access.IdentityReference |
                    Should -Contain "$env:USERDOMAIN\$testUser"
                } 
            }
        }
        Context 'when the script is run again after Action Fix/New' {
            It 'the permissions are unchanged' {
                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'New'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                        [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' } }
                    )
                }

                $testPermissions = {
                    $Actual = (Get-Acl -Path $testParams.Path).Access
                    $Actual.Count | Should -BeExactly 2
                    $Actual[0].IdentityReference | Should -Be 'BUILTIN\Administrators'
                    $Actual[1].IdentityReference | Should -Be "$env:USERDOMAIN\$testUser"

                    $Actual = (Get-Acl -Path "$($testParams.Path)\FolderB").Access
                    $Actual.Count | Should -BeExactly 2
                    $Actual[0].IdentityReference | Should -Be 'BUILTIN\Administrators'
                    $Actual[1].IdentityReference | Should -Be "$env:USERDOMAIN\$testUser2"
                }

                .$testScript @testParams

                & $testPermissions

                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'Fix'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                        [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' } }
                    )
                }
                
                .$testScript @testParams

                & $testPermissions
            } 
            It 'nothing is reported as being incorrect' {
                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'New'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                        [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' } }
                    )
                }

                .$testScript @testParams

                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'Fix'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                        [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' } }
                    )
                }
                
                .$testScript @testParams | Where-Object { $_.Type -notmatch 'Information|Warning' } | Should -BeNullOrEmpty
            } 
        }
    }
    Context 'Check' {
        It "create a FatalError object when the parent folder doesn't exist" {
            $testParams = @{
                Path   = 'NotExistingTestFolder'
                Action = 'Check'
                Matrix = [PSCustomObject]@{Name = 'test' }
            }

            $Actual = .$testScript @testParams

            $Expected = [PSCustomObject]@{
                Type        = 'FatalError'
                Name        = 'Parent folder missing'
                Description = "The folder defined as 'Path' in the worksheet 'Settings' needs to be available on the remote machine. In case the folder structure needs to be created, please use 'Action=New' instead."
                Value       = $testParams.Path
            }

            $Actual.Type | Should -Be $Expected.Type
            $Actual.Name | Should -Be $Expected.Name
            $Actual.Description | Should -Be $Expected.Description
            $Actual.Value | Should -Be $Expected.Value
        } 
        Context 'folders in the matrix that are missing' {
            It 'are not created' {
                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'Check'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                        [PSCustomObject]@{Path = 'FolderB\FolderC'; ACL = @{ } }
                    )
                }
                $null = New-Item -Path $testParams.Path -ItemType Directory

                .$testScript @testParams

                $testParams.Path + '\' + $testParams.Matrix[1].Path | Should -Not -Exist
                $testParams.Path + '\' + $testParams.Matrix[2].Path | Should -Not -Exist
            } 
            It 'are registered in a Warning object' {
                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'Check'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                        [PSCustomObject]@{Path = 'FolderB\FolderC'; ACL = @{ } }
                    )
                }
                New-Item -Path $testParams.Path -ItemType Directory

                $Actual = .$testScript @testParams | Where-Object Name -Like "*child folder*"

                $Actual.Type | Should -Be 'Warning'
                $Actual.Name | Should -Be 'Child folder missing'
                $Actual.Value[0] | Should -Be "$($testParams.Path)\FolderA"
                $Actual.Value[1] | Should -Be "$($testParams.Path)\FolderB\FolderC"
            } 
            It 'are not checked when they are set to ignore' {
                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'Check'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ }; Ignore = $true }
                        [PSCustomObject]@{Path = 'FolderB\FolderC'; ACL = @{ } }
                    )
                }
                New-Item -Path $testParams.Path -ItemType Directory

                .$testScript @testParams | Where-Object { 
                    ($_.Name -Like "*child folder*") -and
                    ($_.Value -contains ($testParams.Path + '\FolderA')) } |
                Should -BeNullOrEmpty

                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'Check'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$env:USERNAME = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                        [PSCustomObject]@{Path = 'FolderB\FolderC'; ACL = @{ } }
                    )
                }
                #Set-Location -Path $TestDrive
                .$testScript @testParams | Where-Object { 
                    ($_.Name -Like "*child folder*") -and
                    ($_.Value -contains ($testParams.Path + '\FolderA')) } |
                Should -Not -BeNullOrEmpty
            } 
        }
        Context 'incorrect folder permissions' {
            Context 'on non inherited folders' {
                It 'are not corrected' {
                    $testParams = @{
                        Path   = $testParentFolder
                        Action = 'Check'
                        Matrix = @(
                            [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                            [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                            [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' } }
                        )
                    }

                    New-Item -Path (Join-Path $testParentFolder '\FolderA') -ItemType Directory -Force
                    New-Item -Path (Join-Path $testParentFolder '\FolderB') -ItemType Directory -Force

                    $Expected = @(
                        (Get-Acl -Path $testParams.Path)
                        (Get-Acl -Path "$($testParams.Path)\FolderA")
                        (Get-Acl -Path "$($testParams.Path)\FolderB")
                    )

                    .$testScript @testParams

                    $Actual = @(
                        (Get-Acl -Path $testParams.Path)
                        (Get-Acl -Path "$($testParams.Path)\FolderA")
                        (Get-Acl -Path "$($testParams.Path)\FolderB")
                    )

                    for ($i = 0; $i -lt $Expected.Count; $i++) {
                        $AssertParams = @{
                            Actual   = $Actual[$i].Access | 
                            Sort-Object IdentityReference
                            Expected = $Expected[$i].Access | 
                            Sort-Object IdentityReference
                        }
                        Assert-Equivalent @AssertParams

                        $AssertParams = @{
                            Actual   = $Actual[$i].Owner | 
                            Sort-Object IdentityReference
                            Expected = $Expected[$i].Owner | 
                            Sort-Object IdentityReference
                        }
                        Assert-Equivalent @AssertParams
                    }
                } 
                Context 'are registered in a Warning object when' {
                    It 'DetailedLog is False only the folder name is saved' {
                        $testParams = @{
                            Path   = $testParentFolder
                            Action = 'Check'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser2 = 'R' } }
                            )
                        }

                        New-Item -Path (Join-Path $testParentFolder '\FolderA') -ItemType Directory -Force

                        $Actual = .$testScript @testParams |
                        Where-Object Name -EQ $ExpectedIncorrectAclNonInheritedFolders.Name

                        $Actual.Type | Should -Be $ExpectedIncorrectAclNonInheritedFolders.Type
                        @(
                            ($testParams.Path),
                            ($testParams.Path + '\FolderA')
                        ).ForEach( {
                                $Actual.Value | Should -Contain $_
                            })

                    } 
                    It 'DetailedLog is True the folder name, the old ACL and the new ACL are saved' {
                        $testParams = @{
                            Path        = $testParentFolder
                            Action      = 'Check'
                            Matrix      = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser2 = 'R' } }
                            )
                            DetailedLog = $true
                        }

                        New-Item -Path (Join-Path $testParentFolder '\FolderA') -ItemType Directory -Force

                        $Actual = .$testScript @testParams |
                        Where-Object Name -EQ $ExpectedIncorrectAclNonInheritedFolders.Name

                        $Actual.Type | Should -Be $ExpectedIncorrectAclNonInheritedFolders.Type
                        $Actual.Value.Count | Should -BeExactly 2 -Because 'two folders have an incorrect ACL'

                        @(
                            ($testParams.Path),
                            ($testParams.Path + '\FolderA')
                        ).ForEach( {
                                $Actual.Value.Keys | Should -Contain $_ -Because 'the folder FullName is expected'
                            })

                        $Actual.Value.GetEnumerator().ForEach( {
                                foreach ($v in @('old', 'new')) {
                                    $_.Value.$v | Should -Not -BeNullOrEmpty -Because 'an ACL is expected'
                                }
                            })

                    } 
                }
            }
            Context 'on inherited folders' {
                It 'are not corrected' {
                    $testParams = @{
                        Path   = $testParentFolder
                        Action = 'New'
                        Matrix = @(
                            [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                            [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser2 = 'R' } }
                        )
                    }

                    .$testScript @testParams

                    $Expected = @(
                        (Get-Acl -Path ($testParams.Path))
                        (Get-Acl -Path ($testParams.Path + '\FolderA'))
                    )

                    $testParams = @{
                        Path   = $testParentFolder
                        Action = 'Check'
                        Matrix = @(
                            [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                            [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                        )
                    }
                    
                    .$testScript @testParams

                    $Actual = @(
                        (Get-Acl -Path ($testParams.Path))
                        (Get-Acl -Path ($testParams.Path + '\FolderA'))
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
                } 
                Context 'are registered in a Warning object when' {
                    It 'DetailedLog is False only the folder name is saved' {
                        $testParams = @{
                            Path   = $testParentFolder
                            Action = 'New'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser2 = 'R' } }
                                [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' } }
                            )
                        }

                        .$testScript @testParams

                        $testParams = @{
                            Path   = $testParentFolder
                            Action = 'Check'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                            )
                        }
                        
                        $Actual = .$testScript @testParams |
                        Where-Object Name -EQ $ExpectedIncorrectAclInheritedFolders.Name

                        $Actual.Type | Should -Be $ExpectedIncorrectAclInheritedFolders.Type
                        @(
                            ($testParams.Path + '\FolderA')
                            ($testParams.Path + '\FolderB')
                        ).ForEach( {
                                $Actual.Value | Should -Contain $_
                            })
                    } 
                    It 'DetailedLog is True the folder name, the old ACL and the new ACL are saved' {
                        $testParams = @{
                            Path   = $testParentFolder
                            Action = 'New'
                            Matrix = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{$testUser2 = 'R' } }
                                [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' } }
                            )
                        }

                        .$testScript @testParams

                        $testParams = @{
                            Path        = $testParentFolder
                            Action      = 'Check'
                            Matrix      = @(
                                [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                                [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                            )
                            DetailedLog = $true
                        }
                        
                        $Actual = .$testScript @testParams |
                        Where-Object Name -EQ $ExpectedIncorrectAclInheritedFolders.Name

                        $Actual.Type | Should -Be $ExpectedIncorrectAclInheritedFolders.Type
                        @(
                            ($testParams.Path + '\FolderA')
                            ($testParams.Path + '\FolderB')
                        ).ForEach( {
                                $Actual.Value.Keys | Should -Contain $_
                            })

                        $Actual.Value.GetEnumerator().ForEach( {
                                $_.Value | Should -Not -BeNullOrEmpty -Because 'an ACL is expected'
                            })
                    } 
                }
            }
            It "don't report missing folders as having incorrect permissions" {
                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'Check'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                        [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' } }
                        [PSCustomObject]@{Path = 'FolderC'; ACL = @{ } }
                    )
                }

                New-Item -Path ($testParams.Path + '\FolderA') -ItemTyp Directory -Force

                $actual = .$testScript @testParams | Where-Object {
                    ($_.Name -EQ $ExpectedIncorrectAclNonInheritedFolders.Name) -or 
                    ($_.Name -EQ $ExpectedIncorrectAclInheritedFolders.Name)
                }

                @(
                    ($testParams.Path + '\FolderB')
                    ($testParams.Path + '\FolderC')
                ).ForEach( {
                        $actual.Value | Should -Not -Contain $_
                    })
            } 
        }
        Context 'incorrect file permissions' {
            It 'are not corrected' {
                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'Check'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                        [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' } }
                    )
                }

                $testFile = New-Item -Path ($testParams.Path + '\FolderB\File.txt') -ItemTyp File -Force
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

                .$testScript @testParams

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
            } 
            Context 'are registered in a Warning object when' {
                BeforeAll {
                    $Expected = [PSCustomObject]@{
                        Type = 'Warning'
                        Name = 'Inherited permissions incorrect'
                    }
                }
                It 'DetailedLog is False only the file name is saved' {
                    $testParams = @{
                        Path   = $testParentFolder
                        Action = 'Check'
                        Matrix = @(
                            [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                            [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                            [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' } }
                        )
                    }

                    $testFile = New-Item -Path ($testParams.Path + '\FolderB\File.txt') -ItemTyp File -Force
                    $testFileItem = Get-Item $testFile

                    #region Remove access and set owner
                    $testOwner = [System.Security.Principal.NTAccount]"$env:USERDOMAIN\$testUser"
                    $testAcl = $testFileItem.GetAccessControl()
                    $testAcl.SetAccessRuleProtection($True, $False)
                    $testAcl.SetOwner($testOwner)
                    $testAcl.Access.ForEach( { $null = $testAcl.RemoveAccessRule($_) })
                    $testFileItem.SetAccessControl($testAcl)
                    #endregion

                    $Actual = .$testScript @testParams |
                    Where-Object Name -EQ $Expected.Name

                    $testFile.FullName | Should -Be $actual.Value
                }
                It 'DetailedLog is True the file name and the the old ACL are saved' {
                    $testParams = @{
                        Path        = $testParentFolder
                        Action      = 'Check'
                        Matrix      = @(
                            [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                            [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                            [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' } }
                        )
                        DetailedLog = $true
                    }

                    $testFile = New-Item -Path ($testParams.Path + '\FolderB\File.txt') -ItemTyp File -Force
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

                    $Actual = .$testScript @testParams | Where-Object Name -EQ $Expected.Name

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
                } 
            }
        }
        Context 'when the script is run again after Action Fix/New' {
            It 'the permissions are unchanged' {
                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'New'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                        [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' } }
                    )
                }

                $testPermissions = {
                    $Actual = (Get-Acl -Path $testParams.Path).Access
                    $Actual.Count | Should -BeExactly 2
                    $Actual[0].IdentityReference | Should -Be 'BUILTIN\Administrators'
                    $Actual[1].IdentityReference | Should -Be "$env:USERDOMAIN\$testUser"

                    $Actual = (Get-Acl -Path "$($testParams.Path)\FolderB").Access
                    $Actual.Count | Should -BeExactly 2
                    $Actual[0].IdentityReference | Should -Be 'BUILTIN\Administrators'
                    $Actual[1].IdentityReference | Should -Be "$env:USERDOMAIN\$testUser2"
                }

                .$testScript @testParams

                & $testPermissions

                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'Fix'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                        [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' } }
                    )
                }
                        
                .$testScript @testParams

                & $testPermissions

                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'Check'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                        [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' } }
                    )
                }
                        
                .$testScript @testParams

                & $testPermissions
            } 
            It 'nothing is reported as being incorrect' {
                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'New'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                        [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' } }
                    )
                }

                .$testScript @testParams

                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'Fix'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                        [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' } }
                    )
                }
                
                .$testScript @testParams | Where-Object { $_.Type -notmatch 'Information|Warning' } | Should -BeNullOrEmpty

                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'Check'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                        [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' } }
                    )
                }
                
                .$testScript @testParams | Where-Object { $_.Type -notmatch 'Information|Warning' } | Should -BeNullOrEmpty
            } 
        }
        Context 'create a Warning object for inaccessible data when' {
            It 'files are found in the deepest folder of a matrix list only path' {
                $testParams = @{
                    Path   = $testParentFolder
                    Action = 'Check'
                    Matrix = @(
                        [PSCustomObject]@{Path = 'Path'; ACL = @{$testUser = 'L' }; Parent = $true }
                        [PSCustomObject]@{Path = 'FolderA'; ACL = @{ } }
                        [PSCustomObject]@{Path = 'FolderB'; ACL = @{$testUser2 = 'R' } }
                        [PSCustomObject]@{Path = 'FolderB\SubfolderB1'; ACL = @{$testUser2 = 'L' } }
                        [PSCustomObject]@{Path = 'FolderB\SubfolderB2'; ACL = @{$testUser2 = 'W' } }
                    )
                }

                $TestFile = New-Item -Path ($testParams.Path + '\FolderB\SubfolderB1\File.txt') -ItemTyp File -Force

                $Actual = .$testScript @testParams | Where-Object Name -EQ $ExpectedInaccessibleData.Name

                @(
                    $TestFile.FullName
                ).ForEach( {
                        $Actual.Value | Should -Contain $_ -Because 'the deepest folder has only list permissions'
                    })
            } -Skip
        }
    }
}

