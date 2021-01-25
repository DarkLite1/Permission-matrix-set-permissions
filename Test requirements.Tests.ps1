#Requires -Version 5.1
#Requires -Modules Pester, SmbShare, Assert

BeforeAll {
    $testSmbShareName = 'TestFolder'
    $testSmbShareName2 = 'TestFolder2'

    Remove-SmbShare -Name $testSmbShareName -Force -EA Ignore
    Remove-SmbShare -Name $testSmbShareName2 -Force -EA Ignore

    $testDirItem = New-Item -Path 'TestDrive:\testShare' -ItemType Directory
    New-SmbShare -Name $testSmbShareName -Path $testDirItem
    Grant-SmbShareAccess -Name $testSmbShareName -AccountName Everyone -AccessRight Full -Force

    $testDirItem2 = New-Item -Path 'TestDrive:\testShare2' -ItemType Directory
    New-SmbShare -Name $testSmbShareName2 -Path $testDirItem2
    Grant-SmbShareAccess -Name $testSmbShareName2 -AccountName Everyone -AccessRight Full -Force

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')

    Function Test-IsRequiredPowerShellVersionHC {}
    
    Mock Get-ItemPropertyValue -MockWith { 461808 }
    Mock Test-IsAdminHC { $true }
    Mock Test-IsRequiredPowerShellVersionHC { $true }
}
AfterAll {
    Remove-SmbShare -Name $testSmbShareName -Force -EA Ignore
    Remove-SmbShare -Name $testSmbShareName2 -Force -EA Ignore
}

Describe 'the mandatory parameters are' {
    It "<_>" -TestCases @('Path', 'Flag') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory | 
        Should -BeTrue
    } 
}
Context 'when the script' {
    Context 'is not started with administrator privileges' {
        AfterAll {
            Mock Test-IsAdminHC { $true }
        }
        It 'return a FatalError object' {
            Mock Test-IsAdminHC { $false }

            $Expected = [PSCustomObject]@{
                Type        = 'FatalError'
                Name        = 'Administrator privileges'
                Description = "Administrator privileges are required to be able to apply permissions."
                Value       = "SamAccountName '$env:USERNAME'"
            }

            $Actual = .$testScript -Path 'NotExistingNotImportant' -Flag $true

            Assert-Equivalent -Actual $Actual -Expected $Expected
        }
    }
    Context 'cannot find PowerShell 5.1 or later' {
        AfterAll {
            # Function Test-IsRequiredPowerShellVersionHC {}
            Mock Test-IsRequiredPowerShellVersionHC { $true }
        }
        It 'return a FatalError object' {
            Mock Test-IsRequiredPowerShellVersionHC { $false }
            $Expected = [PSCustomObject]@{
                Type        = 'FatalError'
                Name        = 'PowerShell version'
                Description = "PowerShell version 5.1 or higher is required to be able to use advanced methods."
                Value       = "PowerShell $($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor)"
            }
            
            $Actual = .$testScript -Path 'NotExistingNotImportant' -Flag $true | Where-Object { $_.Name -eq $Expected.Name }
            
            Assert-Equivalent -Actual $Actual -Expected $Expected
        }
    }
    Context 'cannot find .NET 4.6.2 or later' {
        AfterAll {
            Mock -CommandName Get-ItemPropertyValue -MockWith { 461808 }
        }
        It 'return a FatalError object' {
            Mock -CommandName Get-ItemPropertyValue -MockWith {
                379893
            } -ParameterFilter { $Name -eq 'Release' }

            $Expected = [PSCustomObject]@{
                Type        = 'FatalError'
                Name        = '.NET Framework version'
                Description = "Microsoft .NET Framework version 4.6.2 or higher is required to be able to traverse long path names and use advanced PowerShell methods."
                Value       = $null
            }

            $Actual = .$testScript -Path 'NotExisting' -Flag $true | 
            Where-Object { $_.Name -eq $Expected.Name }

            Assert-Equivalent -Actual $Actual -Expected $Expected
        }
    } 
} 
Context 'set the Access Based Enumeration flag' {
    It "to enabled when the 'Flag' parameter is set to TRUE" {
        Set-SmbShare -Name $testSmbShareName -FolderEnumerationMode Unrestricted -Force

        .$testScript -Path $testDirItem -Flag $true

        (Get-SmbShare -Name $testSmbShareName).FolderEnumerationMode | 
        Should -BeExactly 'AccessBased'
        Test-AccessBasedEnumerationHC -Name $testSmbShareName | Should -BeTrue
    } 
    It "to disabled when the 'Flag' parameter is set to FALSE" {
        Set-SmbShare -Name $testSmbShareName -FolderEnumerationMode AccessBased -Force

        .$testScript -Path $testDirItem -Flag $false

        (Get-SmbShare -Name $testSmbShareName).FolderEnumerationMode | 
        Should -BeExactly 'Unrestricted'
        Test-AccessBasedEnumerationHC -Name $testSmbShareName | Should -BeFalse
    } 
    It 'only on the requested folder, not on other folders' {
        Set-SmbShare -Name $testSmbShareName -FolderEnumerationMode Unrestricted -Force
        Set-SmbShare -Name $testSmbShareName2 -FolderEnumerationMode Unrestricted -Force

        .$testScript -Path $testDirItem -Flag $true

        (Get-SmbShare -Name $testSmbShareName).FolderEnumerationMode | 
        Should -BeExactly 'AccessBased' -Because 'we enabled ABE on this folder'
        (Get-SmbShare -Name $testSmbShareName2).FolderEnumerationMode | 
        Should -BeExactly 'Unrestricted' -Because "we didn't enable ABE on this folder"
    } 
    It 'on multiple folders and ignore duplicates' {
        Set-SmbShare -Name $testSmbShareName -FolderEnumerationMode Unrestricted -Force
        Set-SmbShare -Name $testSmbShareName2 -FolderEnumerationMode Unrestricted -Force

        .$testScript -Path $testDirItem, $testDirItem2, $testDirItem -Flag $true

        (Get-SmbShare -Name $testSmbShareName).FolderEnumerationMode | 
        Should -BeExactly 'AccessBased'
        (Get-SmbShare -Name $testSmbShareName2).FolderEnumerationMode | 
        Should -BeExactly 'AccessBased'
    } 
    It 'on multiple folders and return the results' {
        Set-SmbShare -Name $testSmbShareName -FolderEnumerationMode Unrestricted -Force
        Set-SmbShare -Name $testSmbShareName2 -FolderEnumerationMode Unrestricted -Force

        $Expected = [PSCustomObject]@{
            Type        = 'Warning'
            Name        = 'Access Based Enumeration'
            Description = "Access Based Enumeration should be set to '$true'. This will hide files and folders where the users don't have access to. We fixed this now."
            Value       = @{
                $testSmbShareName  = $testDirItem.FullName
                $testSmbShareName2 = $testDirItem2.FullName
            }
        }
                
        $Actual = .$testScript -Path $testDirItem, $testDirItem2, $testDirItem -Flag $true

        Assert-Equivalent -Actual ($Actual | Where-Object Name -EQ 'Access Based Enumeration') -Expected $Expected
    } 
}
Context "Set share permissions to 'FullControl for Administrators' and 'Read & Executed for Authenticated users'" {
    It 'when they are incorrect' {
        Grant-SmbShareAccess -Name $testSmbShareName -AccountName Administrators -AccessRight Full –Force
        Grant-SmbShareAccess -Name $testSmbShareName -AccountName Everyone -AccessRight Read –Force
        Grant-SmbShareAccess -Name $testSmbShareName -AccountName 'Authenticated users' -AccessRight Read –Force

        $Result = .$testScript -Path $testDirItem -Flag $true

        $Actual = Get-SmbShareAccess -Name $testSmbShareName

        #region Verify share permissions
        $Actual.Count | Should -BeExactly 2

        $Actual.ForEach( {
                $_.Name | Should -Be $testSmbShareName
                $_.AccessControlType | Should -Be 'Allow'
            })
        ($Actual | Where-Object AccountName -EQ 'NT AUTHORITY\Authenticated Users').AccessRight | Should -Be 'Change'
        ($Actual | Where-Object AccountName -EQ 'BUILTIN\Administrators').AccessRight | Should -Be 'Full'
        #endregion

        #verify Script output
        $Expected = [PSCustomObject]@{
            Type        = 'Warning'
            Name        = 'Share permissions'
            Description = "The share permissions are now set to 'Administrators: FullControl' and 'Authenticated users: Change'. The effective permissions are managed on NTFS level."
            Value       = @{$testSmbShareName = @{
                    'NT AUTHORITY\Authenticated Users' = 'Read'
                    'Everyone'                         = 'Read'
                    'BUILTIN\Administrators'           = 'FullControl'
                }
            }
        }
                
        Assert-Equivalent -Actual ($Result | Where-Object Name -EQ $Expected.Name) -Expected $Expected
        #endregion
    } 
    It "but don't change anything when they are already correct" {
        Remove-SmbShare -Name $testSmbShareName -Force -EA Ignore
        New-SmbShare -Name $testSmbShareName -Path $testDirItem
        Grant-SmbShareAccess -Name $testSmbShareName -AccountName Administrators -AccessRight Full -Force
        Grant-SmbShareAccess -Name $testSmbShareName -AccountName 'Authenticated Users' -AccessRight Change -Force
        Set-SmbShare -Name $testSmbShareName -FolderEnumerationMode AccessBased -Force

        $Result = .$testScript -Path $testDirItem -Flag $true

        $Actual = Get-SmbShareAccess -Name $testSmbShareName | Where-Object Name -EQ 'Share permissions' | Should -BeNullOrEmpty
    } 
    It "except when there's no shared folder, do nothing" {
        Remove-SmbShare -Name $testSmbShareName -Force -EA Ignore

        $Result = .$testScript -Path $testDirItem -Flag $true

        ($Result | Where-Object Action -EQ ACL) | Should -BeNullOrEmpty
        #endregion
    } 
}
        

