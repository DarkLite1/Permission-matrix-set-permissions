#Requires -Version 5.1
#Requires -Modules Assert
#Requires -Modules Pester
#Requires -Modules SmbShare

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'

$Skip = $false

Describe $sut {
    Context 'the mandatory parameters are' {
        $TestCases = @('Path', 'Flag').ForEach( {@{Name = $_}})

        it "<Name>" -TestCases $TestCases {
            Param (
                $Name
            )
            (Get-Command "$here\$sut").Parameters[$Name].Attributes.Mandatory | Should -Be $true
        } -Skip:$Skip
    }

    In $TestDrive {
        Context 'the script needs to be run with' {
            it 'adminstrator privileges' {
                Function Test-IsAdmin {}
                Mock Test-IsAdminHC {$false}

                $Expected = [PSCustomObject]@{
                    Type        = 'FatalError'
                    Name        = 'Administrator privileges'
                    Description = "Administrator privileges are required to be able to apply permissions."
                    Value       = "SamAccountName '$env:USERNAME'"
                }

                $Actual = ."$here/$sut" -Path 'NotExistingNotImportant' -Flag $true

                Assert-Equivalent -Actual $Actual -Expected $Expected
            } -Skip:$Skip

            Mock Test-IsAdminHC {$true}

            it 'at least PowerShell 5.1' {
                Function Test-IsRequiredPowerShellVersionHC {}
                Mock Test-IsRequiredPowerShellVersionHC {$false}

                $Expected = [PSCustomObject]@{
                    Type        = 'FatalError'
                    Name        = 'PowerShell version'
                    Description = "PowerShell version 5.1 or higher is required to be able to use advanced methods."
                    Value       = "PowerShell $($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor)"
                }

                $Actual = ."$here/$sut" -Path 'NotExistingNotImportant' -Flag $true | Where-Object {$_.Name -eq $Expected.Name}

                Assert-Equivalent -Actual $Actual -Expected $Expected
            } -Skip:$Skip

            Mock Test-IsRequiredPowerShellVersionHC {$true}

            it 'at least .NET 4.6.2' {
                Mock -CommandName Get-ItemPropertyValue -MockWith {
                    379893
                } -ParameterFilter {
                    $Name -eq 'Release'
                }

                $Expected = [PSCustomObject]@{
                    Type        = 'FatalError'
                    Name        = '.NET Framework version'
                    Description = "Microsoft .NET Framework version 4.6.2 or higher is required to be able to traverse long path names and use advanced PowerShell methods."
                    Value       = $null
                }

                $Actual = ."$here/$sut" -Path 'NotExistingNotImportant' -Flag $true | Where-Object {$_.Name -eq $Expected.Name}

                Assert-Equivalent -Actual $Actual -Expected $Expected
            } -Skip:$Skip

            Mock -CommandName Get-ItemPropertyValue -MockWith {461808}
        }
        
        #region Create first shared folder
        $testFolderName = 'TestFolder'
        $testFolderName2 = 'TestFolder2'
        Remove-SmbShare -Name $testFolderName -Force -EA Ignore
        Remove-SmbShare -Name $testFolderName2 -Force -EA Ignore

        $testDirItem = New-Item -Path ./testShare -ItemType Directory
        New-SmbShare -Name $testFolderName -Path $testDirItem
        Grant-SmbShareAccess -Name $testFolderName -AccountName Everyone -AccessRight Full -Force
        #endregion

        Context 'set the Access Based Enumeration flag' {
            it "to enabled when the 'Flag' parameter is set to TRUE" {
                Set-SmbShare -Name $testFolderName -FolderEnumerationMode Unrestricted -Force

                ."$here/$sut" -Path $testDirItem -Flag $true

                (Get-SmbShare -Name $testFolderName).FolderEnumerationMode | 
                    Should -BeExactly 'AccessBased'
                Test-AccessBasedEnumerationHC -Name $testFolderName | Should -BeTrue
            } -Skip:$Skip
            it "to disabled when the 'Flag' parameter is set to FALSE" {
                Set-SmbShare -Name $testFolderName -FolderEnumerationMode AccessBased -Force

                ."$here/$sut" -Path $testDirItem -Flag $false

                (Get-SmbShare -Name $testFolderName).FolderEnumerationMode | 
                    Should -BeExactly 'Unrestricted'
                Test-AccessBasedEnumerationHC -Name $testFolderName | Should -BeFalse
            } -Skip:$Skip

            #region Create second shared folder
            $testDirItem2 = New-Item -Path ./testShare2 -ItemType Directory
            New-SmbShare -Name $testFolderName2 -Path $testDirItem2
            Grant-SmbShareAccess -Name $testFolderName2 -AccountName Everyone -AccessRight Full -Force
            #endregion

            it 'only on the requested folder, not on other folders' {
                Set-SmbShare -Name $testFolderName -FolderEnumerationMode Unrestricted -Force
                Set-SmbShare -Name $testFolderName2 -FolderEnumerationMode Unrestricted -Force

                ."$here/$sut" -Path $testDirItem -Flag $true

                (Get-SmbShare -Name $testFolderName).FolderEnumerationMode | 
                    Should -BeExactly 'AccessBased' -Because 'we enabled ABE on this folder'
                (Get-SmbShare -Name $testFolderName2).FolderEnumerationMode | 
                    Should -BeExactly 'Unrestricted' -Because "we didn't enable ABE on this folder"
            } -Skip:$Skip
            it 'on multiple folders and ignore duplicates' {
                Set-SmbShare -Name $testFolderName -FolderEnumerationMode Unrestricted -Force
                Set-SmbShare -Name $testFolderName2 -FolderEnumerationMode Unrestricted -Force

                ."$here/$sut" -Path $testDirItem, $testDirItem2, $testDirItem -Flag $true

                (Get-SmbShare -Name $testFolderName).FolderEnumerationMode | 
                    Should -BeExactly 'AccessBased'
                (Get-SmbShare -Name $testFolderName2).FolderEnumerationMode | 
                    Should -BeExactly 'AccessBased'
            } -Skip:$Skip
            it 'on multiple folders and return the results' {
                Set-SmbShare -Name $testFolderName -FolderEnumerationMode Unrestricted -Force
                Set-SmbShare -Name $testFolderName2 -FolderEnumerationMode Unrestricted -Force

                $Expected = [PSCustomObject]@{
                    Type        = 'Warning'
                    Name        = 'Access Based Enumeration'
                    Description = "Access Based Enumeration should be set to '$true'. This will hide files and folders where the users don't have access to. We fixed this now."
                    Value       = @{
                        $testFolderName  = $testDirItem.FullName
                        $testFolderName2 = $testDirItem2.FullName
                    }
                }
                
                $Actual = ."$here/$sut" -Path $testDirItem, $testDirItem2, $testDirItem -Flag $true

                Assert-Equivalent -Actual ($Actual | Where Name -eq 'Access Based Enumeration') -Expected $Expected
            } -Skip:$Skip
        }

        Context "Set share permissions to 'FullControl for Administrators' and 'Read & Executed for Authenticated users'" {
            it 'when they are incorrect' {
                Grant-SmbShareAccess -Name $testFolderName -AccountName Administrators -AccessRight Full –Force
                Grant-SmbShareAccess -Name $testFolderName -AccountName Everyone -AccessRight Read –Force
                Grant-SmbShareAccess -Name $testFolderName -AccountName 'Authenticated users' -AccessRight Read –Force

                $Result = ."$here\$sut" -Path $testDirItem -Flag $true

                $Actual = Get-SmbShareAccess -Name $testFolderName

                #region Verify share permissions
                $Actual.Count | Should -BeExactly 2

                $Actual.ForEach( {
                        $_.Name | Should -Be $testFolderName
                        $_.AccessControlType | Should -Be 'Allow'
                    })
                ($Actual | where AccountName -EQ 'NT AUTHORITY\Authenticated Users').AccessRight | Should -Be 'Change'
                ($Actual | where AccountName -EQ 'BUILTIN\Administrators').AccessRight | Should -Be 'Full'
                #endregion

                #verify Script output
                $Expected = [PSCustomObject]@{
                    Type        = 'Warning'
                    Name        = 'Share permissions'
                    Description = "The share permissions are now set to 'Administrators: FullControl' and 'Authenticated users: Change'. The effective permissions are managed on NTFS level."
                    Value       = @{$testFolderName = @{
                            'NT AUTHORITY\Authenticated Users' = 'Read'
                            'Everyone'                         = 'Read'
                            'BUILTIN\Administrators'           = 'FullControl'
                        }
                    }
                }
                
                Assert-Equivalent -Actual ($Result | where Name -EQ $Expected.Name) -Expected $Expected
                #endregion
            } -Skip:$Skip
            it "but don't change anything when they are already correct" {
                Remove-SmbShare -Name $testFolderName -Force -EA Ignore
                New-SmbShare -Name $testFolderName -Path $testDirItem
                Grant-SmbShareAccess -Name $testFolderName -AccountName Administrators -AccessRight Full -Force
                Grant-SmbShareAccess -Name $testFolderName -AccountName 'Authenticated Users' -AccessRight Change -Force
                Set-SmbShare -Name $testFolderName -FolderEnumerationMode AccessBased -Force

                $Result = ."$here\$sut" -Path $testDirItem -Flag $true

                $Actual = Get-SmbShareAccess -Name $testFolderName | where Name -EQ 'Share permissions' | Should -BeNullOrEmpty
            } -Skip:$Skip
            it "except when there's no shared folder, do nothing" {
                Remove-SmbShare -Name $testFolderName -Force -EA Ignore

                $Result = ."$here\$sut" -Path $testDirItem -Flag $true

                ($Result | where Action -EQ ACL) | Should -BeNullOrEmpty
                #endregion
            } -Skip:$Skip
        }
        
        Remove-SmbShare -Name $testFolderName -Force -EA Ignore
        Remove-SmbShare -Name $testFolderName2 -Force -EA Ignore
    }
}