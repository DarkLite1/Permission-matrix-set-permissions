#Requires -Version 5.1
#Requires -Modules Pester, SmbShare

BeforeAll {
    $testSmbShare = @(
        @{
            Name = 'testShare1'
            Path = (New-Item -Path 'TestDrive:\s1' -ItemType Directory).FullName
        }
        @{
            Name = 'testShare2'
            Path = (New-Item -Path 'TestDrive:\s2' -ItemType Directory).FullName
        }
    )

    $testSmbShare.ForEach(
        { New-SmbShare -Name $_.Name -Path $_.Path }
    )

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')

    Mock Get-ItemPropertyValue -MockWith { 461808 }
    Mock Test-IsAdminHC { $true }
    Mock Write-Warning
}
AfterAll {
    $testSmbShare.ForEach(
        { Remove-SmbShare -Name $_.Name -Force -EA Ignore }
    )
}

Describe 'the mandatory parameters are' {
    It "<_>" -TestCases @('Path', 'Flag') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'return a FatalError object when' {
    It 'the script is not started with administrator privileges' {
        Mock Test-IsAdminHC { $false }

        $expected = [PSCustomObject]@{
            Type        = 'FatalError'
            Name        = 'Administrator privileges'
            Description = "Administrator privileges are required to be able to apply permissions."
            Value       = "SamAccountName '$env:USERNAME'"
        }

        $actual = .$testScript -Path 'NotExistingNotImportant' -Flag $true

        $actual | ConvertTo-Json |
        Should -BeExactly ($expected | ConvertTo-Json)
    }
    It 'the minimal version fo PowerShell is not installed' {
        $expected = [PSCustomObject]@{
            Type        = 'FatalError'
            Name        = 'PowerShell version'
            Description = "PowerShell version 999.33 or higher is required."
            Value       = "PowerShell $($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor)"
        }

        $actual = .$testScript -Path 'NotExistingNotImportant' -Flag $true -MinimumPowerShellVersion @{
            Major = 999
            Minor = 33
        }

        $actual | Where-Object { $_.Name -eq $expected.Name } | ConvertTo-Json |
        Should -BeExactly ($expected | ConvertTo-Json)
    }
    It '.NET 4.6.2 or later is not installed' {
        Mock -CommandName Get-ItemPropertyValue -MockWith {
            379893
        } -ParameterFilter { $Name -eq 'Release' }

        $expected = [PSCustomObject]@{
            Type        = 'FatalError'
            Name        = '.NET Framework version'
            Description = "Microsoft .NET Framework version 4.6.2 or higher is required to be able to traverse long path names and use advanced PowerShell methods."
            Value       = $null
        }

        $actual = .$testScript -Path 'NotExisting' -Flag $true |
        Where-Object { $_.Name -eq $expected.Name }

        $actual | ConvertTo-Json |
        Should -BeExactly ($expected | ConvertTo-Json)
    }
}
Describe 'when the smb share permissions are' {
    Context 'incorrect' {
        BeforeAll {
            @(
                @{
                    AccountName = 'Administrators'
                    AccessRight = 'Full'
                }
                @{
                    AccountName = 'Everyone'
                    AccessRight = 'Read'
                }
                @{
                    AccountName = 'Authenticated users'
                    AccessRight = 'Read'
                }
            ).ForEach(
                {
                    $testGrantParams = $_
                    Grant-SmbShareAccess -Name $testSmbShare[0].Name @testGrantParams -Force
                }
            )

            $Result = .$testScript -Path $testSmbShare[0].Path -Flag $true

            $actual = Get-SmbShareAccess -Name $testSmbShare[0].Name
        }
        Context 'set the smb share permissions to' {
            It 'BUILTIN\Administrators: FullControl' {
                $a = $actual.Where(
                    { $_.AccountName -eq 'BUILTIN\Administrators' }
                )

                $a.AccessRight | Should -BeExactly 'Full'
                $a.AccessControlType | Should -BeExactly 'Allow'
            }
            It 'NT AUTHORITY\Authenticated Users: Change' {
                $a = $actual.Where(
                    { $_.AccountName -eq 'NT AUTHORITY\Authenticated Users' }
                )

                $a.AccessRight | Should -BeExactly 'Change'
                $a.AccessControlType | Should -BeExactly 'Allow'
            }
            It 'with no other permissions' {
                $actual | Should -HaveCount 2
            }
        }
        Context 'return an object of type' {
            It 'Warning' {
                $expected = [PSCustomObject]@{
                    Type        = 'Warning'
                    Name        = 'Share permissions'
                    Description = "The share permissions are now set to 'Administrators: FullControl' and 'Authenticated users: Change'. The effective permissions are managed on NTFS level."
                    Value       = @{$testSmbShare[0].Name = @{
                            'BUILTIN\Administrators'           = 'Full'
                            'Everyone'                         = 'Read'
                            'NT AUTHORITY\Authenticated Users' = 'Read'
                        }
                    }
                }

                (
                    $Result | Where-Object Name -EQ $expected.Name | ConvertTo-Json
                ) |
                Should -BeExactly ($expected | ConvertTo-Json)
            }
        }
    }
    Context 'correct' {
        BeforeAll {
            $testSmbShareAccess = Get-SmbShareAccess -Name $testSmbShare[1].Name
            $testSmbShareAccess.ForEach(
                {
                    $testRevokeParams = @{
                        Name        = $testSmbShare[1].Name
                        AccountName = $_.AccountName
                        Force       = $true
                    }
                    Revoke-SmbShareAccess @testRevokeParams
                }
            )

            @(
                @{
                    AccountName = 'Administrators'
                    AccessRight = 'Full'
                }
                @{
                    AccountName = 'Authenticated users'
                    AccessRight = 'Change'
                }
            ).ForEach(
                {
                    $testGrantParams = $_
                    Grant-SmbShareAccess -Name $testSmbShare[1].Name @testGrantParams -Force
                }
            )

            Mock Revoke-SmbShareAccess
            Mock Grant-SmbShareAccess

            $actual = .$testScript -Path $testSmbShare[1].Path -Flag $false
        }
        It 'do nothing' {
            Should -Not -Invoke Revoke-SmbShareAccess -Scope Context
            Should -Not -Invoke Grant-SmbShareAccess -Scope Context
        }
        It 'return no object' {
            $actual | Should -BeNullOrEmpty
        }
    }
    Context 'not set, because there is no smb share for the folder' {
        It 'do nothing' {
            $testFolder = New-Item -Path 'TestDrive:\s3' -ItemType Directory

            $actual = .$testScript -Path $testFolder.FullName -Flag $true

            $actual | Should -BeNullOrEmpty
        }
    }
}
Describe 'set Access Based Enumeration' {
    Context 'when Flag is TRUE' {
        BeforeAll {
            Set-SmbShare -Name $testSmbShare[1].Name -FolderEnumerationMode 'Unrestricted' -Force

            $actual = .$testScript -Path $testSmbShare[1].Path -Flag $true
        }
        It "to enabled" {
            (Get-SmbShare -Name $testSmbShare[1].Name).FolderEnumerationMode |
            Should -BeExactly 'AccessBased'
        }
        It 'and return a Warning object' {
            $expected = [PSCustomObject]@{
                Type        = 'Warning'
                Name        = 'Access Based Enumeration'
                Description = "Access Based Enumeration should be set to '$true'. This will hide files and folders where the users don't have access to. We fixed this now."
                Value       = @{
                    $testSmbShare[1].Name = $testSmbShare[1].Path
                }
            }

            (
                $actual | Where-Object Name -EQ $expected.Name | ConvertTo-Json
            ) |
            Should -BeExactly ($expected | ConvertTo-Json)
        }
    }
    Context 'when Flag is FALSE' {
        BeforeAll {
            Set-SmbShare -Name $testSmbShare[1].Name -FolderEnumerationMode 'AccessBased' -Force

            $actual = .$testScript -Path $testSmbShare[1].Path -Flag $false
        }
        It "to enabled" {
            (Get-SmbShare -Name $testSmbShare[1].Name).FolderEnumerationMode |
            Should -BeExactly 'Unrestricted'
        }
        It 'and return a Warning object' {
            $expected = [PSCustomObject]@{
                Type        = 'Warning'
                Name        = 'Access Based Enumeration'
                Description = "Access Based Enumeration should be set to '$false'. This will hide files and folders where the users don't have access to. We fixed this now."
                Value       = @{
                    $testSmbShare[1].Name = $testSmbShare[1].Path
                }
            }

            (
                $actual | Where-Object Name -EQ $expected.Name | ConvertTo-Json
            ) |
            Should -BeExactly ($expected | ConvertTo-Json)
        }
    }
}