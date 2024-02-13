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
        {
            Remove-SmbShare -Name $_.Name -Force -EA Ignore
            New-SmbShare -Name $_.Name -Path $_.Path
            Grant-SmbShareAccess -Name $_.Name -AccountName 'Everyone' -AccessRight 'Full' -Force
        }
    )

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')

    Function Test-IsRequiredPowerShellVersionHC {}

    Mock Get-ItemPropertyValue -MockWith { 461808 }
    Mock Test-IsAdminHC { $true }
    Mock Test-IsRequiredPowerShellVersionHC { $true }
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
    It 'PowerShell 5.1 or later is not installed' {
        Mock Test-IsRequiredPowerShellVersionHC { $false }
        $expected = [PSCustomObject]@{
            Type        = 'FatalError'
            Name        = 'PowerShell version'
            Description = "PowerShell version 5.1 or higher is required to be able to use advanced methods."
            Value       = "PowerShell $($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor)"
        }

        $actual = .$testScript -Path 'NotExistingNotImportant' -Flag $true | Where-Object { $_.Name -eq $expected.Name }

        $actual | ConvertTo-Json |
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
Context 'set the Access Based Enumeration flag' {
    It "to enabled when the 'Flag' parameter is set to TRUE" {
        Set-SmbShare -Name $testSmbShare[0].Name -FolderEnumerationMode Unrestricted -Force

        .$testScript -Path $testSmbShare[0].Path -Flag $true

        (Get-SmbShare -Name $testSmbShare[0].Name).FolderEnumerationMode |
        Should -BeExactly 'AccessBased'
        Test-AccessBasedEnumerationHC -Name $testSmbShare[0].Name | Should -BeTrue
    } #-Tag test
    It "to disabled when the 'Flag' parameter is set to FALSE" {
        Set-SmbShare -Name $testSmbShare[0].Name -FolderEnumerationMode AccessBased -Force

        .$testScript -Path $testSmbShare[0].Path -Flag $false

        (Get-SmbShare -Name $testSmbShare[0].Name).FolderEnumerationMode |
        Should -BeExactly 'Unrestricted'
        Test-AccessBasedEnumerationHC -Name $testSmbShare[0].Name | Should -BeFalse
    }
    It 'only on the requested folder, not on other folders' {
        Set-SmbShare -Name $testSmbShare[0].Name -FolderEnumerationMode Unrestricted -Force
        Set-SmbShare -Name $testSmbShare[1].Name -FolderEnumerationMode Unrestricted -Force

        .$testScript -Path $testSmbShare[0].Path -Flag $true

        (Get-SmbShare -Name $testSmbShare[0].Name).FolderEnumerationMode |
        Should -BeExactly 'AccessBased' -Because 'we enabled ABE on this folder'
        (Get-SmbShare -Name $testSmbShare[1].Name).FolderEnumerationMode |
        Should -BeExactly 'Unrestricted' -Because "we didn't enable ABE on this folder"
    }
    It 'on multiple folders and ignore duplicates' {
        Set-SmbShare -Name $testSmbShare[0].Name -FolderEnumerationMode Unrestricted -Force
        Set-SmbShare -Name $testSmbShare[1].Name -FolderEnumerationMode Unrestricted -Force

        .$testScript -Path $testSmbShare[0].Path, $testSmbShare[1].Path, $testSmbShare[0].Path -Flag $true

        (Get-SmbShare -Name $testSmbShare[0].Name).FolderEnumerationMode |
        Should -BeExactly 'AccessBased'
        (Get-SmbShare -Name $testSmbShare[1].Name).FolderEnumerationMode |
        Should -BeExactly 'AccessBased'
    }
    It 'on multiple folders and return the results' {
        Set-SmbShare -Name $testSmbShare[0].Name -FolderEnumerationMode Unrestricted -Force
        Set-SmbShare -Name $testSmbShare[1].Name -FolderEnumerationMode Unrestricted -Force

        $expected = [PSCustomObject]@{
            Type        = 'Warning'
            Name        = 'Access Based Enumeration'
            Description = "Access Based Enumeration should be set to '$true'. This will hide files and folders where the users don't have access to. We fixed this now."
            Value       = @{
                $testSmbShare[0].Name = $testSmbShare[0].Path.FullName
                $testSmbShare[1].Name = $testSmbShare[1].Path.FullName
            }
        }

        $actual = .$testScript -Path $testSmbShare[0].Path, $testSmbShare[1].Path, $testSmbShare[0].Path -Flag $true

        (
            $actual | Where-Object Name -EQ 'Access Based Enumeration' |
            ConvertTo-Json
        ) |
        Should -BeExactly ($expected | ConvertTo-Json)
    }
}
Context "Set share permissions to 'FullControl for Administrators' and 'Read & Executed for Authenticated users'" {
    It 'when they are incorrect' {
        Grant-SmbShareAccess -Name $testSmbShare[0].Name -AccountName Administrators -AccessRight Full –Force
        Grant-SmbShareAccess -Name $testSmbShare[0].Name -AccountName Everyone -AccessRight Read –Force
        Grant-SmbShareAccess -Name $testSmbShare[0].Name -AccountName 'Authenticated users' -AccessRight Read –Force

        $Result = .$testScript -Path $testSmbShare[0].Path -Flag $true

        $actual = Get-SmbShareAccess -Name $testSmbShare[0].Name

        #region Verify share permissions
        $actual.Count | Should -BeExactly 2

        $actual.ForEach( {
                $_.Name | Should -Be $testSmbShare[0].Name
                $_.AccessControlType | Should -Be 'Allow'
            })
        ($actual | Where-Object AccountName -EQ 'NT AUTHORITY\Authenticated Users').AccessRight | Should -Be 'Change'
        ($actual | Where-Object AccountName -EQ 'BUILTIN\Administrators').AccessRight | Should -Be 'Full'
        #endregion

        #verify Script output
        $expected = [PSCustomObject]@{
            Type        = 'Warning'
            Name        = 'Share permissions'
            Description = "The share permissions are now set to 'Administrators: FullControl' and 'Authenticated users: Change'. The effective permissions are managed on NTFS level."
            Value       = @{$testSmbShare[0].Name = @{
                    'NT AUTHORITY\Authenticated Users' = 'Read'
                    'Everyone'                         = 'Read'
                    'BUILTIN\Administrators'           = 'FullControl'
                }
            }
        }

        (
            $Result | Where-Object Name -EQ $expected.Name | ConvertTo-Json
        ) |
        Should -BeExactly ($expected | ConvertTo-Json)
        #endregion
    }
    It "but don't change anything when they are already correct" {
        Remove-SmbShare -Name $testSmbShare[0].Name -Force -EA Ignore
        New-SmbShare -Name $testSmbShare[0].Name -Path $testSmbShare[0].Path
        Grant-SmbShareAccess -Name $testSmbShare[0].Name -AccountName Administrators -AccessRight Full -Force
        Grant-SmbShareAccess -Name $testSmbShare[0].Name -AccountName 'Authenticated Users' -AccessRight Change -Force
        Set-SmbShare -Name $testSmbShare[0].Name -FolderEnumerationMode AccessBased -Force

        $Result = .$testScript -Path $testSmbShare[0].Path -Flag $true

        $actual = Get-SmbShareAccess -Name $testSmbShare[0].Name | Where-Object Name -EQ 'Share permissions' | Should -BeNullOrEmpty
    }
    It "except when there's no shared folder, do nothing" {
        Remove-SmbShare -Name $testSmbShare[0].Name -Force -EA Ignore

        $Result = .$testScript -Path $testSmbShare[0].Path -Flag $true

        ($Result | Where-Object Action -EQ ACL) | Should -BeNullOrEmpty
        #endregion
    }
}


