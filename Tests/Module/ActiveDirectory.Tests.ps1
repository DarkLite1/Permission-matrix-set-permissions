#requires -Modules Pester

Describe 'ActiveDirectory.ps1 - AD Lookup Functions' {

    BeforeAll {
        $root = Split-Path -Parent $MyInvocation.MyCommand.Path
        $file = Join-Path $root '../Modules/Toolbox.PermissionMatrixHC/Private/ActiveDirectory.ps1'
        . $file
    }


    # -------------------------------------------------------------
    # 1. TEST: Get-ADObjectDetailHC
    # -------------------------------------------------------------

    Context 'Get-ADObjectDetailHC - Directory Search & Group Logic' {

        BeforeEach {
            # Mock DirectorySearcher
            Mock -CommandName New-Object -ParameterFilter {
                $TypeName -eq 'System.DirectoryServices.DirectorySearcher'
            } -MockWith {
                # We simulate a DirectorySearcher instance
                $fakeSearcher = [pscustomobject]@{
                    Filter           = $null
                    PropertiesToLoad = New-Object System.Collections.ArrayList
                    FindOne          = $null   # will be injected per test
                }
                return $fakeSearcher
            }

            # Default mock for FindOne
            Mock -CommandName Invoke-Expression {}
        }


        It 'Returns empty adObject when search result is null' {

            # Set DirectorySearcher.FindOne to return null
            Mock -CommandName New-Object -ParameterFilter {
                $TypeName -eq 'System.DirectoryServices.DirectorySearcher'
            } -MockWith {
                [pscustomobject]@{
                    Filter           = $null
                    PropertiesToLoad = New-Object System.Collections.ArrayList
                    FindOne          = { $null }
                }
            }

            $res = Get-ADObjectDetailHC -ADObjectName 'TestUser' -Type SamAccountName -MaxThreads 1
            $res[0].adObject | Should -BeNullOrEmpty
            $res[0].adGroupMember | Should -BeNullOrEmpty
            $res[0].SamAccountName | Should -Be 'TestUser'
        }


        It 'Returns adObject with user properties when user is found' {

            $fakeProps = @{
                'distinguishedname' = @('CN=TestUser,DC=lab,DC=local')
                'samaccountname'    = @('TestUser')
                'name'              = @('Test User')
                'objectclass'       = @('top', 'person', 'organizationalPerson', 'user')
            }

            # Simulate DirectorySearcher -> FindOne returning a "searchResult"
            Mock -CommandName New-Object -ParameterFilter {
                $TypeName -eq 'System.DirectoryServices.DirectorySearcher'
            } -MockWith {
                [pscustomobject]@{
                    Filter           = $null
                    PropertiesToLoad = New-Object System.Collections.ArrayList
                    FindOne          = { [pscustomobject]@{ Properties = $fakeProps } }
                }
            }

            $res = Get-ADObjectDetailHC -ADObjectName 'TestUser' -Type SamAccountName -MaxThreads 1
            $obj = $res[0].adObject

            $obj.SamAccountName | Should -Be 'TestUser'
            $obj.ObjectClass | Should -Be 'user'
            $obj.Name | Should -Be 'Test User'
            $obj.DistinguishedName | Should -Match 'CN=TestUser'
        }


        It 'Returns group object and expands Domain Users special-case' {

            $fakeProps = @{
                'distinguishedname' = @('CN=Domain Users,CN=Users,DC=lab,DC=local')
                'samaccountname'    = @('Domain Users')
                'name'              = @('Domain Users')
                'objectclass'       = @('top', 'group')
            }

            # Simulate DirectorySearcher -> FindOne returning a group
            Mock -CommandName New-Object -ParameterFilter {
                $TypeName -eq 'System.DirectoryServices.DirectorySearcher'
            } -MockWith {
                [pscustomobject]@{
                    Filter           = $null
                    PropertiesToLoad = New-Object System.Collections.ArrayList
                    FindOne          = { [pscustomobject]@{ Properties = $fakeProps } }
                }
            }

            # Mock PrincipalContext and GroupPrincipal
            Mock -CommandName New-Object -ParameterFilter {
                $TypeName -eq 'System.DirectoryServices.AccountManagement.PrincipalContext'
            } -MockWith {
                return [pscustomobject]@{}
            }

            Mock -CommandName [System.DirectoryServices.AccountManagement.GroupPrincipal]::FindByIdentity -MockWith {
                return $null  # Should not be used for Domain Users
            }

            $res = Get-ADObjectDetailHC -ADObjectName 'Domain Users' -Type SamAccountName -MaxThreads 1

            $members = $res[0].adGroupMember
            $members[0].Name | Should -Be 'All users'
        }


        It 'Expands normal group members via GroupPrincipal' {

            $fakeProps = @{
                'distinguishedname' = @('CN=MyGroup,CN=Users,DC=lab,DC=local')
                'samaccountname'    = @('MyGroup')
                'name'              = @('MyGroup')
                'objectclass'       = @('top', 'group')
            }

            # Mock DirectorySearcher
            Mock -CommandName New-Object -ParameterFilter {
                $TypeName -eq 'System.DirectoryServices.DirectorySearcher'
            } -MockWith {
                [pscustomobject]@{
                    Filter           = $null
                    PropertiesToLoad = New-Object System.Collections.ArrayList
                    FindOne          = { [pscustomobject]@{ Properties = $fakeProps } }
                }
            }

            # Mock PrincipalContext
            Mock -CommandName New-Object -ParameterFilter {
                $TypeName -eq 'System.DirectoryServices.AccountManagement.PrincipalContext'
            } -MockWith { [pscustomobject]@{} }

            # Mock GroupPrincipal
            Mock -CommandName [System.DirectoryServices.AccountManagement.GroupPrincipal]::FindByIdentity -MockWith {
                # Fake group principal
                return [pscustomobject]@{
                    GetMembers = {
                        @(
                            [pscustomobject]@{
                                StructuralObjectClass = 'user'
                                Name                  = 'UserA'
                                SamAccountName        = 'UserA'
                                DistinguishedName     = 'CN=UserA,DC=lab,DC=local'
                            }
                        )
                    }
                    Dispose    = {}
                }
            }

            $res = Get-ADObjectDetailHC -ADObjectName 'MyGroup' -Type SamAccountName -MaxThreads 1
            $res[0].adGroupMember[0].SamAccountName | Should -Be 'UserA'
        }
    }



    # -------------------------------------------------------------
    # 2. TEST: Get-AdUserPrincipalNameHC
    # -------------------------------------------------------------

    Context 'Get-AdUserPrincipalNameHC - Email → UPN conversion' {

        It 'Returns notFound list when no AD matches' {

            Mock Get-ADObject { $null }

            $res = Get-AdUserPrincipalNameHC -Name 'nobody@example.com'

            $res.notFound | Should -Contain 'nobody@example.com'
            $res.userPrincipalName.Count | Should -Be 0
        }


        It 'Returns UPN for user mail match' {

            Mock Get-ADObject {
                [pscustomobject]@{
                    ObjectClass    = 'user'
                    SamAccountName = 'u1'
                    Mail           = 'u1@example.com'
                }
            }

            Mock Get-ADUser {
                [pscustomobject]@{
                    Enabled           = $true
                    Mail              = 'u1@example.com'
                    SamAccountName    = 'u1'
                    UserPrincipalName = 'u1@domain.local'
                }
            }

            $res = Get-AdUserPrincipalNameHC -Name 'u1@example.com'

            $res.userPrincipalName | Should -Contain 'u1@domain.local'
        }


        It 'Expands group membership and filters disabled users' {

            # 1. Mock Get-ADObject for a group
            Mock Get-ADObject {
                return [pscustomobject]@{
                    ObjectClass = 'group'
                    Name        = 'MyGroup'
                }
            }

            # 2. Group members
            Mock Get-ADGroupMember {
                return @(
                    [pscustomobject]@{ SamAccountName = 'uA' },
                    [pscustomobject]@{ SamAccountName = 'uB' }
                )
            }

            # 3. Tell Get-ADUser which users are valid
            Mock Get-ADUser {
                param($Identity)

                switch ($Identity.SamAccountName) {
                    'uA' {
                        return [pscustomobject]@{
                            Enabled           = $true
                            Mail              = 'uA@example.com'
                            SamAccountName    = 'uA'
                            UserPrincipalName = 'uA@domain.local'
                        }
                    }
                    'uB' {
                        return [pscustomobject]@{
                            Enabled           = $false
                            Mail              = 'uB@example.com'
                            SamAccountName    = 'uB'
                            UserPrincipalName = 'uB@domain.local'
                        }
                    }
                }
            }

            $res = Get-AdUserPrincipalNameHC -Name 'MyGroup@example.com'

            $res.userPrincipalName | Should -Contain 'uA@domain.local'
            $res.userPrincipalName | Should -Not -Contain 'uB@domain.local'
        }


        It 'Excludes SamAccountNames passed via -ExcludeSamAccountName' {

            Mock Get-ADObject {
                return [pscustomobject]@{
                    ObjectClass    = 'user'
                    SamAccountName = 'skipME'
                    Mail           = 'skip@example.com'
                }
            }

            Mock Get-ADUser {
                return [pscustomobject]@{
                    Enabled           = $true
                    Mail              = 'skip@example.com'
                    SamAccountName    = 'skipME'
                    UserPrincipalName = 'skip@domain.local'
                }
            }

            $res = Get-AdUserPrincipalNameHC `
                -Name 'skip@example.com' `
                -ExcludeSamAccountName 'skipME'

            $res.userPrincipalName.Count | Should -Be 0
        }
    }
}