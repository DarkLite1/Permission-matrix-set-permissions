#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }
#Requires -Module ImportExcel

Describe 'Validation.ps1 - Updated Validation Functions' {
    BeforeDiscovery {
        . "$PSScriptRoot/../../Helpers/Fixtures.Matrix.ps1"

        $script:PermissionFixtures = Get-MatrixPermissionsFixtures
    }

    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        Get-ChildItem "$moduleRoot\Private" -Filter '*.ps1' -File |
        ForEach-Object { . $_.FullName }

        . "$root/Tests/Helpers/Fixtures.Excel.ps1"
        . "$root/Tests/Helpers/Fixtures.Matrix.ps1"
        . "$root/Tests/Helpers/Fixtures.Json.ps1"

        function Get-RoundTripPermissions {
            param([Parameter(Mandatory)][string]$Scenario)

            $spec = New-MatrixPermissionsFixtureRows -Scenario $Scenario

            $dir = Join-Path 'TestDrive:' 'Matrix'
            if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
            $path = Join-Path $dir "Permissions_$Scenario.xlsx"

            New-MatrixPermissionsExcelFixture -Path $path -Spec $spec | Out-Null

            return @(Import-Excel -Path $path -WorksheetName 'Permissions' -NoHeader -DataOnly -ErrorAction Stop)
        }
    }

    Context 'Test-AdObjectInMatrixHC' {
        BeforeAll {
            # A matrix is an array of folder objects; only .ACL.Keys is read by
            # the function under test, so the rest is left minimal on purpose.
            function New-FakeMatrixHC {
                param([Parameter(Mandatory)][string[]]$AdObjectName)

                $acl = @{}
                foreach ($name in $AdObjectName) { $acl[$name] = 'W' }

                return @(
                    [pscustomobject]@{
                        Path = '\\srv\share'
                        ACL  = $acl
                    }
                )
            }

            # Mirrors one element of the Get-ADObjectDetailHC output shape.
            function New-FakeAdDetailHC {
                param(
                    [Parameter(Mandatory)][string]$SamAccountName,
                    [ValidateSet('user', 'group')][string]$ObjectClass = 'group',
                    [object[]]$Member = @(),
                    [switch]$NotFound
                )

                $adObject = if ($NotFound) { $null }
                else {
                    [pscustomobject]@{
                        DistinguishedName = "CN=$SamAccountName,DC=contoso,DC=com"
                        SamAccountName    = $SamAccountName
                        ObjectSid         = "S-1-5-21-1-1-1-$($SamAccountName.Length)"
                        ObjectClass       = $ObjectClass
                        Name              = $SamAccountName
                    }
                }

                [pscustomobject]@{
                    SamAccountName = $SamAccountName
                    adObject       = $adObject
                    adGroupMember  = $Member
                }
            }

            # Enabled is deliberately [object] so $null (nested groups, and the
            # synthetic 'All users' member of 'Domain Users') can be expressed.
            function New-FakeAdMemberHC {
                param(
                    [Parameter(Mandatory)][string]$SamAccountName,
                    [object]$Enabled = $true,
                    [string]$ObjectClass = 'user'
                )

                [pscustomobject]@{
                    objectClass       = $ObjectClass
                    Name              = $SamAccountName
                    SamAccountName    = $SamAccountName
                    DistinguishedName = "CN=$SamAccountName,DC=contoso,DC=com"
                    Enabled           = $Enabled
                }
            }

            function Get-EmptyGroupCheckHC {
                param([object[]]$Checks)

                @($Checks | Where-Object { $_.Name -eq 'AD groups without members' })
            }

            <#
             'Value' is an array of objects with a Name and a Reason, one per
             group. Flatten it to text for the assertions that only care
             whether a group is mentioned.
            #>
            function Get-EmptyGroupTextHC {
                param([object]$Check)

                (@($Check.Value) | ForEach-Object { "$($_.Name) $($_.Reason)" }) -join "`n"
            }

            # Return the reason reported for a single group
            function Get-EmptyGroupReasonHC {
                param([object]$Check, [string]$GroupName)

                @($Check.Value) |
                Where-Object { $_.Name -eq $GroupName } |
                Select-Object -First 1 -ExpandProperty 'Reason'
            }
        }

        Context 'no notice is raised' {
            It 'when the group holds an enabled, non-placeholder member' {
                $matrix = New-FakeMatrixHC -AdObjectName 'GRP-HR'
                $ad = @(
                    New-FakeAdDetailHC -SamAccountName 'GRP-HR' -Member @(
                        New-FakeAdMemberHC -SamAccountName 'bmarley'
                    )
                )

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad `
                    -ExcludedSamAccountName @('cnorris')

                Get-EmptyGroupCheckHC $res | Should-BeCollection @()
            }

            It 'when the group holds a placeholder AND a real member' {
                $matrix = New-FakeMatrixHC -AdObjectName 'GRP-HR'
                $ad = @(
                    New-FakeAdDetailHC -SamAccountName 'GRP-HR' -Member @(
                        New-FakeAdMemberHC -SamAccountName 'cnorris'
                        New-FakeAdMemberHC -SamAccountName 'bmarley'
                    )
                )

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad `
                    -ExcludedSamAccountName @('cnorris')

                Get-EmptyGroupCheckHC $res | Should-BeCollection @()
            }

            It 'when the matrix references a user account instead of a group' {
                # A user object has no adGroupMember; it must never be flagged.
                $matrix = New-FakeMatrixHC -AdObjectName 'bmarley'
                $ad = @(
                    New-FakeAdDetailHC -SamAccountName 'bmarley' -ObjectClass 'user'
                )

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad `
                    -ExcludedSamAccountName @('cnorris')

                Get-EmptyGroupCheckHC $res | Should-BeCollection @()
            }

            It 'when a member has a null Enabled value (nested group)' {
                # GetMembers($true) returns $null Enabled for non-authenticable
                # principals. These count as real members.
                $matrix = New-FakeMatrixHC -AdObjectName 'GRP-HR'
                $ad = @(
                    New-FakeAdDetailHC -SamAccountName 'GRP-HR' -Member @(
                        New-FakeAdMemberHC -SamAccountName 'GRP-NESTED' `
                            -Enabled $null -ObjectClass 'group'
                    )
                )

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad

                Get-EmptyGroupCheckHC $res | Should-BeCollection @()
            }

            It "when the group is 'Domain Users' with its synthetic member" {
                $matrix = New-FakeMatrixHC -AdObjectName 'Domain Users'
                $ad = @(
                    New-FakeAdDetailHC -SamAccountName 'Domain Users' -Member @(
                        New-FakeAdMemberHC -SamAccountName 'All users' -Enabled $null
                    )
                )

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad

                Get-EmptyGroupCheckHC $res | Should-BeCollection @()
            }

            It 'when the matrix has no AD objects at all' {
                $matrix = @([pscustomobject]@{ Path = '\\srv\share'; ACL = @{} })

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject @()

                $res | Should-BeCollection @()
            }
        }

        Context 'a notice is raised' {
            It 'when the placeholder account is the only member' {
                $matrix = New-FakeMatrixHC -AdObjectName 'GRP-HR'
                $ad = @(
                    New-FakeAdDetailHC -SamAccountName 'GRP-HR' -Member @(
                        New-FakeAdMemberHC -SamAccountName 'cnorris'
                    )
                )

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad `
                    -ExcludedSamAccountName @('cnorris')

                $check = @(Get-EmptyGroupCheckHC $res)
                $check.Count | Should-Be 1
                $check[0].Type | Should-Be 'Information'
                Get-EmptyGroupTextHC $check[0] | Should-MatchString 'GRP-HR'
            }

            It 'when the group has zero members' {
                $matrix = New-FakeMatrixHC -AdObjectName 'GRP-HR'
                $ad = @(New-FakeAdDetailHC -SamAccountName 'GRP-HR' -Member @())

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad `
                    -ExcludedSamAccountName @('cnorris')

                @(Get-EmptyGroupCheckHC $res).Count | Should-Be 1
            }

            It 'when adGroupMember is null rather than an empty array' {
                # Get-ADObjectDetailHC leaves adGroupMember at $null when the
                # group could not be expanded.
                $matrix = New-FakeMatrixHC -AdObjectName 'GRP-HR'
                $ad = @(
                    [pscustomobject]@{
                        SamAccountName = 'GRP-HR'
                        adObject       = [pscustomobject]@{
                            SamAccountName = 'GRP-HR'
                            ObjectClass    = 'group'
                            ObjectSid      = 'S-1-5-21-1-1-1-1'
                        }
                        adGroupMember  = $null
                    }
                )

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad

                @(Get-EmptyGroupCheckHC $res).Count | Should-Be 1
            }

            It 'when every member is a disabled account' {
                $matrix = New-FakeMatrixHC -AdObjectName 'GRP-HR'
                $ad = @(
                    New-FakeAdDetailHC -SamAccountName 'GRP-HR' -Member @(
                        New-FakeAdMemberHC -SamAccountName 'bmarley' -Enabled $false
                        New-FakeAdMemberHC -SamAccountName 'jsmith' -Enabled $false
                    )
                )

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad

                @(Get-EmptyGroupCheckHC $res).Count | Should-Be 1
            }

            It 'when members are only placeholders and disabled accounts' {
                $matrix = New-FakeMatrixHC -AdObjectName 'GRP-HR'
                $ad = @(
                    New-FakeAdDetailHC -SamAccountName 'GRP-HR' -Member @(
                        New-FakeAdMemberHC -SamAccountName 'cnorris'
                        New-FakeAdMemberHC -SamAccountName 'bmarley' -Enabled $false
                    )
                )

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad `
                    -ExcludedSamAccountName @('cnorris')

                @(Get-EmptyGroupCheckHC $res).Count | Should-Be 1
            }

            It 'when no ExcludedSamAccountName is supplied at all' {
                $matrix = New-FakeMatrixHC -AdObjectName 'GRP-HR'
                $ad = @(New-FakeAdDetailHC -SamAccountName 'GRP-HR' -Member @())

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad

                @(Get-EmptyGroupCheckHC $res).Count | Should-Be 1
            }

            It 'when ExcludedSamAccountName is passed as $null' {
                # Matrix.ExcludedSamAccountName is optional in the JSON.
                $matrix = New-FakeMatrixHC -AdObjectName 'GRP-HR'
                $ad = @(
                    New-FakeAdDetailHC -SamAccountName 'GRP-HR' -Member @(
                        New-FakeAdMemberHC -SamAccountName 'bmarley'
                    )
                )

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad `
                    -ExcludedSamAccountName $null

                Get-EmptyGroupCheckHC $res | Should-BeCollection @()
            }
        }

        Context 'placeholder matching' {
            It 'is case insensitive' {
                $matrix = New-FakeMatrixHC -AdObjectName 'GRP-HR'
                $ad = @(
                    New-FakeAdDetailHC -SamAccountName 'GRP-HR' -Member @(
                        New-FakeAdMemberHC -SamAccountName 'CNorris'
                    )
                )

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad `
                    -ExcludedSamAccountName @('cnorris')

                @(Get-EmptyGroupCheckHC $res).Count | Should-Be 1
            }

            It 'honours several placeholder accounts' {
                $matrix = New-FakeMatrixHC -AdObjectName 'GRP-HR'
                $ad = @(
                    New-FakeAdDetailHC -SamAccountName 'GRP-HR' -Member @(
                        New-FakeAdMemberHC -SamAccountName 'cnorris'
                        New-FakeAdMemberHC -SamAccountName 'svc-placeholder'
                    )
                )

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad `
                    -ExcludedSamAccountName @('cnorris', 'svc-placeholder')

                @(Get-EmptyGroupCheckHC $res).Count | Should-Be 1
            }
        }

        Context 'reporting shape' {
            It 'collects every empty group into a single check' {
                $matrix = New-FakeMatrixHC -AdObjectName @(
                    'GRP-HR', 'GRP-FIN', 'GRP-IT'
                )
                $ad = @(
                    New-FakeAdDetailHC -SamAccountName 'GRP-HR' -Member @(
                        New-FakeAdMemberHC -SamAccountName 'cnorris'
                    )
                    New-FakeAdDetailHC -SamAccountName 'GRP-FIN' -Member @()
                    New-FakeAdDetailHC -SamAccountName 'GRP-IT' -Member @(
                        New-FakeAdMemberHC -SamAccountName 'bmarley'
                    )
                )

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad `
                    -ExcludedSamAccountName @('cnorris')

                $check = @(Get-EmptyGroupCheckHC $res)
                $check.Count | Should-Be 1

                $text = Get-EmptyGroupTextHC $check[0]
                $text | Should-MatchString 'GRP-HR'
                $text | Should-MatchString 'GRP-FIN'
                $text | Should-NotMatchString 'GRP-IT'
            }

            It 'reports the same group only once when granted on several folders' {
                $matrix = @(
                    [pscustomobject]@{ Path = '\\srv\a'; ACL = @{ 'GRP-HR' = 'W' } }
                    [pscustomobject]@{ Path = '\\srv\b'; ACL = @{ 'GRP-HR' = 'R' } }
                )
                $ad = @(New-FakeAdDetailHC -SamAccountName 'GRP-HR' -Member @())

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad

                $check = @(Get-EmptyGroupCheckHC $res)
                $check.Count | Should-Be 1

                $text = Get-EmptyGroupTextHC $check[0]
                ([regex]::Matches($text, 'GRP-HR')).Count | Should-Be 1
            }

            It 'emits one array entry per group instead of a single joined string' {
                # Mirrors 'Inherited permissions incorrect', whose OldAcl /
                # NewAcl / MatrixFileAcl arrays are one entry per line.
                $matrix = New-FakeMatrixHC -AdObjectName @('GRP-HR', 'GRP-FIN')
                $ad = @(
                    New-FakeAdDetailHC -SamAccountName 'GRP-HR' -Member @()
                    New-FakeAdDetailHC -SamAccountName 'GRP-FIN' -Member @()
                )

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad

                $check = @(Get-EmptyGroupCheckHC $res)
                @($check[0].Value).Count | Should-Be 2
                Get-EmptyGroupTextHC $check[0] | Should-NotMatchString ','
            }

            It 'sorts the entries by group name' {
                $matrix = New-FakeMatrixHC -AdObjectName @(
                    'GRP-IT', 'GRP-FIN', 'GRP-HR'
                )
                $ad = @(
                    New-FakeAdDetailHC -SamAccountName 'GRP-IT' -Member @()
                    New-FakeAdDetailHC -SamAccountName 'GRP-FIN' -Member @()
                    New-FakeAdDetailHC -SamAccountName 'GRP-HR' -Member @()
                )

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad

                $check = @(Get-EmptyGroupCheckHC $res)

                $names = @($check[0].Value) | ForEach-Object { $_.Name }

                <#
                 Compared as one joined string rather than with
                 Should-BeCollection: that assertion checks membership in
                 any order, so it would pass however the entries were
                 sorted - the exact thing this test exists to catch.
                #>
                $names -join ', ' | Should-Be 'GRP-FIN, GRP-HR, GRP-IT'
            }

            It 'emits a Name and a Reason per group, free of formatting' {
                # Replaces the old 'pads the group names so the reasons line
                # up' test: alignment is a rendering concern now, and the
                # value must carry the raw name with no padding or quotes.
                $matrix = New-FakeMatrixHC -AdObjectName @(
                    'GRP-HR', 'GRP-A-VERY-LONG-GROUP-NAME'
                )
                $ad = @(
                    New-FakeAdDetailHC -SamAccountName 'GRP-HR' -Member @()
                    New-FakeAdDetailHC -SamAccountName 'GRP-A-VERY-LONG-GROUP-NAME' -Member @()
                )

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad

                $check = @(Get-EmptyGroupCheckHC $res)
                $entries = @($check[0].Value)

                $entries.Count | Should-Be 2

                foreach ($entry in $entries) {
                    $entry.PSObject.Properties.Name |
                    Should-ContainCollection 'Name'
                    $entry.PSObject.Properties.Name |
                    Should-ContainCollection 'Reason'
                }

                # The short name is not padded out to the long one
                $short = $entries | Where-Object { $_.Name -eq 'GRP-HR' }
                $short.Name | Should-Be 'GRP-HR'
                $short.Reason | Should-Be 'no members'
            }

            It 'mentions the placeholder configuration in the description' {
                $matrix = New-FakeMatrixHC -AdObjectName 'GRP-HR'
                $ad = @(New-FakeAdDetailHC -SamAccountName 'GRP-HR' -Member @())

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad

                $check = @(Get-EmptyGroupCheckHC $res)
                $check[0].Description | Should-MatchString 'ExcludedSamAccountName'
                $check[0].Description | Should-MatchString 'no effective members'
            }

            It 'resolves the configured placeholder accounts in the description' {
                $matrix = New-FakeMatrixHC -AdObjectName 'GRP-HR'
                $ad = @(
                    New-FakeAdDetailHC -SamAccountName 'GRP-HR' -Member @(
                        New-FakeAdMemberHC -SamAccountName 'cnorris'
                    )
                )

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad `
                    -ExcludedSamAccountName @('svc-placeholder', 'cnorris')

                $check = @(Get-EmptyGroupCheckHC $res)
                $check[0].Description | Should-MatchString 'cnorris, svc-placeholder'
            }

            It 'states in the description when no placeholder is configured' {
                $matrix = New-FakeMatrixHC -AdObjectName 'GRP-HR'
                $ad = @(New-FakeAdDetailHC -SamAccountName 'GRP-HR' -Member @())

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad

                $check = @(Get-EmptyGroupCheckHC $res)
                $check[0].Description |
                Should-MatchString 'No placeholder accounts are configured'
            }
        }

        Context 'the reported reason' {
            It "is 'no members' when the group is empty" {
                $matrix = New-FakeMatrixHC -AdObjectName 'GRP-HR'
                $ad = @(New-FakeAdDetailHC -SamAccountName 'GRP-HR' -Member @())

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad `
                    -ExcludedSamAccountName @('cnorris')

                $check = @(Get-EmptyGroupCheckHC $res)
                Get-EmptyGroupReasonHC $check[0] 'GRP-HR' |
                Should-Be 'no members'
            }

            It "is 'no members' when adGroupMember is null" {
                $matrix = New-FakeMatrixHC -AdObjectName 'GRP-HR'
                $ad = @(
                    [pscustomobject]@{
                        SamAccountName = 'GRP-HR'
                        adObject       = [pscustomobject]@{
                            SamAccountName = 'GRP-HR'
                            ObjectClass    = 'group'
                            ObjectSid      = 'S-1-5-21-1-1-1-1'
                        }
                        adGroupMember  = $null
                    }
                )

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad

                $check = @(Get-EmptyGroupCheckHC $res)
                Get-EmptyGroupReasonHC $check[0] 'GRP-HR' |
                Should-Be 'no members'
            }

            It "is 'only placeholder accounts' when only placeholders are left" {
                $matrix = New-FakeMatrixHC -AdObjectName 'GRP-HR'
                $ad = @(
                    New-FakeAdDetailHC -SamAccountName 'GRP-HR' -Member @(
                        New-FakeAdMemberHC -SamAccountName 'cnorris'
                    )
                )

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad `
                    -ExcludedSamAccountName @('cnorris')

                $check = @(Get-EmptyGroupCheckHC $res)
                Get-EmptyGroupReasonHC $check[0] 'GRP-HR' |
                Should-Be 'only placeholder accounts'
            }

            It "is 'only disabled accounts' when every member is disabled" {
                $matrix = New-FakeMatrixHC -AdObjectName 'GRP-HR'
                $ad = @(
                    New-FakeAdDetailHC -SamAccountName 'GRP-HR' -Member @(
                        New-FakeAdMemberHC -SamAccountName 'bmarley' -Enabled $false
                        New-FakeAdMemberHC -SamAccountName 'jsmith' -Enabled $false
                    )
                )

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad `
                    -ExcludedSamAccountName @('cnorris')

                $check = @(Get-EmptyGroupCheckHC $res)
                Get-EmptyGroupReasonHC $check[0] 'GRP-HR' |
                Should-Be 'only disabled accounts'
            }

            It "is 'only placeholder and disabled accounts' for the mix" {
                $matrix = New-FakeMatrixHC -AdObjectName 'GRP-HR'
                $ad = @(
                    New-FakeAdDetailHC -SamAccountName 'GRP-HR' -Member @(
                        New-FakeAdMemberHC -SamAccountName 'cnorris'
                        New-FakeAdMemberHC -SamAccountName 'bmarley' -Enabled $false
                    )
                )

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad `
                    -ExcludedSamAccountName @('cnorris')

                $check = @(Get-EmptyGroupCheckHC $res)
                Get-EmptyGroupReasonHC $check[0] 'GRP-HR' |
                Should-Be 'only placeholder and disabled accounts'
            }

            It 'counts a disabled placeholder as a placeholder only' {
                # An account that is both the placeholder AND disabled must not
                # inflate the disabled count, or the reason would read as a mix.
                $matrix = New-FakeMatrixHC -AdObjectName 'GRP-HR'
                $ad = @(
                    New-FakeAdDetailHC -SamAccountName 'GRP-HR' -Member @(
                        New-FakeAdMemberHC -SamAccountName 'cnorris' -Enabled $false
                    )
                )

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad `
                    -ExcludedSamAccountName @('cnorris')

                $check = @(Get-EmptyGroupCheckHC $res)
                Get-EmptyGroupReasonHC $check[0] 'GRP-HR' |
                Should-Be 'only placeholder accounts'
            }

            It 'reports a different reason per group in the same check' {
                $matrix = New-FakeMatrixHC -AdObjectName @('GRP-HR', 'GRP-FIN')
                $ad = @(
                    New-FakeAdDetailHC -SamAccountName 'GRP-HR' -Member @()
                    New-FakeAdDetailHC -SamAccountName 'GRP-FIN' -Member @(
                        New-FakeAdMemberHC -SamAccountName 'bmarley' -Enabled $false
                    )
                )

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad

                $check = @(Get-EmptyGroupCheckHC $res)
                Get-EmptyGroupReasonHC $check[0] 'GRP-HR' |
                Should-Be 'no members'
                Get-EmptyGroupReasonHC $check[0] 'GRP-FIN' |
                Should-Be 'only disabled accounts'
            }
        }

        Context 'interaction with the unknown AD object check' {
            It 'still reports unknown AD objects as a FatalError' {
                $matrix = New-FakeMatrixHC -AdObjectName 'GRP-GHOST'
                $ad = @(New-FakeAdDetailHC -SamAccountName 'GRP-GHOST' -NotFound)

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad

                $fatal = @(
                    $res | Where-Object { $_.Name -eq 'Unknown AD Objects in Matrix' }
                )
                $fatal.Count | Should-Be 1
                $fatal[0].Type | Should-Be 'FatalError'
            }

            It 'does not also flag an unresolvable object as an empty group' {
                $matrix = New-FakeMatrixHC -AdObjectName 'GRP-GHOST'
                $ad = @(New-FakeAdDetailHC -SamAccountName 'GRP-GHOST' -NotFound)

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad

                Get-EmptyGroupCheckHC $res | Should-BeCollection @()
            }

            It 'reports both findings when the matrix mixes the two cases' {
                $matrix = New-FakeMatrixHC -AdObjectName @('GRP-GHOST', 'GRP-HR')
                $ad = @(
                    New-FakeAdDetailHC -SamAccountName 'GRP-GHOST' -NotFound
                    New-FakeAdDetailHC -SamAccountName 'GRP-HR' -Member @(
                        New-FakeAdMemberHC -SamAccountName 'cnorris'
                    )
                )

                $res = Test-AdObjectInMatrixHC -Matrix $matrix -ADObject $ad `
                    -ExcludedSamAccountName @('cnorris')

                @($res).Count | Should-Be 2
                $res.Type | Should-ContainCollection 'FatalError'
                $res.Type | Should-ContainCollection 'Information'
            }
        }
    }

    Context 'Test-MatrixFileHC' {
        It 'Warns for missing settings' {
            $M = @{ Settings = @(); Permissions = @('x') }
            $res = Test-MatrixFileHC -MatrixObject $M
            $res.Type | Should-ContainCollection 'Warning'
        }

        It 'Errors for missing permissions' {
            $M = @{ Settings = @('x'); Permissions = @() }
            $res = Test-MatrixFileHC -MatrixObject $M
            $res.Type | Should-ContainCollection 'FatalError'
        }
    }

    Context 'Test-MatrixPermissionsHC' {

        Context 'Happy path' {
            It 'returns nothing when the Valid fixture is supplied' {
                $perms = Get-RoundTripPermissions -Scenario 'Valid'

                $result = Test-MatrixPermissionsHC -Permissions $perms

                # Function only returns $checks when Count -gt 0, so success => $null.
                $result | Should-BeFalsy
            }
        }

        Context 'Data-driven checks from Get-MatrixPermissionsFixtures' {
            It 'flags <Issue> with check name <Expected>' -ForEach $PermissionFixtures {

                # The fixture 'Mutation' strings map 1:1 to a scenario name in
                # New-MatrixPermissionsFixtureRows; derive it from the Issue so we
                # can round-trip in-process rather than Invoke-Expression a string
                # that writes its own file.
                $scenario = switch ($Issue) {
                    'MissingADObjectName' { 'MissingADObjectName' }
                    'InvalidPermissionChar' { 'InvalidPermissionChar' }
                    'MissingRows' { 'MissingRows' }
                    'MissingColumns' { 'MissingColumns' }
                    'MissingFolderName' { 'MissingFolderName' }
                    'DuplicateFolderName' { 'DuplicateFolderName' }
                    'InaccessibleFolders' { 'InaccessibleFolders' }
                    default { throw "No scenario mapping for Issue '$Issue'" }
                }

                $perms = Get-RoundTripPermissions -Scenario $scenario

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $result | Should-BeTruthy
                ($result.Name) | Should-ContainCollection $Expected
            }
        }

        Context 'fatal errors exit immediately' {
            It 'returns ONLY "Missing rows" for the MissingRows fixture' {
                $perms = Get-RoundTripPermissions -Scenario 'MissingRows'

                $result = Test-MatrixPermissionsHC -Permissions $perms

                @($result).Count | Should-Be 1
                $result[0].Type | Should-Be 'FatalError'
                $result[0].Name | Should-Be 'Missing rows'
            }

            It 'returns ONLY "Missing columns" for the MissingColumns fixture' {
                $perms = Get-RoundTripPermissions -Scenario 'MissingColumns'

                $result = Test-MatrixPermissionsHC -Permissions $perms

                @($result).Count | Should-Be 1
                $result[0].Type | Should-Be 'FatalError'
                $result[0].Name | Should-Be 'Missing columns'
            }
        }

        Context 'Check types are correct' {
            It 'classifies InaccessibleFolders as a Warning, not a FatalError' {
                $perms = Get-RoundTripPermissions -Scenario 'InaccessibleFolders'
                $result = Test-MatrixPermissionsHC -Permissions $perms

                $warn = $result | Where-Object Name -EQ 'Inaccessible folders'
                $warn | Should-BeTruthy
                $warn.Type | Should-Be 'Warning'
            }

            It 'classifies InvalidPermissionChar as a FatalError' {
                $perms = Get-RoundTripPermissions -Scenario 'InvalidPermissionChar'
                $result = Test-MatrixPermissionsHC -Permissions $perms

                $err = $result | Where-Object Name -EQ 'Invalid permission character'
                $err | Should-BeTruthy
                $err.Type | Should-Be 'FatalError'
            }
        }

        Context 'Inaccessible folders: deepest folder detection' {
            BeforeAll {
                <#
                 Build an in-memory Permissions sheet shaped exactly like the
                 output of 'Import-Excel -NoHeader' (properties P1..Pn):
                 - rows 0-2 : header rows (SamAccountName per column)
                 - row 3    : parent folder permissions
                 - rows 4+  : folder rows
                 Building rows in-memory keeps each test focused on the
                 parent/child path logic without an Excel round trip.
                #>
                function New-PermissionsSheet {
                    param(
                        # One permission char (or $null) per AD object column
                        [array]$ParentPermissions = @('L', 'L'),
                        # Folder rows: @{ Path = '...'; Permissions = @(...) }
                        [array]$FolderRows
                    )

                    $permColCount = $ParentPermissions.Count
                    $colNames = @(1..($permColCount + 1)).ForEach({ "P$_" })

                    $newRow = {
                        param($firstColumn, $permissions)

                        $props = [ordered]@{ $colNames[0] = $firstColumn }
                        for ($i = 0; $i -lt $permColCount; $i++) {
                            $props[$colNames[$i + 1]] = if (
                                $permissions -and ($i -lt $permissions.Count)
                            ) {
                                $permissions[$i]
                            }
                            else { $null }
                        }
                        [pscustomobject]$props
                    }

                    $rows = [System.Collections.Generic.List[object]]::new()

                    # Header rows: SamAccountName on the first header row
                    $rows.Add(
                        (& $newRow $null @(1..$permColCount).ForEach({ "group$_" }))
                    )
                    $rows.Add((& $newRow $null $null))
                    $rows.Add((& $newRow $null $null))

                    # Parent folder permissions (row index 3)
                    $rows.Add((& $newRow $null $ParentPermissions))

                    foreach ($f in $FolderRows) {
                        $rows.Add((& $newRow $f.Path $f.Permissions))
                    }

                    return , $rows.ToArray()
                }
            }

            It 'does not flag a parent typed with a trailing backslash when its children grant access' {
                # Regression: 'BEL\L&D\Certificates\' (trailing backslash, only
                # L) was wrongly reported as a deepest folder even though
                # 'BEL\L&D\Certificates\AGG' below it grants W.
                $perms = New-PermissionsSheet -ParentPermissions @('L', 'L') -FolderRows @(
                    @{ Path = 'BEL\L&D\Certificates\'; Permissions = @('L', 'L') }
                    @{ Path = 'BEL\L&D\Certificates\AGG'; Permissions = @('W', 'W') }
                )

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $result | Should-BeFalsy
            }

            It 'flags a genuinely deepest folder with only List permissions' {
                $perms = New-PermissionsSheet -ParentPermissions @('L', 'L') -FolderRows @(
                    @{ Path = 'BEL\Marketing'; Permissions = @('L', $null) }
                )

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $warn = $result | Where-Object Name -EQ 'Inaccessible folders'
                $warn | Should-BeTruthy
                $warn.Value | Should-MatchString ([regex]::Escape('BEL\Marketing'))
            }

            It 'reports only the truly inaccessible folder, not the trailing-backslash parent' {
                $perms = New-PermissionsSheet -ParentPermissions @('L', 'L') -FolderRows @(
                    @{ Path = 'BEL\L&D\Certificates\'; Permissions = @('L', 'L') }
                    @{ Path = 'BEL\L&D\Certificates\AGG'; Permissions = @('W', 'W') }
                    @{ Path = 'BEL\Dead'; Permissions = @('L', $null) }
                )

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $warn = $result | Where-Object Name -EQ 'Inaccessible folders'
                $warn | Should-BeTruthy
                $warn.Value | Should-MatchString ([regex]::Escape('BEL\Dead'))
                $warn.Value | Should-NotMatchString 'Certificates'
            }

            It 'matches parent and child paths case-insensitively' {
                $perms = New-PermissionsSheet -ParentPermissions @('L', 'L') -FolderRows @(
                    @{ Path = 'BEL\ARCHIVE'; Permissions = @('L', 'L') }
                    @{ Path = 'bel\archive\2020'; Permissions = @('W', $null) }
                )

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $result | Should-BeFalsy
            }

            It 'handles wildcard characters in folder names' {
                # '-like' would treat '[2026]' as a wildcard set and break
                # parent/child matching; String.StartsWith must not.
                $perms = New-PermissionsSheet -ParentPermissions @('L', 'L') -FolderRows @(
                    @{ Path = 'BEL\Reports[2026]'; Permissions = @('L', 'L') }
                    @{ Path = 'BEL\Reports[2026]\Q1'; Permissions = @('W', $null) }
                )

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $result | Should-BeFalsy
            }

            It 'flags a List-only deepest folder even when the parent row grants access' {
                # An explicit 'L' sets a List-only ACL on the folder itself,
                # so the parent's W cannot make it accessible.
                $perms = New-PermissionsSheet -ParentPermissions @('W', 'L') -FolderRows @(
                    @{ Path = 'BEL\Marketing'; Permissions = @('L', $null) }
                )

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $warn = $result | Where-Object Name -EQ 'Inaccessible folders'
                $warn | Should-BeTruthy
                $warn.Value | Should-MatchString ([regex]::Escape('BEL\Marketing'))
            }

            It 'does not flag a blank deepest folder when the parent row grants access' {
                # A row without any permission inherits the parent ACL, so
                # the parent's W makes it accessible.
                $perms = New-PermissionsSheet -ParentPermissions @('W', 'L') -FolderRows @(
                    @{ Path = 'BEL\Marketing'; Permissions = @($null, $null) }
                )

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $result | Should-BeFalsy
            }

            It 'flags a blank deepest folder when the parent row grants no access' {
                # Inheriting from a parent that only grants L leaves the
                # folder without read or write access.
                $perms = New-PermissionsSheet -ParentPermissions @('L', 'L') -FolderRows @(
                    @{ Path = 'BEL\Marketing'; Permissions = @($null, $null) }
                )

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $warn = $result | Where-Object Name -EQ 'Inaccessible folders'
                $warn | Should-BeTruthy
                $warn.Value | Should-MatchString ([regex]::Escape('BEL\Marketing'))
            }

            It 'does not treat a folder with a similar name prefix as a child' {
                # 'BEL\App2' must not count as a child of 'BEL\App' — only
                # 'BEL\App\...' qualifies. 'BEL\App' has only L and no real
                # children, so it must be flagged.
                $perms = New-PermissionsSheet -ParentPermissions @('L', 'L') -FolderRows @(
                    @{ Path = 'BEL\App'; Permissions = @('L', $null) }
                    @{ Path = 'BEL\App2'; Permissions = @('W', $null) }
                )

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $warn = $result | Where-Object Name -EQ 'Inaccessible folders'
                $warn | Should-BeTruthy
                $warn.Value | Should-MatchString ([regex]::Escape('BEL\App'))
            }

            It 'reports the folder exactly as typed in Excel, trailing backslash included' {
                $perms = New-PermissionsSheet -ParentPermissions @('L', 'L') -FolderRows @(
                    @{ Path = 'BEL\Lonely\'; Permissions = @('L', $null) }
                )

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $warn = $result | Where-Object Name -EQ 'Inaccessible folders'
                $warn | Should-BeTruthy
                $warn.Value | Should-MatchString ([regex]::Escape('BEL\Lonely\'))
            }

            It 'does not flag a deepest folder marked with I (Ignore)' {
                $perms = New-PermissionsSheet -ParentPermissions @('L', 'L') -FolderRows @(
                    @{ Path = 'BEL\Temp'; Permissions = @('I', $null) }
                )

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $result | Should-BeFalsy
            }

            It 'does not flag subfolders of a folder marked with I (Ignore)' {
                # 'BEL\Temp\Cache' is List-only and deepest, but it sits
                # under an ignored folder, so the matrix does not manage it.
                $perms = New-PermissionsSheet -ParentPermissions @('L', 'L') -FolderRows @(
                    @{ Path = 'BEL\Temp'; Permissions = @('I', $null) }
                    @{ Path = 'BEL\Temp\Cache'; Permissions = @('L', $null) }
                )

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $result | Should-BeFalsy
            }

            It 'still flags inaccessible folders outside an ignored subtree' {
                $perms = New-PermissionsSheet -ParentPermissions @('L', 'L') -FolderRows @(
                    @{ Path = 'BEL\Temp'; Permissions = @('I', $null) }
                    @{ Path = 'BEL\Temp\Cache'; Permissions = @('L', $null) }
                    @{ Path = 'BEL\Dead'; Permissions = @('L', $null) }
                )

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $warn = $result | Where-Object Name -EQ 'Inaccessible folders'
                $warn | Should-BeTruthy
                $warn.Value | Should-MatchString ([regex]::Escape('BEL\Dead'))
                $warn.Value | Should-NotMatchString 'Temp'
            }

            It 'matches ignored subtrees case-insensitively and with trailing backslashes' {
                $perms = New-PermissionsSheet -ParentPermissions @('L', 'L') -FolderRows @(
                    @{ Path = 'BEL\TEMP\'; Permissions = @('I', $null) }
                    @{ Path = 'bel\temp\cache'; Permissions = @('L', $null) }
                )

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $result | Should-BeFalsy
            }
        }

        Context 'Error handling' {
            It 'rejects an empty array at parameter binding' {
                { Test-MatrixPermissionsHC -Permissions @() } | Should-Throw -ExceptionMessage '*empty array*'
            }
        }
    }

    Context 'Test-MatrixFormDataHC' {
        BeforeAll {
            # A single, fully-valid FormData row. Negative cases call this and
            # mutate one property so each test isolates exactly one failure.
            function New-ValidFormDataRow {
                [pscustomobject]@{
                    MatrixFormStatus        = 'Enabled'
                    MatrixCategoryName      = 'Default'
                    MatrixSubCategoryName   = 'General'
                    MatrixResponsible       = 'owner@example.com'
                    MatrixFolderDisplayName = 'Finance'
                    MatrixFolderPath        = 'E:\Folder'
                }
            }
        }

        Context 'when FormData is missing' {
            It 'returns a Warning when FormData is $null' {
                $result = Test-MatrixFormDataHC -FormData $null
                $result.Type | Should-Be 'Warning'
                $result.Name | Should-Be 'Missing FormData'
            }

            It 'returns a Warning when FormData is an empty array' {
                $result = Test-MatrixFormDataHC -FormData @()
                $result.Type | Should-Be 'Warning'
                $result.Name | Should-Be 'Missing FormData'
            }
        }

        Context 'row count' {
            It 'returns nothing for a single fully-valid row' {
                $result = Test-MatrixFormDataHC -FormData (New-ValidFormDataRow)
                $result | Should-BeFalsy
            }

            It 'flags a FatalError when more than one row is supplied' {
                $rows = @((New-ValidFormDataRow), (New-ValidFormDataRow))
                $result = Test-MatrixFormDataHC -FormData $rows
                $result.Type | Should-Be 'FatalError'
                $result.Name | Should-Be 'Incorrect row count'
                $result.Value | Should-Be 2
            }
        }

        Context 'mandatory column headers' {
            It 'flags a FatalError when a mandatory column is absent' {
                # MatrixFolderDisplayName omitted entirely.
                $row = [pscustomobject]@{
                    MatrixFormStatus      = 'Enabled'
                    MatrixCategoryName    = 'Default'
                    MatrixSubCategoryName = 'General'
                    MatrixResponsible     = 'owner@example.com'
                    MatrixFolderPath      = 'E:\Folder'
                }
                $result = Test-MatrixFormDataHC -FormData $row
                $result.Type | Should-Be 'FatalError'
                $result.Name | Should-Be 'Missing column header'
                $result.Value | Should-MatchString 'MatrixFolderDisplayName'
            }

            It 'lists every absent column in the Value' {
                $row = [pscustomobject]@{
                    MatrixFormStatus   = 'Enabled'
                    MatrixCategoryName = 'Default'
                }
                $result = Test-MatrixFormDataHC -FormData $row
                $result.Name | Should-Be 'Missing column header'
                $result.Value | Should-MatchString 'MatrixSubCategoryName'
                $result.Value | Should-MatchString 'MatrixResponsible'
                $result.Value | Should-MatchString 'MatrixFolderDisplayName'
                $result.Value | Should-MatchString 'MatrixFolderPath'
            }

            It 'flags absent columns even when the row is Disabled' {
                # The header check runs regardless of status.
                $row = [pscustomobject]@{
                    MatrixFormStatus   = 'Disabled'
                    MatrixCategoryName = 'Default'
                }
                $result = Test-MatrixFormDataHC -FormData $row
                $result.Type | Should-Be 'FatalError'
                $result.Name | Should-Be 'Missing column header'
            }
        }

        Context 'mandatory values when status is Enabled' {
            It 'flags a FatalError when an Enabled row has a blank value' {
                $row = New-ValidFormDataRow
                $row.MatrixResponsible = ''
                $result = Test-MatrixFormDataHC -FormData $row
                $result.Type | Should-Be 'FatalError'
                $result.Name | Should-Be 'Missing value'
                $result.Value | Should-MatchString 'MatrixResponsible'
            }

            It 'treats a whitespace-only value as blank' {
                $row = New-ValidFormDataRow
                $row.MatrixFolderPath = '   '
                $result = Test-MatrixFormDataHC -FormData $row
                $result.Name | Should-Be 'Missing value'
                $result.Value | Should-MatchString 'MatrixFolderPath'
            }

            It 'reports every blank mandatory value at once' {
                $row = New-ValidFormDataRow
                $row.MatrixResponsible = ''
                $row.MatrixFolderPath = ''
                $result = Test-MatrixFormDataHC -FormData $row
                $result.Name | Should-Be 'Missing value'
                $result.Value | Should-MatchString 'MatrixResponsible'
                $result.Value | Should-MatchString 'MatrixFolderPath'
            }

            It 'matches the Enabled status case-insensitively' {
                # 'enabled' still triggers the value check (PowerShell -eq is
                # case-insensitive), so a blank value is still flagged.
                $row = New-ValidFormDataRow
                $row.MatrixFormStatus = 'enabled'
                $row.MatrixResponsible = ''
                $result = Test-MatrixFormDataHC -FormData $row
                $result.Name | Should-Be 'Missing value'
            }
        }

        Context 'when status is not Enabled' {
            It 'skips the value checks for a Disabled row with blank values' {
                $row = New-ValidFormDataRow
                $row.MatrixFormStatus = 'Disabled'
                $row.MatrixResponsible = ''
                $row.MatrixFolderPath = ''
                $row.MatrixFolderDisplayName = ''
                $result = Test-MatrixFormDataHC -FormData $row
                $result | Should-BeFalsy
            }

            It 'treats a blank status as "not Enabled" and skips value checks' {
                # Documents current behavior: only the literal 'Enabled' triggers
                # value validation, so a blank/typo status silently passes.
                $row = New-ValidFormDataRow
                $row.MatrixFormStatus = ''
                $result = Test-MatrixFormDataHC -FormData $row
                $result | Should-BeFalsy
            }
        }
    }

    Context 'Test-MatrixSettingRowHC' {
        It 'Validates missing properties' {
            $S = @{ }
            $r = Test-MatrixSettingRowHC -SettingRow $S
            $r.Type | Should-ContainCollection 'FatalError'
        }
    }

    Describe 'Test-ConfigurationStructureHC' {
        BeforeAll {
            $script:ValidFolder = Join-Path 'TestDrive:' 'MatrixFolder'
            $script:ValidDefaults = Join-Path 'TestDrive:' 'defaults.json'
            $script:ValidLogDir = Join-Path 'TestDrive:' 'Logs'
            New-Item -ItemType Directory -Path $script:ValidFolder -Force | Out-Null
            New-Item -ItemType Directory -Path $script:ValidLogDir -Force | Out-Null
            Set-Content -Path $script:ValidDefaults -Value '{}' -Force

            function ConvertTo-JsonObject {
                param([Parameter(Mandatory)]$Hashtable)
                $Hashtable | ConvertTo-Json -Depth 20 | ConvertFrom-Json
            }

            function Set-ValidPaths {
                param([Parameter(Mandatory)][hashtable]$Json)
                # Only fill a branch if it exists. Missing-property fixtures remove
                # whole top-level blocks (e.g. Matrix, Settings); those tests assert on
                # the absence and do not need valid paths underneath.
                if ($Json.ContainsKey('Matrix')) {
                    $Json.Matrix.FolderPath = $script:ValidFolder
                    $Json.Matrix.DefaultsFile = $script:ValidDefaults
                }
                if ($Json.ContainsKey('Settings')) {
                    $Json.Settings.SaveLogFiles.Where.Folder = $script:ValidLogDir
                }
                return $Json
            }

            function Invoke-Validation {
                param([Parameter(Mandatory)][hashtable]$Json)

                $errors = [System.Collections.Generic.List[object]]::new()
                $obj = ConvertTo-JsonObject -Hashtable $Json
                Test-ConfigurationStructureHC -Json $obj -SystemErrors ([ref]$errors)
                return $errors
            }

            function Get-ErrorNames {
                param($Errors)
                @($Errors | ForEach-Object { $_.Name })
            }
        }

        Context 'Happy path' {
            It 'records no errors for a fully valid configuration' {
                $json = Set-ValidPaths (New-JsonFixture)
                $errors = Invoke-Validation -Json $json

                $errors.Count | Should-Be 0
            }
        }

        Context 'Top-level required properties' {
            It "records a 'Missing <_>' error when <_> is absent" -ForEach @(
                'Matrix', 'Export', 'ServiceNow', 'MaxConcurrent', 'PSSessionConfiguration', 'Settings'
            ) {
                $json = Set-ValidPaths (New-JsonFixtureWithMissingProperty -Property $_)
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-ContainCollection "Missing '$_'"
            }
        }

        Context 'Settings block' {
            It 'flags non-boolean Settings.SaveLogFiles.Detailed' {
                $json = Set-ValidPaths (New-JsonFixtureWithInvalidBoolean -Path 'Settings.SaveLogFiles.Detailed')
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-ContainCollection "Incorrect 'Settings.SaveLogFiles.Detailed'"
            }

            It 'reports a non-boolean Settings.SaveLogFiles.Detailed exactly once' {
                # Regression: a duplicate check produced two identical errors.
                $json = Set-ValidPaths (New-JsonFixtureWithInvalidBoolean -Path 'Settings.SaveLogFiles.Detailed')
                $errors = Invoke-Validation -Json $json

                @(Get-ErrorNames $errors).Where(
                    { $_ -eq "Incorrect 'Settings.SaveLogFiles.Detailed'" }
                ).Count | Should-Be 1
            }

            It 'reports a missing Settings.SaveLogFiles.Detailed as missing, not as incorrect' {
                <# Regression: the duplicate check tested $null -isnot [bool],
                so an absent value was reported both as missing and as not
                being a boolean.

                New-JsonFixtureWithMissingProperty only removes top-level
                keys, so the nested one is removed directly here. #>
                $json = Set-ValidPaths (New-JsonFixture)
                $json.Settings.SaveLogFiles.Remove('Detailed')

                $errors = Invoke-Validation -Json $json
                $names = @(Get-ErrorNames $errors)

                $names | Should-ContainCollection "Missing 'Settings.SaveLogFiles.Detailed'"
                $names | Should-NotContainCollection "Incorrect 'Settings.SaveLogFiles.Detailed'"
            }

            It 'accepts Settings.SaveLogFiles.Detailed set to false' {
                <# Guard: the missing check must stay '$null -eq', not '-not'.
                PowerShell coerces $false to falsy, so '-not' would report a
                legitimate 'false' as a missing value. #>
                $json = Set-ValidPaths (New-JsonFixture)
                $json.Settings.SaveLogFiles.Detailed = $false

                $errors = Invoke-Validation -Json $json

                $errors.Count | Should-Be 0
            }

            It 'flags non-boolean Settings.SaveInEventLog.Save' {
                $json = Set-ValidPaths (New-JsonFixtureWithInvalidBoolean -Path 'Settings.SaveInEventLog.Save')
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-ContainCollection "Incorrect 'Settings.SaveInEventLog.Save'"
            }

            It 'flags missing Settings.SaveLogFiles.Where.Folder' {
                $json = New-JsonFixtureWithModifiedValue -Path 'Settings.SaveLogFiles.Where.Folder' -Value ''
                $json.Matrix.FolderPath = $script:ValidFolder
                $json.Matrix.DefaultsFile = $script:ValidDefaults
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-ContainCollection "Missing 'Settings.SaveLogFiles.Where.Folder'"
            }

            It 'flags missing Settings.ScriptName' {
                $json = Set-ValidPaths (New-JsonFixtureWithModifiedValue -Path 'Settings.ScriptName' -Value '')
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-ContainCollection "Missing 'Settings.ScriptName'"
            }
        }

        Context 'Settings.SendMail nested block' {
            It 'flags missing SendMail.From' {
                $json = Set-ValidPaths (New-JsonFixtureWithModifiedValue -Path 'Settings.SendMail.From' -Value '')
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-ContainCollection "Missing 'Settings.SendMail.From'"
            }

            It 'flags missing SendMail.Body' {
                # $null Body triggers the "Missing 'Settings.SendMail.Body'" check.
                # The builder's -Value is Mandatory and rejects $null, so set it
                # directly on the hashtable instead.
                $json = Set-ValidPaths (New-JsonFixture)
                $json.Settings.SendMail.Body = $null
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-ContainCollection "Missing 'Settings.SendMail.Body'"
            }

            It 'flags non-numeric SendMail.Smtp.Port' {
                $json = Set-ValidPaths (New-JsonFixtureWithModifiedValue -Path 'Settings.SendMail.Smtp.Port' -Value 'abc')
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-ContainCollection "Incorrect 'SendMail.Smtp.Port'"
            }

            It 'flags an invalid SendMail.Smtp.ConnectionType' {
                $json = Set-ValidPaths (New-JsonFixtureWithModifiedValue -Path 'Settings.SendMail.Smtp.ConnectionType' -Value 'Carrier Pigeon')
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-ContainCollection "Incorrect 'Settings.SendMail.Smtp.ConnectionType'"
            }

            It 'accepts every valid ConnectionType <_>' -ForEach @(
                'None', 'Auto', 'SslOnConnect', 'StartTls', 'StartTlsWhenAvailable'
            ) {
                $json = Set-ValidPaths (New-JsonFixtureWithModifiedValue -Path 'Settings.SendMail.Smtp.ConnectionType' -Value $_)
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-NotContainCollection "Incorrect 'Settings.SendMail.Smtp.ConnectionType'"
            }

            It 'flags a completely missing SendMail block as mandatory' {
                # "Completely missing" = the key is absent. The builder's -Value is
                # Mandatory and rejects $null, so remove the key on the hashtable.
                $json = Set-ValidPaths (New-JsonFixture)
                $json.Settings.Remove('SendMail') | Out-Null
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-ContainCollection "Missing 'Settings.SendMail'"
            }
        }

        Context 'Matrix block' {
            It 'flags missing Matrix.FolderPath' {
                $json = New-JsonFixtureWithModifiedValue -Path 'Matrix.FolderPath' -Value ''
                $json.Matrix.DefaultsFile = $script:ValidDefaults
                $json.Settings.SaveLogFiles.Where.Folder = $script:ValidLogDir
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-ContainCollection "Missing 'Matrix.FolderPath'"
            }

            It 'flags a non-existent Matrix.FolderPath' {
                $json = New-JsonFixtureWithModifiedValue -Path 'Matrix.FolderPath' -Value 'TestDrive:\does\not\exist'
                $json.Matrix.DefaultsFile = $script:ValidDefaults
                $json.Settings.SaveLogFiles.Where.Folder = $script:ValidLogDir
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-ContainCollection "Incorrect 'Matrix.FolderPath'"
            }

            It 'flags missing Matrix.DefaultsFile' {
                $json = New-JsonFixtureWithModifiedValue -Path 'Matrix.DefaultsFile' -Value ''
                $json.Matrix.FolderPath = $script:ValidFolder
                $json.Settings.SaveLogFiles.Where.Folder = $script:ValidLogDir
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-ContainCollection "Missing 'Matrix.DefaultsFile'"
            }

            It 'flags a non-existent Matrix.DefaultsFile' {
                $json = New-JsonFixtureWithModifiedValue -Path 'Matrix.DefaultsFile' -Value 'TestDrive:\nope.json'
                $json.Matrix.FolderPath = $script:ValidFolder
                $json.Settings.SaveLogFiles.Where.Folder = $script:ValidLogDir
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-ContainCollection "Incorrect 'Matrix.DefaultsFile'"
            }

            It 'flags a non-array Matrix.AdGroupPlaceHolders' {
                $json = Set-ValidPaths (New-JsonFixtureWithInvalidArray -Path 'Matrix.AdGroupPlaceHolders')
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-ContainCollection "Incorrect 'Matrix.AdGroupPlaceHolders'"
            }

            It 'flags a non-boolean Matrix.Archive' {
                $json = Set-ValidPaths (New-JsonFixtureWithInvalidBoolean -Path 'Matrix.Archive')
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-ContainCollection "Incorrect 'Matrix.Archive'"
            }
        }

        Context 'MaxConcurrent block' {
            It 'flags non-numeric MaxConcurrent.<_>' -ForEach @(
                'JobsTotal', 'JobsPerComputer', 'FoldersPerMatrix'
            ) {
                $json = Set-ValidPaths (New-JsonFixtureWithInvalidInteger -Path "MaxConcurrent.$_")
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-ContainCollection "Incorrect 'MaxConcurrent.$_'"
            }

            It 'flags JobsPerComputer greater than JobsTotal' {
                # A per-computer cap above the total is unreachable: the overall
                # throttle stops the run before one computer could ever get
                # there, so the configured number would quietly mean something
                # other than what it says.
                $json = Set-ValidPaths (New-JsonFixture)
                $json.MaxConcurrent.JobsTotal = 5
                $json.MaxConcurrent.JobsPerComputer = 10

                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-ContainCollection "Incorrect 'MaxConcurrent.JobsPerComputer'"
            }

            It 'accepts JobsPerComputer equal to JobsTotal' {
                # One computer being allowed to consume the entire budget is a
                # legitimate configuration, not an error.
                $json = Set-ValidPaths (New-JsonFixture)
                $json.MaxConcurrent.JobsTotal = 5
                $json.MaxConcurrent.JobsPerComputer = 5

                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-NotContainCollection "Incorrect 'MaxConcurrent.JobsPerComputer'"
            }
        }

        Context 'Export block' {
            It 'flags PermissionsExcelFile not ending in .xlsx' {
                $json = Set-ValidPaths (New-JsonFixtureWithModifiedValue -Path 'Export.PermissionsExcelFile' -Value 'out.csv')
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-ContainCollection "Incorrect 'Export.PermissionsExcelFile'"
            }

            It 'flags OverviewHtmlFile not ending in .html' {
                $json = Set-ValidPaths (New-JsonFixtureWithModifiedValue -Path 'Export.OverviewHtmlFile' -Value 'report.pdf')
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-ContainCollection "Incorrect 'Export.OverviewHtmlFile'"
            }

            It 'flags ServiceNowFormDataExcelFile not ending in .xlsx' {
                $json = Set-ValidPaths (New-JsonFixtureWithModifiedValue -Path 'Export.ServiceNowFormDataExcelFile' -Value 'forms.csv')
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-ContainCollection "Incorrect 'Export.ServiceNowFormDataExcelFile'"
            }
        }

        Context 'Export.ServiceNowFormDataExcelFile cross-dependency on ServiceNow' {
            # The Export region now correctly reads $Json.ServiceNow. When a
            # ServiceNowFormDataExcelFile is set, ServiceNow must exist and have
            # CredentialsFilePath / TableName / Environment populated.

            It "emits 'Incorrect configuration' when ServiceNow is absent" {
                $json = Set-ValidPaths (New-JsonFixtureWithModifiedValue -Path 'Export.ServiceNowFormDataExcelFile' -Value 'forms.xlsx')
                $json.Remove('ServiceNow') | Out-Null
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-ContainCollection 'Incorrect configuration'
            }

            It 'records no ServiceNow errors when the block is present and fully populated' {
                $json = Set-ValidPaths (New-JsonFixtureWithModifiedValue -Path 'Export.ServiceNowFormDataExcelFile' -Value 'forms.xlsx')
                # The fixture ships CredentialsFilePath = '' (blank), so fill it to
                # get a genuinely complete ServiceNow block.
                $json.ServiceNow.CredentialsFilePath = 'TestDrive:\snow.cred'
                $errors = Invoke-Validation -Json $json
                $names = Get-ErrorNames $errors

                $names | Should-NotContainCollection 'Incorrect configuration'
                $names | Should-NotContainCollection "Missing 'ServiceNow.CredentialsFilePath'"
                $names | Should-NotContainCollection "Missing 'ServiceNow.TableName'"
                $names | Should-NotContainCollection "Missing 'ServiceNow.Environment'"
            }

            It 'flags missing ServiceNow.<_> when that property is blank' -ForEach @(
                'CredentialsFilePath', 'TableName', 'Environment'
            ) {
                $json = Set-ValidPaths (New-JsonFixtureWithModifiedValue -Path 'Export.ServiceNowFormDataExcelFile' -Value 'forms.xlsx')
                # Start from a complete block, then blank the one under test.
                $json.ServiceNow.CredentialsFilePath = 'TestDrive:\snow.cred'
                $json.ServiceNow.$_ = ''
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-ContainCollection "Missing 'ServiceNow.$_'"
            }
        }

        Context 'SharePoint block' {
            BeforeAll {
                # SharePoint is optional, so it is absent from New-JsonFixture.
                # Tests that need it add it here rather than changing the shared
                # fixture, which would make every other test in this file carry a
                # block it does not care about.
                function Add-SharePointBlock {
                    param(
                        [Parameter(Mandatory)][hashtable]$Json,
                        [hashtable]$Override = @{}
                    )

                    $block = @{
                        SiteUrl               = 'https://contoso.sharepoint.com/sites/IT'
                        DocumentLibraryName   = 'Documents'
                        ClientId              = 'ENV:AZURE_CLIENT_ID'
                        TenantId              = 'ENV:AZURE_TENANT_ID'
                        CertificateThumbprint = 'ENV:AZURE_POWERSHELL_CERTIFICATE_THUMBPRINT'
                    }

                    foreach ($key in $Override.Keys) {
                        $block[$key] = $Override[$key]
                    }

                    $Json.SharePoint = $block
                    $Json
                }
            }

            It 'records no errors when the SharePoint block is absent entirely' {
                # The most important case in this Context: SharePoint was added to
                # the schema after configurations were already in production, and
                # it must stay optional. If this fails, every existing config
                # stops loading.
                $json = Set-ValidPaths (New-JsonFixture)

                $json.ContainsKey('SharePoint') | Should-BeFalse

                $errors = Invoke-Validation -Json $json

                $errors.Count | Should-Be 0
            }

            It 'records no errors when SharePoint is present but SiteUrl is empty' {
                # A block left in place with the upload switched off must not be
                # treated as a misconfiguration.
                $json = Set-ValidPaths (New-JsonFixture)
                $json = Add-SharePointBlock -Json $json -Override @{ SiteUrl = '' }
                $json.Export.OverviewHtmlFile = $null

                $errors = Invoke-Validation -Json $json

                $errors.Count | Should-Be 0
            }

            It 'records no errors for a complete SharePoint block' {
                $json = Set-ValidPaths (New-JsonFixture)
                $json = Add-SharePointBlock -Json $json
                $json.Export.OverviewHtmlFile = 'TestDrive:\Overview.html'

                $errors = Invoke-Validation -Json $json

                $errors.Count | Should-Be 0
            }

            It 'accepts an optional FolderPath and FileName' {
                $json = Set-ValidPaths (New-JsonFixture)
                $json = Add-SharePointBlock -Json $json -Override @{
                    FolderPath = 'Reports/Permission matrix'
                    FileName   = 'Permission matrix overview.html'
                }
                $json.Export.OverviewHtmlFile = 'TestDrive:\Overview.html'

                $errors = Invoke-Validation -Json $json

                $errors.Count | Should-Be 0
            }

            It 'flags a SiteUrl that does not start with https://' -ForEach @(
                'http://contoso.sharepoint.com/sites/IT'
                'contoso.sharepoint.com/sites/IT'
                'ftp://contoso.sharepoint.com'
            ) {
                $json = Set-ValidPaths (New-JsonFixture)
                $json = Add-SharePointBlock -Json $json -Override @{ SiteUrl = $_ }
                $json.Export.OverviewHtmlFile = 'TestDrive:\Overview.html'

                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-ContainCollection "Incorrect 'SharePoint.SiteUrl'"
            }

            It 'flags missing SharePoint.<_> when that property is blank' -ForEach @(
                'DocumentLibraryName', 'ClientId', 'TenantId', 'CertificateThumbprint'
            ) {
                $json = Set-ValidPaths (New-JsonFixture)
                $json = Add-SharePointBlock -Json $json -Override @{ $_ = '' }
                $json.Export.OverviewHtmlFile = 'TestDrive:\Overview.html'

                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-ContainCollection "Missing 'SharePoint.$_'"
            }

            It 'flags every missing property at once rather than stopping at the first' {
                # The validation loop must report all four, so a user fixing the
                # config sees the whole list in one run instead of rediscovering
                # them one at a time.
                $json = Set-ValidPaths (New-JsonFixture)
                $json = Add-SharePointBlock -Json $json -Override @{
                    DocumentLibraryName   = ''
                    ClientId              = ''
                    TenantId              = ''
                    CertificateThumbprint = ''
                }
                $json.Export.OverviewHtmlFile = 'TestDrive:\Overview.html'

                $errors = Invoke-Validation -Json $json
                $names = Get-ErrorNames $errors

                $names | Should-ContainCollection "Missing 'SharePoint.DocumentLibraryName'"
                $names | Should-ContainCollection "Missing 'SharePoint.ClientId'"
                $names | Should-ContainCollection "Missing 'SharePoint.TenantId'"
                $names | Should-ContainCollection "Missing 'SharePoint.CertificateThumbprint'"
            }

            It 'flags a SharePoint upload configured without Export.OverviewHtmlFile' {
                # There would be nothing to upload: the html the whole feature
                # exists to publish is never produced.
                $json = Set-ValidPaths (New-JsonFixture)
                $json = Add-SharePointBlock -Json $json
                $json.Export.OverviewHtmlFile = $null

                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-ContainCollection 'Incorrect configuration'
            }

            It 'does not flag the OverviewHtmlFile dependency when SharePoint is unused' {
                $json = Set-ValidPaths (New-JsonFixture)
                $json.Export.OverviewHtmlFile = $null

                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should-NotContainCollection 'Incorrect configuration'
            }

            It 'raises FatalError, not Warning, for an incomplete SharePoint block' {
                # These are fatal on purpose: a half-configured upload should stop
                # the run at validation rather than fail at 02:00 after the
                # permissions have already been applied.
                $json = Set-ValidPaths (New-JsonFixture)
                $json = Add-SharePointBlock -Json $json -Override @{ ClientId = '' }
                $json.Export.OverviewHtmlFile = 'TestDrive:\Overview.html'

                $errors = Invoke-Validation -Json $json

                $sharePointErrors = @($errors | Where-Object { $_.Name -like '*SharePoint*' })

                $sharePointErrors.Count | Should-BeGreaterThan 0
                $sharePointErrors.Type | Should-NotContainCollection 'Warning'
                $sharePointErrors.Type | Select-Object -Unique | Should-BeString 'FatalError' -CaseSensitive
            }
        }
    }
}