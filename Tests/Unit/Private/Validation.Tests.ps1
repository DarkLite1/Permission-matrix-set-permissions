#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
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

    Context 'Test-MatrixFileHC' {
        It 'Warns for missing settings' {
            $M = @{ Settings = @(); Permissions = @('x') }
            $res = Test-MatrixFileHC -MatrixObject $M
            $res.Type | Should -Contain 'Warning'
        }

        It 'Errors for missing permissions' {
            $M = @{ Settings = @('x'); Permissions = @() }
            $res = Test-MatrixFileHC -MatrixObject $M
            $res.Type | Should -Contain 'FatalError'
        }
    }

    Context 'Test-MatrixPermissionsHC' {

        Context 'Happy path' {
            It 'returns nothing when the Valid fixture is supplied' {
                $perms = Get-RoundTripPermissions -Scenario 'Valid'

                $result = Test-MatrixPermissionsHC -Permissions $perms

                # Function only returns $checks when Count -gt 0, so success => $null.
                $result | Should -BeNullOrEmpty
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

                $result | Should -Not -BeNullOrEmpty
                ($result.Name) | Should -Contain $Expected
            }
        }

        Context 'fatal errors exit immediately' {
            It 'returns ONLY "Missing rows" for the MissingRows fixture' {
                $perms = Get-RoundTripPermissions -Scenario 'MissingRows'

                $result = Test-MatrixPermissionsHC -Permissions $perms

                @($result).Count | Should -Be 1
                $result[0].Type | Should -Be 'FatalError'
                $result[0].Name | Should -Be 'Missing rows'
            }

            It 'returns ONLY "Missing columns" for the MissingColumns fixture' {
                $perms = Get-RoundTripPermissions -Scenario 'MissingColumns'

                $result = Test-MatrixPermissionsHC -Permissions $perms

                @($result).Count | Should -Be 1
                $result[0].Type | Should -Be 'FatalError'
                $result[0].Name | Should -Be 'Missing columns'
            }
        }

        Context 'Check types are correct' {
            It 'classifies InaccessibleFolders as a Warning, not a FatalError' {
                $perms = Get-RoundTripPermissions -Scenario 'InaccessibleFolders'
                $result = Test-MatrixPermissionsHC -Permissions $perms

                $warn = $result | Where-Object Name -EQ 'Inaccessible folders'
                $warn | Should -Not -BeNullOrEmpty
                $warn.Type | Should -Be 'Warning'
            }

            It 'classifies InvalidPermissionChar as a FatalError' {
                $perms = Get-RoundTripPermissions -Scenario 'InvalidPermissionChar'
                $result = Test-MatrixPermissionsHC -Permissions $perms

                $err = $result | Where-Object Name -EQ 'Invalid permission character'
                $err | Should -Not -BeNullOrEmpty
                $err.Type | Should -Be 'FatalError'
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

                $result | Should -BeNullOrEmpty
            }

            It 'flags a genuinely deepest folder with only List permissions' {
                $perms = New-PermissionsSheet -ParentPermissions @('L', 'L') -FolderRows @(
                    @{ Path = 'BEL\Marketing'; Permissions = @('L', $null) }
                )

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $warn = $result | Where-Object Name -EQ 'Inaccessible folders'
                $warn | Should -Not -BeNullOrEmpty
                $warn.Value | Should -Match ([regex]::Escape('BEL\Marketing'))
            }

            It 'reports only the truly inaccessible folder, not the trailing-backslash parent' {
                $perms = New-PermissionsSheet -ParentPermissions @('L', 'L') -FolderRows @(
                    @{ Path = 'BEL\L&D\Certificates\'; Permissions = @('L', 'L') }
                    @{ Path = 'BEL\L&D\Certificates\AGG'; Permissions = @('W', 'W') }
                    @{ Path = 'BEL\Dead'; Permissions = @('L', $null) }
                )

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $warn = $result | Where-Object Name -EQ 'Inaccessible folders'
                $warn | Should -Not -BeNullOrEmpty
                $warn.Value | Should -Match ([regex]::Escape('BEL\Dead'))
                $warn.Value | Should -Not -Match 'Certificates'
            }

            It 'matches parent and child paths case-insensitively' {
                $perms = New-PermissionsSheet -ParentPermissions @('L', 'L') -FolderRows @(
                    @{ Path = 'BEL\ARCHIVE'; Permissions = @('L', 'L') }
                    @{ Path = 'bel\archive\2020'; Permissions = @('W', $null) }
                )

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $result | Should -BeNullOrEmpty
            }

            It 'handles wildcard characters in folder names' {
                # '-like' would treat '[2026]' as a wildcard set and break
                # parent/child matching; String.StartsWith must not.
                $perms = New-PermissionsSheet -ParentPermissions @('L', 'L') -FolderRows @(
                    @{ Path = 'BEL\Reports[2026]'; Permissions = @('L', 'L') }
                    @{ Path = 'BEL\Reports[2026]\Q1'; Permissions = @('W', $null) }
                )

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $result | Should -BeNullOrEmpty
            }

            It 'flags a List-only deepest folder even when the parent row grants access' {
                # An explicit 'L' sets a List-only ACL on the folder itself,
                # so the parent's W cannot make it accessible.
                $perms = New-PermissionsSheet -ParentPermissions @('W', 'L') -FolderRows @(
                    @{ Path = 'BEL\Marketing'; Permissions = @('L', $null) }
                )

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $warn = $result | Where-Object Name -EQ 'Inaccessible folders'
                $warn | Should -Not -BeNullOrEmpty
                $warn.Value | Should -Match ([regex]::Escape('BEL\Marketing'))
            }

            It 'does not flag a blank deepest folder when the parent row grants access' {
                # A row without any permission inherits the parent ACL, so
                # the parent's W makes it accessible.
                $perms = New-PermissionsSheet -ParentPermissions @('W', 'L') -FolderRows @(
                    @{ Path = 'BEL\Marketing'; Permissions = @($null, $null) }
                )

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $result | Should -BeNullOrEmpty
            }

            It 'flags a blank deepest folder when the parent row grants no access' {
                # Inheriting from a parent that only grants L leaves the
                # folder without read or write access.
                $perms = New-PermissionsSheet -ParentPermissions @('L', 'L') -FolderRows @(
                    @{ Path = 'BEL\Marketing'; Permissions = @($null, $null) }
                )

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $warn = $result | Where-Object Name -EQ 'Inaccessible folders'
                $warn | Should -Not -BeNullOrEmpty
                $warn.Value | Should -Match ([regex]::Escape('BEL\Marketing'))
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
                $warn | Should -Not -BeNullOrEmpty
                $warn.Value | Should -Match ([regex]::Escape('BEL\App'))
            }

            It 'reports the folder exactly as typed in Excel, trailing backslash included' {
                $perms = New-PermissionsSheet -ParentPermissions @('L', 'L') -FolderRows @(
                    @{ Path = 'BEL\Lonely\'; Permissions = @('L', $null) }
                )

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $warn = $result | Where-Object Name -EQ 'Inaccessible folders'
                $warn | Should -Not -BeNullOrEmpty
                $warn.Value | Should -Match ([regex]::Escape('BEL\Lonely\'))
            }

            It 'does not flag a deepest folder marked with I (Ignore)' {
                $perms = New-PermissionsSheet -ParentPermissions @('L', 'L') -FolderRows @(
                    @{ Path = 'BEL\Temp'; Permissions = @('I', $null) }
                )

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $result | Should -BeNullOrEmpty
            }

            It 'does not flag subfolders of a folder marked with I (Ignore)' {
                # 'BEL\Temp\Cache' is List-only and deepest, but it sits
                # under an ignored folder, so the matrix does not manage it.
                $perms = New-PermissionsSheet -ParentPermissions @('L', 'L') -FolderRows @(
                    @{ Path = 'BEL\Temp'; Permissions = @('I', $null) }
                    @{ Path = 'BEL\Temp\Cache'; Permissions = @('L', $null) }
                )

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $result | Should -BeNullOrEmpty
            }

            It 'still flags inaccessible folders outside an ignored subtree' {
                $perms = New-PermissionsSheet -ParentPermissions @('L', 'L') -FolderRows @(
                    @{ Path = 'BEL\Temp'; Permissions = @('I', $null) }
                    @{ Path = 'BEL\Temp\Cache'; Permissions = @('L', $null) }
                    @{ Path = 'BEL\Dead'; Permissions = @('L', $null) }
                )

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $warn = $result | Where-Object Name -EQ 'Inaccessible folders'
                $warn | Should -Not -BeNullOrEmpty
                $warn.Value | Should -Match ([regex]::Escape('BEL\Dead'))
                $warn.Value | Should -Not -Match 'Temp'
            }

            It 'matches ignored subtrees case-insensitively and with trailing backslashes' {
                $perms = New-PermissionsSheet -ParentPermissions @('L', 'L') -FolderRows @(
                    @{ Path = 'BEL\TEMP\'; Permissions = @('I', $null) }
                    @{ Path = 'bel\temp\cache'; Permissions = @('L', $null) }
                )

                $result = Test-MatrixPermissionsHC -Permissions $perms

                $result | Should -BeNullOrEmpty
            }
        }

        Context 'Error handling' {
            It 'rejects an empty array at parameter binding' {
                { Test-MatrixPermissionsHC -Permissions @() } |
                Should -Throw -ExpectedMessage '*empty array*'
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
                $result.Type | Should -Be 'Warning'
                $result.Name | Should -Be 'Missing FormData'
            }

            It 'returns a Warning when FormData is an empty array' {
                $result = Test-MatrixFormDataHC -FormData @()
                $result.Type | Should -Be 'Warning'
                $result.Name | Should -Be 'Missing FormData'
            }
        }

        Context 'row count' {
            It 'returns nothing for a single fully-valid row' {
                $result = Test-MatrixFormDataHC -FormData (New-ValidFormDataRow)
                $result | Should -BeNullOrEmpty
            }

            It 'flags a FatalError when more than one row is supplied' {
                $rows = @((New-ValidFormDataRow), (New-ValidFormDataRow))
                $result = Test-MatrixFormDataHC -FormData $rows
                $result.Type | Should -Be 'FatalError'
                $result.Name | Should -Be 'Incorrect row count'
                $result.Value | Should -Be 2
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
                $result.Type | Should -Be 'FatalError'
                $result.Name | Should -Be 'Missing column header'
                $result.Value | Should -Match 'MatrixFolderDisplayName'
            }

            It 'lists every absent column in the Value' {
                $row = [pscustomobject]@{
                    MatrixFormStatus   = 'Enabled'
                    MatrixCategoryName = 'Default'
                }
                $result = Test-MatrixFormDataHC -FormData $row
                $result.Name | Should -Be 'Missing column header'
                $result.Value | Should -Match 'MatrixSubCategoryName'
                $result.Value | Should -Match 'MatrixResponsible'
                $result.Value | Should -Match 'MatrixFolderDisplayName'
                $result.Value | Should -Match 'MatrixFolderPath'
            }

            It 'flags absent columns even when the row is Disabled' {
                # The header check runs regardless of status.
                $row = [pscustomobject]@{
                    MatrixFormStatus   = 'Disabled'
                    MatrixCategoryName = 'Default'
                }
                $result = Test-MatrixFormDataHC -FormData $row
                $result.Type | Should -Be 'FatalError'
                $result.Name | Should -Be 'Missing column header'
            }
        }

        Context 'mandatory values when status is Enabled' {
            It 'flags a FatalError when an Enabled row has a blank value' {
                $row = New-ValidFormDataRow
                $row.MatrixResponsible = ''
                $result = Test-MatrixFormDataHC -FormData $row
                $result.Type | Should -Be 'FatalError'
                $result.Name | Should -Be 'Missing value'
                $result.Value | Should -Match 'MatrixResponsible'
            }

            It 'treats a whitespace-only value as blank' {
                $row = New-ValidFormDataRow
                $row.MatrixFolderPath = '   '
                $result = Test-MatrixFormDataHC -FormData $row
                $result.Name | Should -Be 'Missing value'
                $result.Value | Should -Match 'MatrixFolderPath'
            }

            It 'reports every blank mandatory value at once' {
                $row = New-ValidFormDataRow
                $row.MatrixResponsible = ''
                $row.MatrixFolderPath = ''
                $result = Test-MatrixFormDataHC -FormData $row
                $result.Name | Should -Be 'Missing value'
                $result.Value | Should -Match 'MatrixResponsible'
                $result.Value | Should -Match 'MatrixFolderPath'
            }

            It 'matches the Enabled status case-insensitively' {
                # 'enabled' still triggers the value check (PowerShell -eq is
                # case-insensitive), so a blank value is still flagged.
                $row = New-ValidFormDataRow
                $row.MatrixFormStatus = 'enabled'
                $row.MatrixResponsible = ''
                $result = Test-MatrixFormDataHC -FormData $row
                $result.Name | Should -Be 'Missing value'
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
                $result | Should -BeNullOrEmpty
            }

            It 'treats a blank status as "not Enabled" and skips value checks' {
                # Documents current behavior: only the literal 'Enabled' triggers
                # value validation, so a blank/typo status silently passes.
                $row = New-ValidFormDataRow
                $row.MatrixFormStatus = ''
                $result = Test-MatrixFormDataHC -FormData $row
                $result | Should -BeNullOrEmpty
            }
        }
    }

    Context 'Test-MatrixSettingRowHC' {
        It 'Validates missing properties' {
            $S = @{ }
            $r = Test-MatrixSettingRowHC -SettingRow $S
            $r.Type | Should -Contain 'FatalError'
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

                $errors.Count | Should -Be 0
            }
        }

        Context 'Top-level required properties' {
            It "records a 'Missing <_>' error when <_> is absent" -ForEach @(
                'Matrix', 'Export', 'ServiceNow', 'MaxConcurrent', 'PSSessionConfiguration', 'Settings'
            ) {
                $json = Set-ValidPaths (New-JsonFixtureWithMissingProperty -Property $_)
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should -Contain "Missing '$_'"
            }
        }

        Context 'Settings block' {
            It 'flags non-boolean Settings.SaveLogFiles.Detailed' {
                $json = Set-ValidPaths (New-JsonFixtureWithInvalidBoolean -Path 'Settings.SaveLogFiles.Detailed')
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should -Contain "Incorrect 'Settings.SaveLogFiles.Detailed'"
            }

            It 'flags non-boolean Settings.SaveInEventLog.Save' {
                $json = Set-ValidPaths (New-JsonFixtureWithInvalidBoolean -Path 'Settings.SaveInEventLog.Save')
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should -Contain "Incorrect 'Settings.SaveInEventLog.Save'"
            }

            It 'flags missing Settings.SaveLogFiles.Where.Folder' {
                $json = New-JsonFixtureWithModifiedValue -Path 'Settings.SaveLogFiles.Where.Folder' -Value ''
                $json.Matrix.FolderPath = $script:ValidFolder
                $json.Matrix.DefaultsFile = $script:ValidDefaults
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should -Contain "Missing 'Settings.SaveLogFiles.Where.Folder'"
            }

            It 'flags missing Settings.ScriptName' {
                $json = Set-ValidPaths (New-JsonFixtureWithModifiedValue -Path 'Settings.ScriptName' -Value '')
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should -Contain "Missing 'Settings.ScriptName'"
            }
        }

        Context 'Settings.SendMail nested block' {
            It 'flags missing SendMail.From' {
                $json = Set-ValidPaths (New-JsonFixtureWithModifiedValue -Path 'Settings.SendMail.From' -Value '')
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should -Contain "Missing 'Settings.SendMail.From'"
            }

            It 'flags missing SendMail.Body' {
                # $null Body triggers the "Missing 'Settings.SendMail.Body'" check.
                # The builder's -Value is Mandatory and rejects $null, so set it
                # directly on the hashtable instead.
                $json = Set-ValidPaths (New-JsonFixture)
                $json.Settings.SendMail.Body = $null
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should -Contain "Missing 'Settings.SendMail.Body'"
            }

            It 'flags non-numeric SendMail.Smtp.Port' {
                $json = Set-ValidPaths (New-JsonFixtureWithModifiedValue -Path 'Settings.SendMail.Smtp.Port' -Value 'abc')
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should -Contain "Incorrect 'SendMail.Smtp.Port'"
            }

            It 'flags an invalid SendMail.Smtp.ConnectionType' {
                $json = Set-ValidPaths (New-JsonFixtureWithModifiedValue -Path 'Settings.SendMail.Smtp.ConnectionType' -Value 'Carrier Pigeon')
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should -Contain "Incorrect 'Settings.SendMail.Smtp.ConnectionType'"
            }

            It 'accepts every valid ConnectionType <_>' -ForEach @(
                'None', 'Auto', 'SslOnConnect', 'StartTls', 'StartTlsWhenAvailable'
            ) {
                $json = Set-ValidPaths (New-JsonFixtureWithModifiedValue -Path 'Settings.SendMail.Smtp.ConnectionType' -Value $_)
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should -Not -Contain "Incorrect 'Settings.SendMail.Smtp.ConnectionType'"
            }

            It 'flags a completely missing SendMail block as mandatory' {
                # "Completely missing" = the key is absent. The builder's -Value is
                # Mandatory and rejects $null, so remove the key on the hashtable.
                $json = Set-ValidPaths (New-JsonFixture)
                $json.Settings.Remove('SendMail') | Out-Null
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should -Contain "Missing 'Settings.SendMail'"
            }
        }

        Context 'Matrix block' {
            It 'flags missing Matrix.FolderPath' {
                $json = New-JsonFixtureWithModifiedValue -Path 'Matrix.FolderPath' -Value ''
                $json.Matrix.DefaultsFile = $script:ValidDefaults
                $json.Settings.SaveLogFiles.Where.Folder = $script:ValidLogDir
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should -Contain "Missing 'Matrix.FolderPath'"
            }

            It 'flags a non-existent Matrix.FolderPath' {
                $json = New-JsonFixtureWithModifiedValue -Path 'Matrix.FolderPath' -Value 'TestDrive:\does\not\exist'
                $json.Matrix.DefaultsFile = $script:ValidDefaults
                $json.Settings.SaveLogFiles.Where.Folder = $script:ValidLogDir
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should -Contain "Incorrect 'Matrix.FolderPath'"
            }

            It 'flags missing Matrix.DefaultsFile' {
                $json = New-JsonFixtureWithModifiedValue -Path 'Matrix.DefaultsFile' -Value ''
                $json.Matrix.FolderPath = $script:ValidFolder
                $json.Settings.SaveLogFiles.Where.Folder = $script:ValidLogDir
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should -Contain "Missing 'Matrix.DefaultsFile'"
            }

            It 'flags a non-existent Matrix.DefaultsFile' {
                $json = New-JsonFixtureWithModifiedValue -Path 'Matrix.DefaultsFile' -Value 'TestDrive:\nope.json'
                $json.Matrix.FolderPath = $script:ValidFolder
                $json.Settings.SaveLogFiles.Where.Folder = $script:ValidLogDir
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should -Contain "Incorrect 'Matrix.DefaultsFile'"
            }

            It 'flags a non-array Matrix.AdGroupPlaceHolders' {
                $json = Set-ValidPaths (New-JsonFixtureWithInvalidArray -Path 'Matrix.AdGroupPlaceHolders')
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should -Contain "Incorrect 'Matrix.AdGroupPlaceHolders'"
            }

            It 'flags a non-boolean Matrix.Archive' {
                $json = Set-ValidPaths (New-JsonFixtureWithInvalidBoolean -Path 'Matrix.Archive')
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should -Contain "Incorrect 'Matrix.Archive'"
            }
        }

        Context 'MaxConcurrent block' {
            It 'flags non-numeric MaxConcurrent.<_>' -ForEach @(
                'Computers', 'FoldersPerMatrix', 'JobsPerRemoteComputer'
            ) {
                $json = Set-ValidPaths (New-JsonFixtureWithInvalidInteger -Path "MaxConcurrent.$_")
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should -Contain "Incorrect 'MaxConcurrent.$_'"
            }
        }

        Context 'Export block' {
            It 'flags PermissionsExcelFile not ending in .xlsx' {
                $json = Set-ValidPaths (New-JsonFixtureWithModifiedValue -Path 'Export.PermissionsExcelFile' -Value 'out.csv')
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should -Contain "Incorrect 'Export.PermissionsExcelFile'"
            }

            It 'flags OverviewHtmlFile not ending in .html' {
                $json = Set-ValidPaths (New-JsonFixtureWithModifiedValue -Path 'Export.OverviewHtmlFile' -Value 'report.pdf')
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should -Contain "Incorrect 'Export.OverviewHtmlFile'"
            }

            It 'flags ServiceNowFormDataExcelFile not ending in .xlsx' {
                $json = Set-ValidPaths (New-JsonFixtureWithModifiedValue -Path 'Export.ServiceNowFormDataExcelFile' -Value 'forms.csv')
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should -Contain "Incorrect 'Export.ServiceNowFormDataExcelFile'"
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

                Get-ErrorNames $errors | Should -Contain 'Incorrect configuration'
            }

            It 'records no ServiceNow errors when the block is present and fully populated' {
                $json = Set-ValidPaths (New-JsonFixtureWithModifiedValue -Path 'Export.ServiceNowFormDataExcelFile' -Value 'forms.xlsx')
                # The fixture ships CredentialsFilePath = '' (blank), so fill it to
                # get a genuinely complete ServiceNow block.
                $json.ServiceNow.CredentialsFilePath = 'TestDrive:\snow.cred'
                $errors = Invoke-Validation -Json $json
                $names = Get-ErrorNames $errors

                $names | Should -Not -Contain 'Incorrect configuration'
                $names | Should -Not -Contain "Missing 'ServiceNow.CredentialsFilePath'"
                $names | Should -Not -Contain "Missing 'ServiceNow.TableName'"
                $names | Should -Not -Contain "Missing 'ServiceNow.Environment'"
            }

            It 'flags missing ServiceNow.<_> when that property is blank' -ForEach @(
                'CredentialsFilePath', 'TableName', 'Environment'
            ) {
                $json = Set-ValidPaths (New-JsonFixtureWithModifiedValue -Path 'Export.ServiceNowFormDataExcelFile' -Value 'forms.xlsx')
                # Start from a complete block, then blank the one under test.
                $json.ServiceNow.CredentialsFilePath = 'TestDrive:\snow.cred'
                $json.ServiceNow.$_ = ''
                $errors = Invoke-Validation -Json $json

                Get-ErrorNames $errors | Should -Contain "Missing 'ServiceNow.$_'"
            }
        }
    }
}