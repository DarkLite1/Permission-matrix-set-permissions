#Requires -Module @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

<#
    Tests for Validate-ConfigurationStructureHC.

    How input is built
    -------------------
    Fixtures.Json.ps1 returns the config as a *hashtable*. Production feeds this
    function a PSCustomObject parsed by ConvertFrom-Json, so each test round-trips
    the fixture through ConvertTo-Json | ConvertFrom-Json. That also gives the
    correct "missing property reads as $null" semantics the function relies on.

    How results are asserted
    ------------------------
    The function does not return anything; it appends errors to the [ref]
    $SystemErrors via Add-JsonSchemaErrorHC (a thin wrapper over Add-ErrorHC in
    Utils.ps1, fixing Category='JsonSchema'). Tests inspect the collected list.

    >> CONTRACT ASSUMPTION (verify against Utils.ps1): each recorded error exposes
       a .Name property matching the -Name passed in, and a .Type property. If
       Add-ErrorHC stores under different property names, adjust Get-ErrorNames
       and the .Type assertions in one place below.

    Filesystem
    ----------
    Matrix.FolderPath / Matrix.DefaultsFile / Settings.SaveLogFiles.Where.Folder
    are validated with Test-Path. The happy-path fixture leaves them blank, so
    BeforeAll fills them with real TestDrive paths. "Missing"/"incorrect" cases
    leave them blank or point them at non-existent paths.

    NOTE: Adjust dot-source paths if this file is relocated. Assumes
    Tests\Unit\Private\ with helpers under Tests\Helpers\.
#>

BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\PermissionMatrix"

    . "$moduleRoot\Private\Utils.ps1"
    . "$moduleRoot\Private\Validate-ConfigurationStructureHC.ps1"
    . "$root\Tests\Helpers\Fixtures.Json.ps1"

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
        Validate-ConfigurationStructureHC -Json $obj -SystemErrors ([ref]$errors)
        return $errors
    }

    # Single place that knows how an error exposes its name. Adjust here if
    # Utils.ps1 stores under a different property.
    function Get-ErrorNames {
        param($Errors)
        @($Errors | ForEach-Object { $_.Name })
    }
}

Describe 'Validate-ConfigurationStructureHC' {

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

    Context 'KNOWN BUG: Export reads $ServiceNow (local) not $Json.ServiceNow' {
        # The Export region checks the bare variable $ServiceNow, which is never
        # assigned in function scope, so it is always $null. As a result, setting
        # a valid ServiceNowFormDataExcelFile ALWAYS emits 'Incorrect configuration'
        # even though $Json.ServiceNow is fully populated, and the per-property
        # ServiceNow checks (else branch) are unreachable.
        # These tests PIN that current behavior. If the function is fixed to read
        # $Json.ServiceNow, both expectations below must be inverted.
        It 'emits "Incorrect configuration" despite a valid ServiceNow block' {
            $json = Set-ValidPaths (New-JsonFixtureWithModifiedValue -Path 'Export.ServiceNowFormDataExcelFile' -Value 'forms.xlsx')
            $errors = Invoke-Validation -Json $json

            Get-ErrorNames $errors | Should -Contain 'Incorrect configuration'
        }

        It 'never reaches the per-property ServiceNow checks (else branch is dead)' {
            $json = Set-ValidPaths (New-JsonFixtureWithModifiedValue -Path 'Export.ServiceNowFormDataExcelFile' -Value 'forms.xlsx')
            $errors = Invoke-Validation -Json $json
            $names = Get-ErrorNames $errors

            $names | Should -Not -Contain "Missing 'ServiceNow.CredentialsFilePath'"
            $names | Should -Not -Contain "Missing 'ServiceNow.TableName'"
            $names | Should -Not -Contain "Missing 'ServiceNow.Environment'"
        }
    }
}