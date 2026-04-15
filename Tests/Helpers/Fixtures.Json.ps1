# factory for building JSON input

# JSON Fixture Factory Functions
# Safe for Pester BeforeDiscovery
# No top-level code, no TestDrive calls, no external variable references.

function New-JsonFixture {
    <#
        .SYNOPSIS
            Creates a complete, valid JSON input object for the script.
        .DESCRIPTION
            This object contains all required top-level and nested properties.
            All filesystem paths must be assigned in the test file (BeforeAll).
        .NOTES
            This fixture returns a *hashtable*, not JSON. 
            Tests should use Save-TestJson to write it out.
    #>

    return @{
        MaxConcurrent          = @{
            Computers             = 1
            JobsPerRemoteComputer = 1
            FoldersPerMatrix      = 1
        }
        Matrix                 = @{
            FolderPath             = ''      # Filled in tests
            DefaultsFile           = ''      # Filled in tests
            Archive                = $false
            ExcludedSamAccountName = @()
        }
        Export                 = @{
            ServiceNowFormDataExcelFile = $null
            OverviewHtmlFile            = $null
            PermissionsExcelFile        = $null
        }
        ServiceNow             = @{
            CredentialsFilePath = ''
            Environment         = 'Test'
            TableName           = 'roles'
        }
        PSSessionConfiguration = 'PowerShell.7'
        Settings               = @{
            ScriptName     = 'Test (Brecht)'

            Advanced       = @{
                DriveMapMaxRetries   = 1
                DriveMapSleepSeconds = 0
            }

            SendMail       = @{
                From         = 'm@example.com'
                To           = '007@example.com'
                Subject      = 'Email subject'
                Body         = 'Email body'
                Smtp         = @{
                    ServerName     = 'SMTP_SERVER'
                    Port           = 25
                    ConnectionType = 'StartTls'
                    UserName       = 'bob'
                    Password       = 'pass'
                }
                AssemblyPath = @{
                    MailKit = 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll'
                    MimeKit = 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll'
                }
            }

            SaveLogFiles   = @{
                Detailed            = $true
                Where               = @{
                    Folder = ''   # Filled in tests
                }
                DeleteLogsAfterDays = 30
            }

            SaveInEventLog = @{
                Save    = $true
                LogName = 'Scripts'
            }
        }
    }
}

function New-JsonFixtureWithMissingProperty {
    <#
        .SYNOPSIS
            Returns a JSON fixture with one top-level property removed.
        .PARAMETER Property
            The name of the top-level property to remove.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Property
    )

    $json = New-JsonFixture
    $json.Remove($Property) | Out-Null
    return $json
}

function New-JsonFixtureWithModifiedValue {
    <#
        .SYNOPSIS
            Returns a JSON fixture with a nested property set to a specific value.
        .PARAMETER Path
            Dot-notation path to the property (e.g. "Matrix.FolderPath")
        .PARAMETER Value
            The new value to assign
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        $Value
    )

    $json = New-JsonFixture

    # Resolve nested path
    $segments = $Path.Split('.')
    $target = $json

    for ($i = 0; $i -lt $segments.Count - 1; $i++) {
        $target = $target[$segments[$i]]
    }

    $target[$segments[-1]] = $Value
    return $json
}

function New-JsonFixtureWithInvalidBoolean {
    <#
        .SYNOPSIS
            Returns a JSON fixture with a boolean replaced by invalid value.
        .PARAMETER Path
            Path to the boolean property (e.g. "Matrix.Archive")
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    return New-JsonFixtureWithModifiedValue -Path $Path -Value 'notABoolean'
}

function New-JsonFixtureWithInvalidInteger {
    <#
        .SYNOPSIS
            Replace integer field with invalid non-integer.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    return New-JsonFixtureWithModifiedValue -Path $Path -Value 'abc'
}

function New-JsonFixtureWithInvalidArray {
    <#
        .SYNOPSIS
            Replace array field with invalid non-array.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    return New-JsonFixtureWithModifiedValue -Path $Path -Value 'notAnArray'
}

function New-ValidDefaultsExcelFixture {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    # Mandatory column headers for the Settings sheet
    $rows = @(
        [PSCustomObject]@{
            MailTo       = 'test@example.com'
            ADObjectName = 'TestGroup'
            Permission   = 'R'
        }
    )

    # Create Excel file with sheet "Settings"
    $rows | Export-Excel `
        -Path $Path `
        -WorksheetName 'Settings' `
        -TableName 'Settings' `
        -AutoSize `
        -FreezeTopRow `
        -ErrorAction Stop

    return $Path
}