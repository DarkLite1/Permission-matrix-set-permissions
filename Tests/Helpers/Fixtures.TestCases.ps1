# Shared TestCase definitions
# Safe for Pester BeforeDiscovery — no commands, no TestDrive, no global references.

function Get-MissingTopLevelProperties {
    <#
        Returns:
            @{ Property = 'MaxConcurrent' }
            @{ Property = 'Matrix' }
            ...
    #>

    return @(
        @{ Property = 'MaxConcurrent' }
        @{ Property = 'Matrix' }
        @{ Property = 'Export' }
        @{ Property = 'ServiceNow' }
        @{ Property = 'PSSessionConfiguration' }
        @{ Property = 'Settings' }
    )
}

function Get-MissingMaxConcurrentProperties {
    return @(
        @{ Property = 'Computers' }
        @{ Property = 'FoldersPerMatrix' }
        @{ Property = 'JobsPerRemoteComputer' }
    )
}

function Get-MissingMatrixProperties {
    return @(
        @{ Property = 'FolderPath' }
        @{ Property = 'DefaultsFile' }
    )
}

function Get-InvalidMatrixPaths {
    # Use consistent hashtable format for param-binding
    return @(
        @{ Property = 'Matrix.FolderPath'  ; Value = 'TestDrove:\NotExisting' }
        @{ Property = 'Matrix.DefaultsFile'; Value = 'TestDrive:\NotExisting.xlsx' }
    )
}

function Get-InvalidBooleanCases {
    <#
        Useful for validating boolean logic:
        - Settings.SaveLogFiles.Detailed
        - Matrix.Archive
        etc.
    #>

    return @(
        @{ Path = 'Settings.SaveLogFiles.Detailed' ; Value = 'abc' }
        @{ Path = 'Matrix.Archive'                 ; Value = 'hello' }
    )
}

function Get-InvalidIntegerCases {
    return @(
        @{ Path = 'MaxConcurrent.Computers'        ; Value = 'hello' }
        @{ Path = 'MaxConcurrent.FoldersPerMatrix' ; Value = 'world' }
        @{ Path = 'MaxConcurrent.JobsPerRemoteComputer'; Value = 'abc123' }
    )
}

function Get-InvalidArrayCases {
    return @(
        @{ Path = 'Matrix.AdGroupPlaceHolders' ; Value = 'not-an-array' }
    )
}