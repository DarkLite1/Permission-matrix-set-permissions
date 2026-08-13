#Requires -Version 7
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.0.0' }

<#
    The scripts in Scripts\Operations are deliberately self-contained: they are
    invoked by path from configuration ($Context.ScriptPath.*), so they cannot
    assume the PermissionMatrix module is importable, and they resolve 'ENV:'
    secrets before any module could be loaded anyway.

    SetPermissions.ps1 duplicates helpers a second time for the same reason at a
    smaller scale: its worker scriptblock is stringified and rehydrated in a
    remote runspace, where the definitions in the begin block are out of scope.

    The cost of both choices is copies that have to be changed together. This
    file makes that cost visible:
        1. a helper copied into several files must stay identical
        2. a helper copied twice inside one file must stay identical
        3. a name shared by two different implementations must be a recorded
           decision rather than an accident
#>

BeforeAll {
    <#
        Walk up from this file until the repository root is found, so the test
        keeps working wherever it is filed inside Tests\.
    #>
    $script:repoRoot = $null
    $candidate = [System.IO.DirectoryInfo]::new($PSScriptRoot)

    while ($candidate) {
        $hasModules = Test-Path -LiteralPath (Join-Path $candidate.FullName 'Modules\PermissionMatrix') -PathType Container
        $hasScripts = Test-Path -LiteralPath (Join-Path $candidate.FullName 'Scripts\Operations') -PathType Container

        if ($hasModules -and $hasScripts) {
            $script:repoRoot = $candidate.FullName
            break
        }

        $candidate = $candidate.Parent
    }

    if (-not $script:repoRoot) {
        throw "Could not locate the repository root above '$PSScriptRoot'. Expected a folder holding both 'Modules\PermissionMatrix' and 'Scripts\Operations'."
    }

    $script:operationsFolder = Join-Path $script:repoRoot 'Scripts\Operations'

    <#
        Reduce every definition of a function to a comparable signature: the
        token stream of the definition with comments and line breaks removed.

        Tokenizing rather than comparing raw text means indentation, blank lines
        and help blocks can differ freely while any change to the actual code is
        caught. Returns one entry per definition, so a function defined twice in
        the same file yields two.
    #>
    function Get-FunctionTokenSignature {
        param(
            [Parameter(Mandatory)][string]$Path,
            [Parameter(Mandatory)][string]$FunctionName
        )

        $tokens = $null
        $parseErrors = $null

        $ast = [System.Management.Automation.Language.Parser]::ParseFile(
            $Path, [ref]$tokens, [ref]$parseErrors
        )

        if ($parseErrors) {
            throw "Failed to parse '$Path': $($parseErrors[0].Message)"
        }

        $functions = $ast.FindAll(
            {
                param($node)
                $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                $node.Name -eq $FunctionName
            }, $true
        )

        foreach ($function in $functions) {
            $start = $function.Extent.StartOffset
            $end = $function.Extent.EndOffset

            $relevant = $tokens | Where-Object {
                $_.Extent.StartOffset -ge $start -and
                $_.Extent.EndOffset -le $end -and
                $_.Kind -ne 'Comment' -and
                $_.Kind -ne 'NewLine' -and
                $_.Kind -ne 'EndOfInput'
            }

            [PSCustomObject]@{
                Line      = $function.Extent.StartLineNumber
                Signature = ($relevant.Text -join ' ')
            }
        }
    }

    function Get-DefinedFunctionName {
        param([Parameter(Mandatory)][string]$Path)

        $parseErrors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile(
            $Path, [ref]$null, [ref]$parseErrors
        )

        if ($parseErrors) {
            throw "Failed to parse '$Path': $($parseErrors[0].Message)"
        }

        return $ast.FindAll(
            {
                param($node)
                $node -is [System.Management.Automation.Language.FunctionDefinitionAst]
            }, $true
        ).Name
    }

    $script:operationsScript = @(
        Get-ChildItem -Path $script:operationsFolder -Filter '*.ps1' -File
    )

    if ($script:operationsScript.Count -eq 0) {
        throw "No scripts found in '$script:operationsFolder'."
    }

    <#
        Names deliberately shared by two DIFFERENT implementations. Listing one
        here is a statement that the difference is intentional and that the two
        should not be merged; renaming one of them is the better fix.

        Invoke-WithRetryHC
            UpdateServiceNow   - returns $true/$false, has a -Critical switch,
                                 retries every failure at a fixed 3s interval
            UploadToSharePoint - returns the action's own result, throws on
                                 non-retryable Graph errors, honours Retry-After
                                 and backs off exponentially

        Test-AclEqualHC (twice inside SetPermissions.ps1)
            begin block  - takes the reference ACEs as an array and builds a
                           fingerprint HashSet on every call
            scriptblock  - takes the HashSet already built by the caller and
                           early-exits on the first unmatched ACE, so the set is
                           built once per folder tree instead of once per folder
            Same comparison, same result; the second is the hot-path variant.
    #>
    $script:knownDivergent = @(
        'Invoke-WithRetryHC'
        'Test-AclEqualHC'
    )

    # Names duplicated ACROSS Operations scripts that must stay identical
    $script:knownShared = @('Get-StringValueHC')

    <#
        Names duplicated INSIDE one Operations script that must stay identical.
        Kept in step with the 'copied twice inside one file' test below.
    #>
    $script:knownInFileShared = @(
        'ConvertTo-HashtableHC'
        'ConvertTo-MatrixAdObjectHC'
        'New-UnreadableAclEntryHC'
        'Get-DirectoryAclSafeHC'
    )
}

Describe 'A helper copied into several files stays identical' {
    <#
        Each entry lists every file that must carry an identical copy. When a
        helper is copied into another file, add that path here so it is covered.
    #>
    It '<Name> is identical in every file that defines it' -ForEach @(
        @{
            Name  = 'Get-StringValueHC'
            Files = @(
                'Modules\PermissionMatrix\Private\Utils.ps1'
                'Scripts\Operations\UpdateServiceNow.ps1'
                'Scripts\Operations\UploadToSharePoint.ps1'
            )
        }
    ) {
        $signatures = foreach ($relativePath in $Files) {
            $fullPath = Join-Path $script:repoRoot $relativePath

            if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
                throw "Expected '$relativePath' to exist and define $Name."
            }

            $found = @(
                Get-FunctionTokenSignature -Path $fullPath -FunctionName $Name
            )

            if ($found.Count -eq 0) {
                throw "'$relativePath' no longer defines $Name. Update the expected file list in this test."
            }

            [PSCustomObject]@{
                File      = $relativePath
                Signature = $found[0].Signature
            }
        }

        $distinct = @($signatures.Signature | Select-Object -Unique)

        if ($distinct.Count -ne 1) {
            $detail = ($signatures | ForEach-Object { "  $($_.File)" }) -join "`n"
            throw "$Name has drifted between copies. Every copy must behave identically:`n$detail"
        }

        $distinct.Count | Should-Be 1
    }
}

Describe 'A helper copied twice inside one file stays identical' {
    <#
        SetPermissions.ps1 defines these once in its begin block and again
        inside the worker scriptblock, which runs in a remote runspace that
        cannot see the begin block. Both copies must behave the same.
    #>
    It '<Name> is identical in both copies inside <Script>' -ForEach @(
        @{ Script = 'SetPermissions.ps1'; Name = 'ConvertTo-HashtableHC' }
        @{ Script = 'SetPermissions.ps1'; Name = 'ConvertTo-MatrixAdObjectHC' }
        @{ Script = 'SetPermissions.ps1'; Name = 'New-UnreadableAclEntryHC' }
        @{ Script = 'SetPermissions.ps1'; Name = 'Get-DirectoryAclSafeHC' }
    ) {
        $fullPath = Join-Path $script:operationsFolder $Script

        $found = @(
            Get-FunctionTokenSignature -Path $fullPath -FunctionName $Name
        )

        if ($found.Count -lt 2) {
            throw "$Name is defined $($found.Count) time(s) in $Script. It is no longer duplicated, so remove it from this test."
        }

        $distinct = @($found.Signature | Select-Object -Unique)

        if ($distinct.Count -ne 1) {
            $lines = ($found | ForEach-Object { "  line $($_.Line)" }) -join "`n"
            throw "$Name has drifted between its copies in $Script. Both must behave identically:`n$lines"
        }

        $distinct.Count | Should-Be 1
    }
}

Describe 'A function name used twice is a recorded decision' {
    <#
        Same name, different behaviour is harder to spot than plain drift: a
        reader who has seen one copy will assume the other matches. Anything not
        already recorded above has to be looked at.
    #>
    It 'introduces no unreviewed duplicate name across the Operations scripts' {
        $byName = @{}

        foreach ($file in $script:operationsScript) {
            $names = @(Get-DefinedFunctionName -Path $file.FullName) |
            Select-Object -Unique

            foreach ($name in $names) {
                if (-not $byName.ContainsKey($name)) {
                    $byName[$name] = [System.Collections.Generic.List[string]]::new()
                }
                $byName[$name].Add($file.Name)
            }
        }

        $unreviewed = $byName.GetEnumerator() |
        Where-Object {
            $_.Value.Count -gt 1 -and
            $_.Key -notin $script:knownShared -and
            $_.Key -notin $script:knownDivergent
        }

        if ($unreviewed) {
            $detail = (
                $unreviewed | ForEach-Object {
                    "  $($_.Key): $($_.Value -join ', ')"
                }
            ) -join "`n"

            throw @"
These names are defined in more than one Operations script but are not recorded
in this test:
$detail
Either keep the copies identical and add the name to `$knownShared, or give them
distinct names. If they are deliberately different, add the name to
`$knownDivergent with a note explaining why.
"@
        }

        $true | Should-BeTrue
    }

    It 'introduces no unreviewed duplicate name inside a single Operations script' {
        $duplicate = foreach ($file in $script:operationsScript) {
            $names = @(Get-DefinedFunctionName -Path $file.FullName)

            $repeated = $names |
            Group-Object |
            Where-Object { $_.Count -gt 1 } |
            Select-Object -ExpandProperty Name

            foreach ($name in $repeated) {
                if (
                    $name -in $script:knownShared -or
                    $name -in $script:knownDivergent -or
                    $name -in $script:knownInFileShared
                ) { continue }

                "  $($file.Name): $name"
            }
        }

        if ($duplicate) {
            throw @"
These names are defined more than once inside a single Operations script but are
not recorded in this test:
$($duplicate -join "`n")
Add the name to the 'copied twice inside one file' test when both copies must
match, or to `$knownDivergent when they are deliberately different.
"@
        }

        $true | Should-BeTrue
    }

    It '<_> is still defined more than once' -ForEach @(
        'Get-StringValueHC'
        'Invoke-WithRetryHC'
        'Test-AclEqualHC'
    ) {
        <#
            Guards the lists above against going stale: once a helper exists in
            only one place its entry should be removed from this test.
        #>
        $functionName = $_
        $definitionCount = 0

        foreach ($file in $script:operationsScript) {
            $definitionCount += @(
                Get-DefinedFunctionName -Path $file.FullName |
                Where-Object { $_ -eq $functionName }
            ).Count
        }

        $definitionCount | Should-BeGreaterThan 1
    }
}