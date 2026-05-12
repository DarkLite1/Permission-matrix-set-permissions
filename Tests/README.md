# Tests

Pester 5 test suite for the Permission Matrix module. PowerShell 7+ required.

## Folder layout

```
Tests/
├── Helpers/         Shared fixtures and utilities (dot-sourced by tests)
├── Unit/            One test file per source file; mocks everything external
├── Integration/     Multi-component scenarios; shells out to entrypoint scripts
└── Scripts/         Tests for the Operations\ scripts
```

### Unit/

Mirrors the production source tree. Source file at
`Modules/PermissionMatrix/Private/<path>/Foo.ps1` has its tests at
`Tests/Unit/Private/<path>/Foo.Tests.ps1`. Walking from a source file to its
tests is mechanical — no thinking required.

```
Unit/
├── Private/
│   ├── ActiveDirectory.Tests.ps1
│   ├── Core.Tests.ps1
│   ├── Export.Tests.ps1
│   ├── Html.Tests.ps1
│   ├── Mail.Tests.ps1
│   ├── Matrix.Tests.ps1
│   ├── Utils.Tests.ps1
│   ├── Validation.Tests.ps1
│   ├── Core/
│   │   ├── Invoke-PermissionMatrixBeginHC.Tests.ps1
│   │   ├── Invoke-PermissionMatrixEndHC.Tests.ps1
│   │   └── Invoke-PermissionMatrixProcessHC.Tests.ps1
│   ├── Logging/
│   ├── Matrix/
│   └── Validation/
└── Public/
    └── Invoke-PermissionMatrix.Tests.ps1
```

Unit tests dot-source the private files they exercise (and any direct
dependencies), then call the function under test directly. Everything external
— AD, filesystem outside `TestDrive:`, sub-helpers — is mocked.

### Integration/

Cross-cutting tests that exercise more than one component, typically by
shelling out to an entrypoint script with a real (TestDrive) config. These run
slower and serve as a safety net for end-to-end flows.

### Scripts/

Tests for the standalone scripts under `Operations\`: `SetPermissions.ps1`,
`TestRequirements.ps1`, `UpdateServiceNow.ps1`.

### Helpers/

Shared utilities every test can dot-source:

- `Helpers.HC.ps1` — `Copy-ObjectHC`, `Save-TestJson`, `Set-NestedPropertyHC`,
  log-folder helpers, comparison utilities
- `Fixtures.Json.ps1` — `New-JsonFixture` (valid base config) and
  `New-...InvalidFixture` variants, `New-ValidDefaultsExcelFixture`
- `Fixtures.Excel.ps1` — Excel file builders for matrix and defaults files
- `Fixtures.Matrix.ps1` — parametrized `-TestCases` data
- `Fixtures.TestCases.ps1` — bad-input test case tables

## Running

```powershell
# Full suite
Invoke-Pester -Path .\Tests

# Fast feedback (unit only)
Invoke-Pester -Path .\Tests\Unit

# One file, verbose
Invoke-Pester -Path .\Tests\Unit\Private\Core\Invoke-PermissionMatrixBeginHC.Tests.ps1 -Output Detailed

# Single test by tag (add -Tag test to the It block you're iterating on)
Invoke-Pester -Path .\Tests -Tag test
```

## File naming

- **Test files end with `.Tests.ps1`** (Pester's default discovery pattern).
  `Foo_Tests.ps1` won't be picked up automatically.
- **No spaces in filenames** — they confuse some shells and CI tools. The
  source script `Set permissions.ps1` is tested by `SetPermissions.Tests.ps1`.

## Writing a new unit test

A typical unit test file looks like:

```powershell
#Requires -Version 7
#Requires -Modules Pester

Describe 'Function-Under-Test' {
    BeforeAll {
        $root = Resolve-Path "$PSScriptRoot\..\..\..\.."
        $moduleRoot = "$root\Modules\PermissionMatrix"

        . "$root\Tests\Helpers\Helpers.HC.ps1"
        . "$moduleRoot\Private\Utils.ps1"
        # ...other dot-sources as needed
        . "$moduleRoot\Private\Path\To\Function-Under-Test.ps1"

        # Local fixture builders, if useful
        function New-FakeThing { ... }
    }

    BeforeEach {
        # Per-test state and default-safe mocks
        $systemErrors = [System.Collections.Generic.List[object]]::new()
        Mock Some-Helper { }
    }

    Context 'a specific behavior' {
        It 'does what it should' {
            # arrange
            # act
            # assert
        }
    }
}
```

### Conventions worth following

**One Context per behavior, one It per concrete claim.** A 25-line `It`
that's clear in isolation beats a 5-line `It` that hides setup in helpers.
Tests aren't production code — readability at the point of failure matters
more than DRY.

**Context-level `BeforeEach` for shared preconditions.** If every test in a
Context needs the same fixture (a matrix file, a particular mock), put it in
that Context's `BeforeEach`, not in each `It`.

**Outer `BeforeEach` for default-safe mocks.** Functions the system-under-test
calls should be mocked to no-op by default in the top-level `BeforeEach`. Each
`It` overrides specific mocks as needed.

**Local fixture builders > shared helpers.** Each test file's `BeforeAll`
should define its own `New-FakeX` builders that produce the exact shape the
function under test consumes. Cross-file fixture sharing tends to drift from
production shapes over time.

## Mocking patterns and traps

### Mock names must match the actual function name

```powershell
# In source: function Import-MatrixDefaultsFileHC { ... }
# In test:
Mock Import-MatrixDefaultsFileHC { ... }    # correct
Mock Import-MatrixDefaultsFile     { ... }  # silently no-ops, doesn't throw
```

A `Mock <name>` for a function that doesn't exist in the current scope **does
not always throw** — sometimes it silently does nothing. If a test mysteriously
isn't intercepting, check spelling against the function definition, not the
filename. PowerShell file-name conventions are conventional, not enforced.

### Pester respects the real function's parameter binding

Mocks honor `[Parameter(Mandatory)]` and type constraints from the original
function signature. The mock body never runs if the binder rejects the
arguments first. Common case:

```powershell
function Real-Function {
    param([Parameter(Mandatory)][hashtable[]]$Matrix)
    # ...
}

Mock Real-Function { return @() }
Real-Function -Matrix @()    # ← throws BEFORE the mock body runs
```

If your fake objects fail to bind to a real function's parameters, the test
fails in confusing ways. Either give the fixture builder defaults that satisfy
the binder, or refactor the production code's signature to be looser.

### Mocks don't cross runspace boundaries

`Mock Inner-Helper` won't intercept calls made inside an
`Invoke-WithOptionalParallelismHC` scriptblock that runs in a separate
runspace. For tests that need to control what comes out of parallel work, mock
`Invoke-WithOptionalParallelismHC` itself and return canned results.

## Diagnosing failing tests

When a test fails in a way that doesn't match your hypothesis after one round
of reading the source, instrument before guessing again. A useful template:

```powershell
It 'failing test' {
    # ... setup ...

    Write-Host "===== DIAG ====="
    Write-Host "Function exists: $(Get-Command Function-Under-Test -EA SilentlyContinue)"
    # Direct call to verify mocks intercept as expected:
    Write-Host "Direct mock call: $(Mocked-Function -SomeArg 'value')"

    # ... call system under test ...

    Write-Host "Result state: $($result | Out-String)"
    Write-Host "SystemErrors count: $($systemErrors.Count)"
    foreach ($e in $systemErrors) {
        Write-Host "  [$($e.Type)] $($e.Name): $($e.Message)"
    }
    Write-Host "================"

    # ... assertions ...
}
```

This template answers in one run:

- Was the function/mock even discoverable?
- Did the mock fire?
- What errors got logged?
- What's the actual resulting state?

It's faster than a third round of "I think the bug is..."

## Skill issues we've hit before

A list of recurring footguns, written down so the next person doesn't
re-discover them:

- **`$args` is a PowerShell automatic variable.** Using `$args = @{...}; & $foo @args`
  inside an `It` block may or may not work depending on scope. Use `$splat`,
  `$callArgs`, or anything else.
- **`[string]` casts coerce `$null` to `''`.** A `[string]$Name` parameter
  defaulted with `??` won't fall back when `$null` is passed — `$null` became
  `''`. Use `[string]::IsNullOrWhiteSpace()` instead of `-not $value` or `??`
  for blank-string fallbacks. There's a `Get-StringOrDefaultHC` helper in
  `Utils.ps1` that handles this correctly.
- **`-not '0'` is `$true` in PowerShell.** `-not <string>` treats the literal
  string `'0'` as falsy. Affects any string-presence guard using `-not`.
- **Empty arrays and Mandatory parameters don't mix.** `Pester` and PowerShell
  both reject `@()` for `[Parameter(Mandatory)]` (treated as "argument not
  supplied"). Fixture builders should produce non-empty defaults.

## Adding tests for a new source file

1. Locate the source file: `Modules/PermissionMatrix/Private/<Area>/Foo.ps1`
2. Create `Tests/Unit/Private/<Area>/Foo.Tests.ps1`
3. Use the skeleton above. Dot-source `Foo.ps1` and its direct dependencies.
4. Group `It`s into `Context`s by behavior, not by code branch.
5. Mock everything `Foo.ps1` calls that isn't pure (filesystem, AD, sibling
   helpers).
6. For each Mandatory parameter the function under test takes, the test must
   supply a value that satisfies its type and `[ValidateX]` attributes — even
   if the test doesn't care about that value.

If the test ends up shelling out to a real script, dot-sourcing many files, or
asserting on log file contents, it belongs in `Tests/Integration/`, not
`Tests/Unit/`.
