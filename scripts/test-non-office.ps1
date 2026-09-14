#requires -Version 7
<#
.SYNOPSIS
    Runs the non-Office (OfficeIntegration-excluded) test suite (W8 local/CI parity entry point).

.DESCRIPTION
    Single source of truth for the `dotnet test` invocation shared by
    verify-quick.ps1 step 11 and the GitHub CI
    'Test (OfficeIntegration excluded)' step. Inline `dotnet test` in
    ci.yml is forbidden by scripts/ci-parity.Tests.ps1 (W8), for the
    same reason inline Pester is: a test command inlined in the workflow
    can drift from the version-pinned script gate (the Pester 4->5
    `-Script` removal was caught the same way).

    Exit 0 on clean; the dotnet test exit code on failure.
#>

[CmdletBinding()]
param(
    [string]$Solution      = 'GanttCreator.slnx',
    [string]$Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

# dotnet test accepts exactly ONE project or solution argument. Passing
# multiple project paths is rejected by MSBuild (MSB1008: Only one project
# can be specified) -- the b54ccbd regression. Test the whole solution and
# exclude the Office-hosted suite by its xUnit Category trait; every test
# in Office.IntegrationTests is tagged [Trait("Category","OfficeIntegration")]
# (tests/Directory.Build.props). W14 tripwire (scripts/ci-parity.Tests.ps1)
# asserts every scripts/*.ps1 dotnet test line passes exactly one target.
dotnet test $Solution -c $Configuration --no-build --no-restore `
    --filter 'Category!=OfficeIntegration'
exit $LASTEXITCODE
