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
# exclude the host-dependent suites by their xUnit Category traits. The
# traits are per-test [Trait] attributes (the OfficeIntegration convention
# is documented in tests/Directory.Build.props): OfficeIntegration bodies
# need a live Office host, TransferSoak bodies (R7.7) are the 25-transfer
# soak that runs only through its own measured gate (D4, ADR-0001), and
# FontProbe bodies (R8.2a) need the Office-delivered Aptos cloud fonts that a
# GitHub-hosted runner does not have, so they run only through the
# reference-machine evidence command in ADR-0070. The filter is NEGATED, so
# every exclusion is load-bearing: a TransferSoak-only or FontProbe-only test
# matches Category!=OfficeIntegration and would otherwise run on every commit,
# in verify.ps1 and in CI. W14 tripwire (scripts/ci-parity.Tests.ps1) asserts
# every scripts/*.ps1 dotnet test line passes exactly one target; the same
# file pins all three exclusions.
dotnet test $Solution -c $Configuration --no-build --no-restore `
    --filter 'Category!=OfficeIntegration&Category!=TransferSoak&Category!=FontProbe'
exit $LASTEXITCODE
