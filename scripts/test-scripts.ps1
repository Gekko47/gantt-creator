#requires -Version 7
<#
.SYNOPSIS
    Script unit test gate (L7). Runs Pester tests for all scripts/*.ps1 files.
    Wired into verify-quick.ps1 and CI.

.DESCRIPTION
    Discovers Pester test files matching scripts/*Tests.ps1 or scripts/*.Tests.ps1
    and runs them with -PassThru. Aggregates results and exits with the
    appropriate code.

    Tests should be placed alongside the script they test, e.g.:
      scripts/check-status.ps1
      scripts/check-status.Tests.ps1

    Files run in parallel runspaces by default (W15). Measured 2026-11-12: the
    15 files cost 141.6 s sequentially and the per-test cost is dominated by
    child processes and fixtures, not by Pester itself, so the wall time becomes
    the slowest file rather than the sum of all of them. Each file's own output
    is buffered inside its runspace and printed after the run in discovery
    order, so the report stays deterministic even though completion order is
    not. -Serial restores the one-file-at-a-time path with Pester's Detailed
    output, for diagnosing an order-dependent or CPU-contended failure.

    The no-silent-pass policy is unchanged in both modes: an empty discovery is
    a failure, a discovered file that contains zero tests is a failure, and any
    failed test fails the gate. A runspace that throws, or returns no result, is
    a failed file, never a pass; a run that returns fewer results than it
    discovered files is also a failure, so a dropped worker cannot shrink the
    suite silently.

    Exit 0 on all tests pass; exit 1 on any failure.
#>

[CmdletBinding()]
param(
    [string]$ScriptsRoot = 'scripts',
    # Run one file at a time with Pester's Detailed output. Slower; for
    # diagnosing an order-dependent or resource-contended failure.
    [switch]$Serial,
    # 0 = the measured default (4). Measured 2026-11-12 on 8 logical
    # processors: 15 files took 129.8 s at 8 workers, 92.8 s at 4, 94.7 s at 2,
    # and ~142 s sequentially. The suite is process-spawn bound, so oversubscribing
    # the machine costs more than it saves.
    [int]$ThrottleLimit = 0
)

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $PSCommandPath
. (Join-Path $scriptRoot 'verify-helpers.ps1')

$repoRoot = Split-Path -Parent $PSScriptRoot
# Join-Path concatenates rather than resolving a rooted second argument, so an
# absolute -ScriptsRoot (used by this script's own fixture tests) would become
# "<repoRoot>\C:\..." and report the directory as missing.
$scriptsDir = if ([System.IO.Path]::IsPathRooted($ScriptsRoot)) { $ScriptsRoot } else { Join-Path $repoRoot $ScriptsRoot }

if (-not (Test-Path -LiteralPath $scriptsDir)) {
    Write-Error "Scripts directory not found: $scriptsDir"
    exit 1
}

# Ensure Pester 6+ is available. The version is single-sourced from
# scripts/tool-versions.psd1 (W11). Get-Module does not support
# -MinimumVersion, so we check the highest installed version against
# the pinned minimum.
$versionsPath = Join-Path $repoRoot 'scripts\tool-versions.psd1'
$script:toolVersions = Import-PowerShellDataFile -LiteralPath $versionsPath
$script:minPesterMajor = $script:toolVersions.Pester.MinimumMajor

$installedPester = Get-Module -ListAvailable -Name Pester -ErrorAction SilentlyContinue |
    Sort-Object Version -Descending |
    Select-Object -First 1
if (-not $installedPester -or $installedPester.Version.Major -lt $script:minPesterMajor) {
    Write-Host "Installing Pester $($script:minPesterMajor).0+..."
    try {
        Install-Module -Name Pester -MinimumVersion "$($script:minPesterMajor).0" -Scope CurrentUser -Force -SkipPublisherCheck
    }
    catch {
        Write-Error "Failed to install Pester: $_"
        exit 1
    }
}

# Import Pester
try {
    Import-Module Pester -MinimumVersion "$($script:minPesterMajor).0" -ErrorAction Stop
}
catch {
    Write-Error "Pester $($script:minPesterMajor).0+ not available: $_"
    exit 1
}

# Find test files in scripts/. A single '*Tests.ps1' wildcard covers both
# documented name forms ('*.Tests.ps1' and '*Tests.ps1') in one call; the
# previous two-call fallback meant the second pattern was only reached when
# the first found nothing, contrary to the documented discovery contract.
$testFiles = @(Get-ChildItem -Path $scriptsDir -Filter '*Tests.ps1' -ErrorAction SilentlyContinue | Sort-Object Name)

if ($testFiles.Count -eq 0) {
    # No-silent-pass policy: a gate that finds nothing to check proves
    # nothing about scripts/ health, so an empty discovery is a failure
    # with an instruction, never a PASS.
    Write-Error ("test-scripts: No *.Tests.ps1 files found in '$scriptsDir'. " +
        'The script-lint gate requires at least one Pester test file.')
    exit 1
}

Write-Host "Found $($testFiles.Count) script test file(s):"
$testFiles | ForEach-Object { Write-Host "  $($_.Name)" }

# -Parallel passes $using: values into another runspace, so the pinned Pester
# major is copied into a local name rather than referenced through $script:.
$minimumPesterMajor = $script:minPesterMajor

# One result object per file, carrying the fields the judge needs and the text
# of each failure, so the judge does not have to know Pester's result shape.
# The builder is inlined in both paths because a scriptblock cannot be shared
# across runspaces.
if ($Serial) {
    Write-Host 'Mode: serial (-Serial)'
    $results = foreach ($testFile in $testFiles) {
        Write-Host "Running $($testFile.Name)..."
        $r = Invoke-Pester -Path $testFile.FullName -PassThru -Output Detailed
        if ($null -eq $r) {
            [pscustomobject]@{ Name = $testFile.Name; Total = 0; Passed = 0; Failed = 0; Failures = @(); Error = 'Pester returned no result object.' }
        }
        else {
            $failures = @($r.Failed | ForEach-Object {
                    $path = if ($_.PSObject.Properties['ExpandedPath']) { $_.ExpandedPath } else { $_.Name }
                    $first = @($_.ErrorRecord) | Select-Object -First 1
                    $message = if ($first) { $first.Exception.Message } else { 'failed' }
                    "$path : $message"
                })
            [pscustomobject]@{ Name = $testFile.Name; Total = $r.TotalCount; Passed = $r.PassedCount; Failed = $r.FailedCount; Failures = $failures; Error = '' }
        }
    }
}
else {
    # Measured optimum is 4 (see the -ThrottleLimit documentation). Never more
    # than the processor count, so a small CI runner does not oversubscribe.
    $defaultThrottle = [Math]::Min(4, [Environment]::ProcessorCount)
    $throttle = if ($ThrottleLimit -gt 0) { $ThrottleLimit } else { $defaultThrottle }
    Write-Host "Mode: parallel ($throttle worker(s))"

    # Every failure inside a runspace is converted into a result object so a
    # crashed worker is judged as a failed file rather than disappearing.
    $results = $testFiles | ForEach-Object -Parallel {
        $file = $_
        $minimum = $using:minimumPesterMajor
        try {
            Import-Module Pester -MinimumVersion "$minimum.0" -ErrorAction Stop
            # 6>$null drops the information stream a test's own Write-Host would
            # otherwise interleave across workers; the verdict and the failure
            # text are returned as data below.
            $r = Invoke-Pester -Path $file.FullName -PassThru -Output None 6>$null
            if ($null -eq $r) {
                [pscustomobject]@{ Name = $file.Name; Total = 0; Passed = 0; Failed = 0; Failures = @(); Error = 'Pester returned no result object.' }
            }
            else {
                $failures = @($r.Failed | ForEach-Object {
                        $path = if ($_.PSObject.Properties['ExpandedPath']) { $_.ExpandedPath } else { $_.Name }
                        $first = @($_.ErrorRecord) | Select-Object -First 1
                        $message = if ($first) { $first.Exception.Message } else { 'failed' }
                        "$path : $message"
                    })
                [pscustomobject]@{ Name = $file.Name; Total = $r.TotalCount; Passed = $r.PassedCount; Failed = $r.FailedCount; Failures = $failures; Error = '' }
            }
        }
        catch {
            [pscustomobject]@{ Name = $file.Name; Total = 0; Passed = 0; Failed = 0; Failures = @(); Error = $_.Exception.Message }
        }
    } -ThrottleLimit $throttle

    $results = @($results | Sort-Object Name)
}

# The verdict rules live in Get-ScriptTestGateVerdict (scripts/verify-helpers.ps1)
# so the tested code is the code this runner executes. This body owns only the
# streams and the exit code.
$verdict = Get-ScriptTestGateVerdict -Results @($results) -DiscoveredCount $testFiles.Count
foreach ($line in $verdict.Lines) { Write-Host $line }
foreach ($message in $verdict.Errors) { Write-Error -Message $message -ErrorAction Continue }

Write-Host "test-scripts: Total=$($verdict.Total), Passed=$($verdict.Passed), Failed=$($verdict.Failed)"

if ($verdict.ExitCode -ne 0) {
    if ($verdict.Failed -gt 0) {
        Write-Error -Message "test-scripts: $($verdict.Failed) test(s) failed out of $($verdict.Total)" -ErrorAction Continue
    }
    exit $verdict.ExitCode
}

Write-Host "test-scripts: PASS ($($verdict.Total) tests)"
exit 0
