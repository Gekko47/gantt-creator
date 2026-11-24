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

# An unexpected terminating error must still name itself in the check-run
# annotations: without this the step reports only "Process completed with exit
# code 1", which is the gap that made the 2026-11-18 CI failure unreadable.
trap {
    if ($env:GITHUB_ACTIONS -eq 'true') {
        Write-Host (Format-CheckAnnotation -Title 'test-scripts: unhandled error' -Message $_.Exception.Message)
    }
    Write-Error -Message $_.Exception.Message -ErrorAction Continue
    exit 1
}

$repoRoot = Split-Path -Parent $PSScriptRoot
# Join-Path concatenates rather than resolving a rooted second argument, so an
# absolute -ScriptsRoot (used by this script's own fixture tests) would become
# "<repoRoot>\C:\..." and report the directory as missing.
$scriptsDir = if ([System.IO.Path]::IsPathRooted($ScriptsRoot)) { $ScriptsRoot } else { Join-Path $repoRoot $ScriptsRoot }

# Every path this gate hands to a cmdlet is validated first. A null path is not
# a harmless edge case: `Invoke-Pester -Path $null` silently falls back to the
# current directory and still reports a verdict, and PathIntrinsics throws
# "The provided Path argument was null or an empty collection." for a null or
# empty-collection -Path with no attribution (CI 45193295738 / 45376223093).
Assert-UsablePath -Path $repoRoot -Caller 'test-scripts repository root'
Assert-UsablePath -Path $scriptsDir -Caller 'test-scripts scripts directory'

if (-not (Test-Path -LiteralPath $scriptsDir -ErrorAction Stop)) {
    Write-Error "Scripts directory not found: $scriptsDir"
    exit 1
}

# Ensure Pester 6+ is available. The version is single-sourced from
# scripts/tool-versions.psd1 (W11). Get-Module does not support
# -MinimumVersion, so we check the highest installed version against
# the pinned minimum.
$versionsPath = Join-Path $repoRoot 'scripts\tool-versions.psd1'
Assert-UsablePath -Path $versionsPath -Caller 'test-scripts tool-versions'
$script:toolVersions = Import-PowerShellDataFile -LiteralPath $versionsPath -ErrorAction Stop
$script:minPesterMajor = $script:toolVersions.Pester.MinimumMajor
# Exact pin (W11). Absent in a pre-pin psd1, in which case the minimum-major
# behaviour below is preserved rather than failing the gate.
$script:pesterVersion = [string]$script:toolVersions.Pester.Version

$installedPester = Get-Module -ListAvailable -Name Pester -ErrorAction SilentlyContinue |
    Sort-Object Version -Descending |
    Select-Object -First 1
if (-not $installedPester -or $installedPester.Version.Major -lt $script:minPesterMajor) {
    Write-Host "Installing Pester $($script:pesterVersion) (minimum major $($script:minPesterMajor))..."
    try {
        # An open-ended install is what made CI run a different Pester on every
        # hosted runner; install the pinned version when the psd1 carries one.
        $install = @{ Name = 'Pester'; Scope = 'CurrentUser'; Force = $true; SkipPublisherCheck = $true }
        if ($script:pesterVersion) { $install['RequiredVersion'] = $script:pesterVersion }
        else { $install['MinimumVersion'] = "$($script:minPesterMajor).0" }
        Install-Module @install -ErrorAction Stop
    }
    catch {
        if ($env:GITHUB_ACTIONS -eq 'true') {
            Write-Host (Format-CheckAnnotation -Title 'test-scripts: Pester install failed' -Message $_.Exception.Message)
        }
        Write-Error "Failed to install Pester: $_"
        exit 1
    }
}

# Import Pester, and state which build of it this gate actually runs: the
# version and module path are the first things to check when a hosted runner
# behaves differently from a developer machine.
try {
    if ($script:pesterVersion) { Import-Module Pester -RequiredVersion $script:pesterVersion -ErrorAction Stop }
    else { Import-Module Pester -MinimumVersion "$($script:minPesterMajor).0" -ErrorAction Stop }
    $loadedPester = Get-Module -Name Pester | Select-Object -First 1
    Write-Host "Pester $($loadedPester.Version) from $($loadedPester.ModuleBase)"
}
catch {
    if ($env:GITHUB_ACTIONS -eq 'true') {
        Write-Host (Format-CheckAnnotation -Title 'test-scripts: Pester import failed' -Message $_.Exception.Message)
    }
    Write-Error "Pester $($script:minPesterMajor).0+ not available: $_"
    exit 1
}

# Find test files in scripts/. A single '*Tests.ps1' wildcard covers both
# documented name forms ('*.Tests.ps1' and '*Tests.ps1') in one call; the
# previous two-call fallback meant the second pattern was only reached when
# the first found nothing, contrary to the documented discovery contract.
#
# Discovery yields FULL PATH STRINGS, not FileInfo objects. A FileInfo crossing
# the ForEach-Object -Parallel runspace boundary is the input shape that
# produced "The provided Path argument was null or an empty collection." on
# hosted runners; a string cannot deserialize into a wrapper whose FullName is
# null.
$testFiles = @(Get-ChildItem -LiteralPath $scriptsDir -Filter '*Tests.ps1' -File -ErrorAction Stop |
        Sort-Object Name |
        ForEach-Object { $_.FullName })

if ($testFiles.Count -eq 0) {
    # No-silent-pass policy: a gate that finds nothing to check proves
    # nothing about scripts/ health, so an empty discovery is a failure
    # with an instruction, never a PASS.
    $message = "test-scripts: No *.Tests.ps1 files found in '$scriptsDir'. " +
        'The script-lint gate requires at least one Pester test file.'
    if ($env:GITHUB_ACTIONS -eq 'true') {
        Write-Host (Format-CheckAnnotation -Title 'test-scripts: no test files' -Message $message)
    }
    Write-Error $message
    exit 1
}

# The array is non-empty here, so a null or blank entry means discovery produced
# a path no worker can use. Refuse it with a named message rather than let it
# reach Invoke-Pester, which silently falls back to the current directory.
Assert-UsablePath -Path $testFiles -Caller 'test-scripts discovery'

Write-Host "Found $($testFiles.Count) script test file(s):"
$testFiles | ForEach-Object { Write-Host "  $([System.IO.Path]::GetFileName($_))" }

# -Parallel passes $using: values into another runspace, so the pinned Pester
# version is copied into explicit [string] locals rather than referenced through
# $script:. Only strings cross that boundary: a FileInfo deserializes into a
# wrapper whose FullName can be null, which is one of the ways this gate used to
# fail with "The provided Path argument was null or an empty collection."
$minimumPesterVersion = "$($script:minPesterMajor).0"
$pinnedPesterVersion = [string]$script:pesterVersion

# One result object per file, carrying the fields the judge needs and the text
# of each failure, so the judge does not have to know Pester's result shape.
# The builder is inlined in both paths because a scriptblock cannot be shared
# across runspaces.
if ($Serial) {
    Write-Host 'Mode: serial (-Serial)'
    $results = foreach ($testFile in $testFiles) {
        $name = [System.IO.Path]::GetFileName($testFile)
        Write-Host "Running $name..."
        # Guarded so a refused or broken path is a named failed file, exactly
        # as the parallel path reports it.
        try {
            if ([string]::IsNullOrWhiteSpace($testFile)) {
                throw 'refusing a null, empty, or whitespace-only test path'
            }
            $r = Invoke-Pester -Path $testFile -PassThru -Output Detailed
            if ($null -eq $r) {
                [pscustomobject]@{ Name = $name; Total = 0; Passed = 0; Failed = 0; Failures = @(); Error = 'Pester returned no result object.' }
            }
            else {
                $failures = @($r.Failed | ForEach-Object {
                        $path = if ($_.PSObject.Properties['ExpandedPath']) { $_.ExpandedPath } else { $_.Name }
                        $first = @($_.ErrorRecord) | Select-Object -First 1
                        $message = if ($first) { $first.Exception.Message } else { 'failed' }
                        "$path : $message"
                    })
                [pscustomobject]@{ Name = $name; Total = $r.TotalCount; Passed = $r.PassedCount; Failed = $r.FailedCount; Failures = $failures; Error = '' }
            }
        }
        catch {
            [pscustomobject]@{ Name = $name; Total = 0; Passed = 0; Failed = 0; Failures = @(); Error = $_.Exception.Message }
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
        # Explicit [string] locals: the worker owns its own copy of the path, so
        # no FileInfo wrapper (and no null FullName) can cross the runspace.
        $testPath = [string]$_
        $name = [System.IO.Path]::GetFileName($testPath)
        $minimum = [string]$using:minimumPesterVersion
        $pin = [string]$using:pinnedPesterVersion
        try {
            # The guards sit INSIDE the try so a refused path becomes a named
            # failed file instead of the unhandled error that named nothing on
            # CI (runs 45193295738 / 45376223093).
            if ([string]::IsNullOrWhiteSpace($testPath)) {
                throw 'refusing a null, empty, or whitespace-only test path'
            }
            # A relative path resolves against the worker's own directory, not
            # the repository; refuse it rather than analyse the wrong tree.
            if (-not [System.IO.Path]::IsPathRooted($testPath)) {
                throw "refusing a non-rooted test path: $testPath"
            }
            if ($pin) { Import-Module Pester -RequiredVersion $pin -ErrorAction Stop }
            else { Import-Module Pester -MinimumVersion $minimum -ErrorAction Stop }
            # 6>$null drops the information stream a test's own Write-Host would
            # otherwise interleave across workers; the verdict and the failure
            # text are returned as data below.
            $r = Invoke-Pester -Path $testPath -PassThru -Output None 6>$null
            if ($null -eq $r) {
                [pscustomobject]@{ Name = $name; Total = 0; Passed = 0; Failed = 0; Failures = @(); Error = 'Pester returned no result object.' }
            }
            else {
                $failures = @($r.Failed | ForEach-Object {
                        $path = if ($_.PSObject.Properties['ExpandedPath']) { $_.ExpandedPath } else { $_.Name }
                        $first = @($_.ErrorRecord) | Select-Object -First 1
                        $message = if ($first) { $first.Exception.Message } else { 'failed' }
                        "$path : $message"
                    })
                [pscustomobject]@{ Name = $name; Total = $r.TotalCount; Passed = $r.PassedCount; Failed = $r.FailedCount; Failures = $failures; Error = '' }
            }
        }
        catch {
            [pscustomobject]@{ Name = $name; Total = 0; Passed = 0; Failed = 0; Failures = @(); Error = $_.Exception.Message }
        }
    } -ThrottleLimit $throttle

    $results = @($results | Sort-Object Name)
}

# The verdict rules live in Get-ScriptTestGateVerdict (scripts/verify-helpers.ps1)
# so the tested code is the code this runner executes. This body owns only the
# streams and the exit code.
$verdict = Get-ScriptTestGateVerdict -Results @($results) -DiscoveredCount $testFiles.Count
foreach ($line in $verdict.Lines) { Write-Host $line }
if ($env:GITHUB_ACTIONS -eq 'true') {
    # A verdict error is not a failed test (a zero-test file or a lost worker),
    # so the per-test loop below cannot report it; annotate it here.
    foreach ($message in $verdict.Errors) {
        Write-Host (Format-CheckAnnotation -Title 'test-scripts: verdict error' -Message $message)
    }
}
foreach ($message in $verdict.Errors) { Write-Error -Message $message -ErrorAction Continue }

Write-Host "test-scripts: Total=$($verdict.Total), Passed=$($verdict.Passed), Failed=$($verdict.Failed)"

# A CI failure must be readable from the check-run annotations, because the job
# log needs an authenticated download. GITHUB_ACTIONS gates the prefix so a
# local run keeps Pester's own output instead of printing workflow commands.
if ($env:GITHUB_ACTIONS -eq 'true') {
    foreach ($result in @($results)) {
        $annotationFile = "scripts/$($result.Name)"
        if ($result.Error) {
            Write-Host (Format-CheckAnnotation -Title "test-scripts: $($result.Name)" -Message $result.Error -File $annotationFile)
            continue
        }
        foreach ($failure in @($result.Failures)) {
            Write-Host (Format-CheckAnnotation -Title "test-scripts: $($result.Name)" -Message $failure -File $annotationFile)
        }
    }
}

if ($verdict.ExitCode -ne 0) {
    if ($verdict.Failed -gt 0) {
        Write-Error -Message "test-scripts: $($verdict.Failed) test(s) failed out of $($verdict.Total)" -ErrorAction Continue
    }
    exit $verdict.ExitCode
}

Write-Host "test-scripts: PASS ($($verdict.Total) tests)"
exit 0
