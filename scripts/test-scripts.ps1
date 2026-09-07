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

    Exit 0 on all tests pass; exit 1 on any failure.
#>

[CmdletBinding()]
param(
    [string]$ScriptsRoot = 'scripts'
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$scriptsDir = Join-Path $repoRoot $ScriptsRoot

if (-not (Test-Path -LiteralPath $scriptsDir)) {
    Write-Error "Scripts directory not found: $scriptsDir"
    exit 1
}

# Ensure Pester 5+ is available. Get-Module does not support
# -MinimumVersion; check the highest installed version and install
# only if it is missing or below 5.0.
$installedPester = Get-Module -ListAvailable -Name Pester -ErrorAction SilentlyContinue |
    Sort-Object Version -Descending |
    Select-Object -First 1
if (-not $installedPester -or $installedPester.Version -lt [Version]'5.0') {
    Write-Host "Installing Pester 5..."
    try {
        Install-Module -Name Pester -MinimumVersion 5.0 -Scope CurrentUser -Force -SkipPublisherCheck
    }
    catch {
        Write-Error "Failed to install Pester: $_"
        exit 1
    }
}

# Import Pester
try {
    Import-Module Pester -MinimumVersion 5.0 -ErrorAction Stop
}
catch {
    Write-Error "Pester 5+ not available: $_"
    exit 1
}

# Find test files: *.Tests.ps1 or *Tests.ps1 in scripts/
$testFiles = Get-ChildItem -Path $scriptsDir -Filter '*.Tests.ps1' -ErrorAction SilentlyContinue
if ($testFiles.Count -eq 0) {
    $testFiles = Get-ChildItem -Path $scriptsDir -Filter '*Tests.ps1' -ErrorAction SilentlyContinue
}

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

# Run Pester on each test file and aggregate results
$allPassed = $true
$totalTests = 0
$passedTests = 0
$failedTests = 0

foreach ($testFile in $testFiles) {
    Write-Host "Running $($testFile.Name)..."
    # Pester 5 and 6 both expose PassedCount / FailedCount / TotalCount on
    # the -PassThru result. The Pester 4-era `.TestResult` collection is NOT
    # version-safe: under Pester 6 it no longer lists every test, which made
    # this gate report "Total=0, Passed=0" while tests actually ran.
    $result = Invoke-Pester -Path $testFile.FullName -PassThru -Output Detailed
    if ($null -eq $result) {
        Write-Error "test-scripts: Pester returned no result object for $($testFile.Name)."
        exit 1
    }
    if ($result.TotalCount -eq 0) {
        # No-silent-pass policy: a discovered test file that contains zero
        # tests proves nothing and must fail the gate.
        Write-Error "test-scripts: $($testFile.Name) contains no Pester tests."
        exit 1
    }
    $totalTests  += $result.TotalCount
    $passedTests += $result.PassedCount
    $failedTests += $result.FailedCount

    if ($result.FailedCount -gt 0) {
        $allPassed = $false
        Write-Host "  FAILED: $($result.FailedCount) of $($result.TotalCount) tests failed in $($testFile.Name)"
    } else {
        Write-Host "  PASSED: $($result.TotalCount) tests"
    }
}

Write-Host "test-scripts: Total=$totalTests, Passed=$passedTests, Failed=$failedTests"

if (-not $allPassed) {
    Write-Error "test-scripts: $failedTests test(s) failed out of $totalTests"
    exit 1
}

Write-Host "test-scripts: PASS ($totalTests tests)"
exit 0