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

if (-not (Test-Path $scriptsDir)) {
    Write-Error "Scripts directory not found: $scriptsDir"
    exit 1
}

# Ensure Pester 5+ is available
if (-not (Get-Module -ListAvailable -Name Pester -ErrorAction SilentlyContinue)) {
    Write-Host "Installing Pester 5..."
    try {
        Install-Module -Name Pester -MinimumVersion 5.0 -MaximumVersion 6 -Scope CurrentUser -Force -SkipPublisherCheck
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
    Write-Host "test-scripts: No test files found (looking for *.Tests.ps1 or *Tests.ps1 in $scriptsDir)"
    Write-Host "test-scripts: PASS (nothing to test)"
    exit 0
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
    $result = Invoke-Pester -Script @($testFile.FullName) -PassThru -Output Detailed
    $totalTests += $result.TestResult.Count
    $passedTests += $result.TestResult.PassedCount
    $failedTests += $result.TestResult.FailedCount

    if ($result.FailedCount -gt 0) {
        $allPassed = $false
        Write-Host "  FAILED: $($result.FailedCount) of $($result.TestResult.Count) tests failed in $($testFile.Name)"
    } else {
        Write-Host "  PASSED: $($result.TestResult.Count) tests"
    }
}

Write-Host "test-scripts: Total=$totalTests, Passed=$passedTests, Failed=$failedTests"

if (-not $allPassed) {
    Write-Error "test-scripts: $failedTests test(s) failed out of $totalTests"
    exit 1
}

Write-Host "test-scripts: PASS ($totalTests tests)"
exit 0