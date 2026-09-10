#requires -Version 7
<#
.SYNOPSIS
    Quick verification gate. Run on every commit.

.DESCRIPTION
    Must complete in under ~2 minutes on a developer machine. Excludes the
    Office-integration suite (per docs/adr/0001). Exits non-zero on any
    failure with the failing step's exit code preserved. Writes a
    human-readable report to scripts/_artifacts/verify-quick.txt.

    Steps:
      1. dotnet --version sanity check
      2. skill-tree drift gate
      3. STATUS.md accuracy gate (hashes, paths, roadmap IDs)
      4. PSScriptAnalyzer over scripts/
      5. Workflow lint (actionlint) over .github/workflows/ci.yml
      6. Script lint (Pester) over scripts/*.ps1
      7. dotnet restore --locked-mode (after a lock file is produced)
      8. dotnet format --verify-no-changes --exclude tests
      9. dotnet build -c Release -warnaserror
     10. dotnet publish GanttCreator.AddIn (packed XLL for AddInAssemblyTests)
     11. dotnet test on Core, Raster, Office contract, AddIn, Architecture
         (OfficeIntegration trait excluded)
#>

[CmdletBinding()]
param(
    [string]$Solution = 'GanttCreator.slnx',
    [string]$Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $PSCommandPath
. (Join-Path $scriptRoot 'verify-helpers.ps1')
# Default the solution against the repository root (not the caller's CWD)
# so the gate works from any working directory. An explicitly supplied
# -Solution is honoured exactly as given.
if (-not $PSBoundParameters.ContainsKey('Solution')) {
    $Solution = Join-Path (Split-Path -Parent $scriptRoot) 'GanttCreator.slnx'
}
$artifacts  = Join-Path $scriptRoot '_artifacts'
if (-not (Test-Path -LiteralPath $artifacts)) { New-Item -ItemType Directory -Path $artifacts | Out-Null }
$report = Join-Path $artifacts 'verify-quick.txt'
$start  = Get-Date
"" | Set-Content -LiteralPath $report

function Invoke-Step {
    param([string]$Name, [scriptblock]$Block)
    $line = "[{0:HH:mm:ss}] {1}" -f (Get-Date), $Name
    Write-Host $line
    Add-Content -LiteralPath $report -Value $line
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    # Deterministic step result: clear the process-wide $LASTEXITCODE so a
    # step that runs only cmdlets (or returns early) cannot inherit a stale
    # exit code from the previous native command and report a false PASS.
    $global:LASTEXITCODE = 0
    try {
        & $Block 2>&1 | Tee-Object -Variable stepOut | ForEach-Object { Add-Content -LiteralPath $report -Value $_ }
        $sw.Stop()
        $status = if ($LASTEXITCODE -eq 0) { 'PASS' } else { "FAIL (exit $LASTEXITCODE)" }
        $line = "  -> {0} in {1:N1}s" -f $status, $sw.Elapsed.TotalSeconds
        Write-Host $line
        Add-Content -LiteralPath $report -Value $line
        if ($LASTEXITCODE -ne 0) {
            Add-Content -LiteralPath $report -Value "Last output:"
            $stepOut | Select-Object -Last 30 | ForEach-Object { Add-Content -LiteralPath $report -Value $_ }
            Write-Error "Step '$Name' failed with exit $LASTEXITCODE. See $report."
            exit $LASTEXITCODE
        }
    } catch {
        $sw.Stop()
        $msg = $_.Exception.Message
        $line = "  -> FAIL (exception) in {0:N1}s: {1}" -f $sw.Elapsed.TotalSeconds, $msg
        Write-Host $line -ForegroundColor Red
        Add-Content -LiteralPath $report -Value $line
        exit 1
    }
}

Invoke-Step 'dotnet --version' { dotnet --version }

Invoke-Step 'skill tree in sync' {
    # Drift gate: re-run the sync and fail if either view
    # (.cline/skills/ or docs/) is out of date. Wired in so
    # a stale skill is caught at every commit.
    pwsh -NoProfile -File (Join-Path $PSScriptRoot 'check-cline-skills.ps1')
}

Invoke-Step 'status accuracy' {
    # Accuracy gate for docs/STATUS.md: every commit hash referenced must
    # resolve, every backticked repo path claimed must exist, and every
    # roadmap ID mentioned must be defined in docs/03-ROADMAP.md.
    pwsh -NoProfile -File (Join-Path $PSScriptRoot 'check-status.ps1')
}

Invoke-Step 'script analyzer (PSScriptAnalyzer)' {
    # Lint every script in scripts/ against scripts/PSScriptAnalyzerSettings.psd1.
    # Exclusions are documented in that file. If the module is missing,
    # install it once per machine:
    #   Install-Module PSScriptAnalyzer -Scope CurrentUser -Force -SkipPublisherCheck
    #
    # The capture-then-judge gate lives in Invoke-PssaGate
    # (scripts/verify-helpers.ps1) so pssa-gate.Tests.ps1 exercises the
    # same code this step runs, never a copy. It must not rely on
    # `-EnableExit`: the function-level `exit` does not survive the
    # Tee/ForEach pipeline in Invoke-Step (W13, CI run 2026-09-07).
    Invoke-PssaGate -Path (Join-Path $PSScriptRoot '.') `
        -Settings (Join-Path $PSScriptRoot 'PSScriptAnalyzerSettings.psd1') `
        -Report $report
}

Invoke-Step 'workflow lint (actionlint)' {
    # Lint .github/workflows/ci.yml with pinned actionlint binary (L6).
    # Downloads the binary once per machine to TEMP.
    pwsh -NoProfile -File (Join-Path $PSScriptRoot 'lint-ci.ps1')
}

Invoke-Step 'script lint (Pester)' {
    # Run Pester unit tests for scripts/*.ps1 (L7).
    # Discovers *.Tests.ps1 files and runs them with -PassThru.
    pwsh -NoProfile -File (Join-Path $PSScriptRoot 'test-scripts.ps1')
}

Invoke-Step 'restore' {
    # Per-project packages.lock.json files are committed for src/ and
    # tests/ (RestorePackagesWithLockFile=true in Directory.Build.props),
    # so --locked-mode is always available. This matches verify.ps1 and
    # proves the lock files are honoured without network access.
    # Pass $Solution explicitly so the step does not depend on the caller's CWD.
    dotnet restore $Solution --locked-mode
}

Invoke-Step 'format (production only; tests/ tolerated per tests/Directory.Build.props)' {
    # Same command as the CI format gate (.github/workflows/ci.yml) so a
    # local PASS predicts a CI PASS. Excludes tests/ where the xUnit style
    # conventions are deliberately tolerated (EnforceCodeStyleInBuild=false).
    dotnet format $Solution --verify-no-changes --exclude tests
}

Invoke-Step 'build Release -warnaserror' {
    # Shared entry point with verify.ps1 step 10 and the GitHub CI
    # 'Build Release' step (W8 local/CI parity).
    pwsh -NoProfile -File (Join-Path $PSScriptRoot 'build-release.ps1') -Solution $Solution -Configuration $Configuration
}

Invoke-Step 'publish AddIn (packed XLL)' {
    # Shared entry point with verify.ps1 step 11 and the GitHub CI
    # 'Publish AddIn' step (W8 local/CI parity).
    pwsh -NoProfile -File (Join-Path $PSScriptRoot 'publish-addin.ps1') -Configuration $Configuration
}

Invoke-Step 'test (OfficeIntegration excluded)' {
    # Shared entry point with the GitHub CI 'Test (OfficeIntegration
    # excluded)' step (W8 local/CI parity).
    pwsh -NoProfile -File (Join-Path $PSScriptRoot 'test-non-office.ps1') -Solution $Solution -Configuration $Configuration
}

$end = Get-Date
$elapsed = ($end - $start).TotalSeconds
$line = "`nverify-quick: PASS in {0:N1}s. Report: {1}" -f $elapsed, $report
Write-Host $line -ForegroundColor Green
Add-Content -LiteralPath $report -Value $line
exit 0
