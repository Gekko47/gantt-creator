#requires -Version 7
<#
.SYNOPSIS
    Full verification gate. Run on every PR and before release.

.DESCRIPTION
    Slower than verify-quick.ps1. Includes all quick gates plus:
      1. markdown link sanity (scripts/check-md-links.ps1)
      2. skill-tree drift gate (scripts/check-cline-skills.ps1)
      3. STATUS.md accuracy (scripts/check-status.ps1)
      4. PSScriptAnalyzer over scripts/
      5. workflow lint (actionlint) over .github/workflows/ci.yml
      6. script lint (Pester) over scripts/*.ps1
      7. dotnet restore --locked-mode
      8. dotnet format --verify-no-changes --exclude tests
      9. dotnet build -c Release -warnaserror
     10. dotnet publish the AddIn in Release configuration
     11. dotnet test on every non-OfficeIntegration project
     12. coverage report and per-project threshold check
     13. dotnet list package --vulnerable --include-transitive
     14. repository hygiene (git status --short, dirty working tree)

    Exits non-zero on any failure. Writes a human-readable report to
    scripts/_artifacts/verify.txt.
#>

[CmdletBinding()]
param(
    [string]$Solution = 'GanttCreator.slnx',
    [string]$Configuration = 'Release'
)

# Coverage thresholds from docs/04-TEST-STRATEGY.md.
$coverageThresholds = @{
    'GanttCreator.Core'        = @{ Line = 95; Branch = 90 }
    'GanttCreator.Raster'      = @{ Line = 90; Branch = 85 }
    'GanttCreator.Office'      = @{ Line = 80; Branch = 70 }
    'GanttCreator.AddIn'       = @{ Line = 75; Branch = 65 }
    'GanttCreator.Architecture.Tests' = @{ Line = 0;  Branch = 0  }  # not subject to coverage
}

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $PSCommandPath
. (Join-Path $scriptRoot 'verify-helpers.ps1')
$artifacts  = Join-Path $scriptRoot '_artifacts'
if (-not (Test-Path -LiteralPath $artifacts)) { New-Item -ItemType Directory -Path $artifacts | Out-Null }
$report = Join-Path $artifacts 'verify.txt'
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
            $stepOut | Select-Object -Last 50 | ForEach-Object { Add-Content -LiteralPath $report -Value $_ }
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

Invoke-Step 'markdown link sanity' {
    pwsh -NoProfile -File (Join-Path $scriptRoot 'check-md-links.ps1')
}

Invoke-Step 'skill tree in sync' {
    pwsh -NoProfile -File (Join-Path $scriptRoot 'check-cline-skills.ps1')
}

Invoke-Step 'status accuracy' {
    pwsh -NoProfile -File (Join-Path $scriptRoot 'check-status.ps1')
}

Invoke-Step 'script analyzer (PSScriptAnalyzer)' {
    # Same gate as verify-quick.ps1: Invoke-PssaGate (scripts/verify-helpers.ps1)
    # captures findings and judges in the step's scope; never `-EnableExit`
    # (its function-level `exit` is swallowed by the Tee/ForEach pipeline).
    Invoke-PssaGate -Path (Join-Path $PSScriptRoot '.') `
        -Settings (Join-Path $PSScriptRoot 'PSScriptAnalyzerSettings.psd1') `
        -Report $report
}

Invoke-Step 'workflow lint (actionlint)' {
    pwsh -NoProfile -File (Join-Path $PSScriptRoot 'lint-ci.ps1')
}

Invoke-Step 'script lint (Pester)' {
    pwsh -NoProfile -File (Join-Path $PSScriptRoot 'test-scripts.ps1')
}

Invoke-Step 'restore' {
    dotnet restore --locked-mode
}

Invoke-Step 'format (production only; tests/ tolerated per tests/Directory.Build.props)' {
    # Same command as the CI format gate (.github/workflows/ci.yml) and
    # verify-quick.ps1 so a local PASS predicts a CI PASS. Excludes tests/
    # where the xUnit style conventions are deliberately tolerated.
    dotnet format $Solution --verify-no-changes --exclude tests
}

Invoke-Step 'build Release -warnaserror' {
    # Shared entry point with verify-quick.ps1 step 10 and the GitHub CI
    # 'Build Release' step (W8 local/CI parity).
    pwsh -NoProfile -File (Join-Path $PSScriptRoot 'build-release.ps1') -Solution $Solution -Configuration $Configuration
}

Invoke-Step 'publish AddIn (packed XLL)' {
    # Shared entry point with verify-quick.ps1 step 11 and the GitHub CI
    # 'Publish AddIn' step (W8 local/CI parity).
    pwsh -NoProfile -File (Join-Path $PSScriptRoot 'publish-addin.ps1') -Configuration $Configuration
}

Invoke-Step 'test (OfficeIntegration excluded)' {
    dotnet test $Solution -c $Configuration --no-build --no-restore `
        --filter 'Category!=OfficeIntegration' `
        --collect:'XPlat Code Coverage' --results-directory (Join-Path $artifacts 'coverage')
}

# The coverage-threshold check is intentionally permissive at R0.x. The
# thresholds in $coverageThresholds above are the production targets, but
# this step only records the targets next to the measured coverage; it does
# not warn and it does not fail. Enforcement is enabled in R3.x when Core
# has real tests (docs/04-TEST-STRATEGY.md).
Invoke-Step 'coverage threshold check' {
    $coverageRoot = Join-Path $artifacts 'coverage'
    if (-not (Test-Path -LiteralPath $coverageRoot)) {
        Write-Host 'coverage root not present; skipping threshold check.'
        return
    }
    $coverageFiles = Get-ChildItem -Path $coverageRoot -Recurse -Filter 'coverage.cobertura.xml'
    if (-not $coverageFiles) {
        Write-Host 'no coverage.cobertura.xml found; skipping threshold check.'
        return
    }
    foreach ($project in $coverageThresholds.Keys) {
        $expected = $coverageThresholds[$project]
        if ($expected.Line -eq 0) { continue }
        $line = "{0}: target line >= {1}%, branch >= {2}% (threshold enforcement deferred to R3.x per docs/04-TEST-STRATEGY.md)" -f $project, $expected.Line, $expected.Branch
        Write-Host "  - $line"
        Add-Content -LiteralPath $report -Value "  $line"
    }
}

Invoke-Step 'package vulnerability scan' {
    dotnet list $Solution package --vulnerable --include-transitive
    # dotnet list does not set a non-zero exit on found vulnerabilities, so
    # we cannot make this step a hard fail until the team approves an
    # explicit vulnerability gate. For now this is informational.
}

Invoke-Step 'working tree hygiene' {
    $dirty = git status --short
    if ($dirty) {
        Add-Content -LiteralPath $report -Value "Dirty working tree:"
        $dirty | ForEach-Object { Add-Content -LiteralPath $report -Value "  $_" }
        Write-Error 'Working tree is dirty. Commit or stash before running verify.ps1.'
        exit 1
    }
}

$end = Get-Date
$elapsed = ($end - $start).TotalSeconds
$line = "`nverify: PASS in {0:N1}s. Report: {1}" -f $elapsed, $report
Write-Host $line -ForegroundColor Green
Add-Content -LiteralPath $report -Value $line
exit 0
