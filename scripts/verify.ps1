#requires -Version 7
<#
.SYNOPSIS
    Full verification gate. Run on every PR and before release.

.DESCRIPTION
    Slower than verify-quick.ps1. Includes all quick gates plus:
      1. working tree hygiene
      2. markdown link sanity
      3. skill tree in sync
      4. status accuracy
      5. script analyzer (PSScriptAnalyzer)
      6. workflow lint (actionlint)
      7. script lint (Pester)
      8. restore
      9. format (production only; tests/ tolerated per tests/Directory.Build.props)
     10. build Release -warnaserror
     11. publish AddIn (packed XLL)
     12. test (OfficeIntegration excluded)
     13. coverage threshold check
     14. package vulnerability scan
     15. mutation gate (changed Core code)

    Exits non-zero on any failure. Writes a human-readable report to
    scripts/_artifacts/verify.txt.
#>

[CmdletBinding()]
param(
    [string]$Solution = 'GanttCreator.slnx',
    [string]$Configuration = 'Release',
    # The diff base the mutation step judges changed Core code against. The
    # default is the integration branch, which is right for CI. On a long-lived
    # branch the diff against that base exceeds the gate's 30-file scope guard,
    # so the gate refuses to run at all and step 15 can never be exercised; this
    # lets a branch-final run use a nearer base without weakening the default.
    [string]$MutationSince = 'origin/main'
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
# Default the solution against the repository root (not the caller's CWD)
# so the gate works from any working directory. An explicitly supplied
# -Solution is honoured exactly as given.
if (-not $PSBoundParameters.ContainsKey('Solution')) {
    $Solution = Join-Path (Split-Path -Parent $scriptRoot) 'GanttCreator.slnx'
}
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

# Runs FIRST so a dirty tree fails in under a second instead of after the
# ~2-minute suite (the R1.1 session paid that cost). The gate certifies a
# committed state: verify-quick.ps1 accepts staged changes (pre-commit),
# but verify.ps1's coverage/publish artifacts must correspond to the
# committed tree, so staged-but-uncommitted work also counts as dirty.
Invoke-Step 'working tree hygiene' {
    $dirty = git -C (Split-Path -Parent $scriptRoot) status --short
    if ($dirty) {
        Add-Content -LiteralPath $report -Value "Dirty working tree:"
        $dirty | ForEach-Object { Add-Content -LiteralPath $report -Value "  $_" }
        Write-Error 'Working tree is dirty. Commit or stash before running verify.ps1.'
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
    # Pass $Solution explicitly so the step does not depend on the caller's CWD.
    dotnet restore $Solution --locked-mode
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
# thresholds in $coverageThresholds above are the production targets; this
# step parses the measured coverage.cobertura.xml rates and prints them next
# to each target, visibly flagging below-target entries as warnings. It does
# not warn-and-fail. Enforcement is enabled in R3.x when Core has real
# tests (docs/04-TEST-STRATEGY.md).
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
    # Parse the measured line/branch rates from the cobertura XML so each
    # target is reported next to the real number. Cobertura stores rates as
    # invariant-culture decimal fractions (0..1) on the <package> element;
    # the test command writes one coverage file per test project, so a
    # package can appear in several files and the best measured rate wins.
    $measured = @{}
    foreach ($file in $coverageFiles) {
        [xml]$xml = Get-Content -LiteralPath $file.FullName -Raw
        foreach ($pkg in $xml.coverage.packages.package) {
            $name = [string]$pkg.name
            if (-not $measured.ContainsKey($name)) {
                $measured[$name] = @{ Line = 0.0; Branch = 0.0 }
            }
            $lineRate = [double]$pkg.'line-rate'
            $branchRate = [double]$pkg.'branch-rate'
            if ($lineRate -gt $measured[$name].Line)   { $measured[$name].Line = $lineRate }
            if ($branchRate -gt $measured[$name].Branch) { $measured[$name].Branch = $branchRate }
        }
    }
    foreach ($project in $coverageThresholds.Keys) {
        $expected = $coverageThresholds[$project]
        if ($expected.Line -eq 0) { continue }
        if (-not $measured.ContainsKey($project)) {
            $line = "{0}: no coverage package found (target line >= {1}%, branch >= {2}%)" -f $project, $expected.Line, $expected.Branch
            Write-Host "  - $line"
            Add-Content -LiteralPath $report -Value "  $line"
            continue
        }
        $linePct = [math]::Round($measured[$project].Line * 100, 2)
        $branchPct = [math]::Round($measured[$project].Branch * 100, 2)
        $flag = ''
        if ($linePct -lt $expected.Line -or $branchPct -lt $expected.Branch) { $flag = '  <-- below target (warning)' }
        $line = "{0}: measured line {1}%, branch {2}% (target line >= {3}%, branch >= {4}%; enforcement deferred to R3.x per docs/04-TEST-STRATEGY.md){5}" -f $project, $linePct, $branchPct, $expected.Line, $expected.Branch, $flag
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

# R3.13 (the R0.5 deferral). Changed-Core-code mutation score, judged against
# the pinned threshold in scripts/tool-versions.psd1. Runs AFTER the test step
# so the score is measured against a suite that has just passed.
#
# The gate script judges the Stryker exit code AFTER capturing its output and
# returns a result object rather than calling `exit`; this step therefore judges
# the returned object. That is the L10 lesson applied twice: `-EnableExit` from
# inside this Tee/ForEach pipeline reported false PASSes for the life of the
# PSScriptAnalyzer gate.
# Re-enabled once the first full-Core baseline was archived
# (scripts/_artifacts/mutation/reports/mutation-report.json; judged score
# 76.22% against the pinned 80%; D5 - archived, not gated: the threshold
# applies from the next Core change). Runs changed-code only via
# --since origin/main so a local verify is not slowed by whole-Core mutation.
# The gate script judges the Stryker exit code AFTER capturing its output and
# returns a result object; this step decides pass/fail from it. That is the
# L10 lesson applied twice: never report PASS from a non-throwing call or a
# silent step body.
Invoke-Step 'mutation gate (changed Core code)' {
    $result = $null
    $halted = $false
    try {
        # The base is passed explicitly so CI keeps judging against the
        # integration branch while a long-lived branch can name a nearer one.
        Write-Host "  -> mutation gate: diff base $MutationSince"
        $result = & $scriptRoot\mutation-gate.ps1 -Since $MutationSince
    } catch {
        $halted = $true
        Write-Host "  -> mutation gate (changed Core code): FAIL (mutation-gate.ps1 halted: $($_.Exception.Message))" -ForegroundColor Red
    }
    if (-not $halted) {
        $scored = $result.Killed + $result.Survived + $result.Timeout
        if ($scored -eq 0) {
            Write-Host "  -> mutation gate (changed Core code): PASS (no scored changed-code mutants)"
        } elseif ($result.ScorePercent -lt $result.Threshold) {
            Write-Host "  -> mutation gate (changed Core code): FAIL (score $($result.ScorePercent)% below threshold $($result.Threshold)%)" -ForegroundColor Red
            exit 1
        } else {
            Write-Host "  -> mutation gate (changed Core code): PASS ($($result.ScorePercent)% >= $($result.Threshold)%)"
        }
    } else {
        exit 1
    }
}

$end = Get-Date
$elapsed = ($end - $start).TotalSeconds
$line = "`nverify: PASS in {0:N1}s. Report: {1}" -f $elapsed, $report
Write-Host $line -ForegroundColor Green
Add-Content -LiteralPath $report -Value $line
exit 0
