#requires -Version 7
<#
.SYNOPSIS
    Mutation-testing gate for changed Core code (R3.13, the R0.5 deferral).

.DESCRIPTION
    Runs Stryker.NET against GanttCreator.Core through the Core test project
    and judges the result against the changed-code threshold from
    docs/04-TEST-STRATEGY.md ("Coverage policy", 80%).

    WHY THIS IS A SEPARATE SCRIPT. D3: the run's exit code is judged AFTER the
    output is captured, never via `-EnableExit`. That is the L10 lesson: a
    function-level `exit` does not propagate through the Invoke-Step
    Tee/ForEach pipeline, so every "-EnableExit reports PASS on failure" claim
    in this repo's history was a false PASS. This script therefore returns a
    result object and lets the CALLER decide, and scripts/verify.ps1 invokes it
    the same way it invokes every other gate.

    SCOPE. Core only, and only changed code:
      * `--project` / `--test-project` are explicit csproj paths, so the gate
        does not depend on Stryker parsing the .slnx solution file. The R3.13
        guide named ".slnx gaps" as a stop condition; the probe showed it does
        not apply, and both flags are verified against the installed CLI.
      * `-Since` (diff-compare) restricts the run to files changed against a
        merge base, which is what makes an 80% threshold on *changed* code
        meaningful rather than a whole-project number that only falls.

    The first full-Core run (no `-Since`) archives the BASELINE. Per D5 the
    threshold does not gate this row retroactively; it applies from the next
    Core change onward. Pass -Baseline to take that archived run.

.PARAMETER Baseline
    Run the whole of Core and archive the baseline score instead of judging
    changed code against the threshold.

.PARAMETER Threshold
    Override the threshold. A change here needs a stated reason; never lower it
    "to make it pass".

.PARAMETER Since
    The committish `--since` diffs against. Defaults to `origin/main`, which
    is this repository's integration branch.

.EXAMPLE
    pwsh -NoProfile -File scripts/mutation-gate.ps1 -Baseline
    pwsh -NoProfile -File scripts/mutation-gate.ps1
#>
[CmdletBinding()]
param(
    [switch]$Baseline,
    [int]$Threshold,
    [string]$Since = 'origin/main',
    [string]$Solution,
    [string]$Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $PSCommandPath
$repoRoot = Split-Path -Parent $scriptRoot
Import-PowerShellDataFile -Path (Join-Path $scriptRoot 'tool-versions.psd1') |
    ForEach-Object { $script:tools = $_ }

if (-not $PSBoundParameters.ContainsKey('Solution')) {
    $Solution = Join-Path $repoRoot 'GanttCreator.slnx'
}

$settings = $script:tools.'dotnet-stryker'
if (-not $settings) {
    throw 'tool-versions.psd1 has no dotnet-stryker pin. The mutation tool must be pinned (W11).'
}
if ($PSBoundParameters.ContainsKey('Threshold')) {
    # An override is allowed but must be visible, so it is echoed. Silent
    # threshold drift is how a coverage gate stops meaning anything.
    Write-Host "mutation-gate: THRESHOLD OVERRIDDEN to $Threshold (pinned value is $($settings.ThresholdHigh))."
} else {
    $Threshold = $settings.ThresholdHigh
}

$coreProject = Join-Path $repoRoot 'src\GanttCreator.Core\GanttCreator.Core.csproj'
$coreTests = Join-Path $repoRoot 'tests\GanttCreator.Core.Tests\GanttCreator.Core.Tests.csproj'
foreach ($required in @($Solution, $coreProject, $coreTests)) {
    if (-not (Test-Path -LiteralPath $required)) {
        throw "Required path not found: $required"
    }
}

$outputDir = Join-Path $scriptRoot '_artifacts\mutation'
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

function Test-ToolInstalled {
    param([string]$ToolId)

    $list = & dotnet tool list --global 2>$null
    return $null -ne ($list | Where-Object { $_ -match [regex]::Escape($ToolId) })
}

# ---------------------------------------------------------------- install
# The tool is installed by the GATE, not by a csproj PackageReference, so it
# is test-only and can never reach a production project.
if (-not (Test-ToolInstalled $settings.ToolId)) {
    Write-Host "mutation-gate: installing $($settings.ToolId) $($settings.Version) as a local tool..."
    & dotnet tool install --global $settings.ToolId --version $settings.Version
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to install $($settings.ToolId) $($settings.Version)."
    }
} else {
    Write-Host "mutation-gate: $($settings.ToolId) already installed."
}

# ---------------------------------------------------------------- run
$arguments = @(
    '--project', $coreProject
    '--test-project', $coreTests
    '--configuration', $Configuration
    # ONE reporter, and only one: `Json`. The gate judges the JSON report
    # file, not the console stream, so a second reporter adds nothing the
    # gate needs. It also removes a whole class of CLI ambiguity -- probing
    # the installed dotnet-stryker 5.0.0 showed that `Json,ClearText` and
    # `--reporter Json --reporter ClearText` both fail at argument parsing
    # with the banner and no report, which is the "no report produced" case
    # this gate must never read as a pass. The captured stdout is still
    # echoed below, so a human reading the CI log sees the score.
    '--reporter', 'Json',
    '--output', $outputDir
    # --threshold-low is the HARD gate. It is passed only when judging
    # changed code; the baseline run deliberately omits it so the archived
    # score is the measurement rather than a pass/fail.
    '--verbosity', 'info'
)
if (-not $Baseline) {
    $arguments += @('--since', $Since, '--threshold-low', "$Threshold")
} else {
    Write-Host 'mutation-gate: BASELINE run over the whole of Core (no --since, no threshold).'
}

Write-Host "mutation-gate: dotnet-stryker $($arguments -join ' ')"
$output = & dotnet-stryker @arguments 2>&1
$exit = $LASTEXITCODE

# The run's text output is echoed so a human reading the CI log sees the
# score, not just a status word.
$output | ForEach-Object { Write-Host $_ }

# ---------------------------------------------------------------- judge
# Judge the REPORT, not the process code alone. A run that died before
# producing a report must never be read as a pass: that is the L10/L18
# "gate reports PASS without running" class.
$reportPath = Join-Path $outputDir 'mutation-report.json'
if (-not (Test-Path -LiteralPath $reportPath)) {
    throw "Stryker produced no report at $reportPath. The run did not complete; treating it as a FAILURE rather than a pass."
}

$report = Get-Content -LiteralPath $reportPath -Raw | ConvertFrom-Json

# Stryker writes the score as a fraction (0..1) on MutationScore.
$score = [double]$report.MutationScore
$scorePct = [math]::Round($score * 100, 2)

$total = [int]$report.Total
$killed = [int]$report.Killed
$survived = [int]$report.Survived
$timeout = [int]$report.Timeout
$compileErrors = [int]$report.CompileErrors

$result = [pscustomobject]@{
    Mode            = if ($Baseline) { 'baseline' } else { 'changed-code' }
    ScorePercent    = $scorePct
    Threshold       = $Threshold
    Total           = $total
    Killed          = $killed
    Survived        = $survived
    Timeout         = $timeout
    CompileErrors   = $compileErrors
    ReportPath      = $reportPath
    StrykerExitCode = $exit
}

Write-Host ''
Write-Host "mutation-gate: $($result.Mode) score $($scorePct)% ($killed/$total killed, $survived survived, $timeout timeout) against a $Threshold% threshold."

if ($Baseline) {
    Write-Host "mutation-gate: BASELINE archived (not judged). Report: $reportPath"
    return $result
}

# `$exit` is judged TOO: a non-zero Stryker exit with a passing score still
# means the tool reported a problem, and swallowing it is how a gate starts
# lying. Both conditions must hold for the step to pass.
if ($exit -ne 0) {
    Write-Error "Stryker exited $exit despite a score of $scorePct%. Failing the gate rather than reporting a pass."
    return $result
}
if ($scorePct -lt $Threshold) {
    Write-Error "Mutation score $scorePct% is below the $Threshold% changed-code threshold. Do NOT lower the threshold to make it pass - investigate the surviving mutants."
    return $result
}

Write-Host 'mutation-gate: PASS'
return $result
