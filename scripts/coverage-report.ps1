#requires -Version 7
<#
.SYNOPSIS
    Merges the per-test-project Cobertura files into ONE report and returns the
    measured line/branch rate per production package.

.DESCRIPTION
    `dotnet test --collect:'XPlat Code Coverage'` writes one
    coverage.cobertura.xml per TEST PROJECT, and each file holds only the
    classes that project's own run happened to execute. The same package
    therefore appears in several files with partial rates: the Core.Tests file
    alone measured GanttCreator.Core at 0.28% line while the union of every
    collected file is ~97%.

    Two consequences follow, and this script exists because of them:

      * No single file may be compared with the thresholds in
        docs/04-TEST-STRATEGY.md. The gate used "the best rate over the files
        wins", which is not the union either - a class covered only by another
        project's run is invisible to every file that did not execute it, so
        the best-file figure under-reports the suite.

      * The merge must be the union over exactly one test run. Collecting into
        a directory that is never cleaned silently mixes runs, and an older
        file can add covered lines the current tree no longer has.

    Steps:
      1. assert the pinned ReportGenerator version (tool-versions.psd1, W11);
      2. install or update the global tool to that pinned version;
      3. run it over every collected file, excluding generated compiler
         sources (**/obj/**) - RegexGenerator output is machine-written and
         cannot be tested, so counting it as uncovered Core code measures the
         compiler, not this repository;
      4. parse the merged Cobertura and return one object per production
         package.

    Measurement only. Judging a rate against a threshold stays in
    scripts/verify.ps1, which owns the policy table.

.PARAMETER ReportsRoot
    Directory holding the collected per-project coverage.cobertura.xml files.
    Defaults to scripts/_artifacts/coverage.
.PARAMETER TargetDir
    Directory the merged report is written to. Defaults to
    scripts/_artifacts/coverage-merged.
.NOTES
    A merge over zero reports is a hard failure: reporting a rate the suite
    never produced is the no-silent-pass defect this gate exists to avoid.
#>
[CmdletBinding()]
param(
    [string]$ReportsRoot,
    [string]$TargetDir
)

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $PSCommandPath
Import-PowerShellDataFile -Path (Join-Path $scriptRoot 'tool-versions.psd1') |
    ForEach-Object { $script:tools = $_ }

if (-not $ReportsRoot) { $ReportsRoot = Join-Path $scriptRoot '_artifacts\coverage' }
if (-not $TargetDir) { $TargetDir = Join-Path $scriptRoot '_artifacts\coverage-merged' }

# ---------------------------------------------------------------- helper functions (available when dot-sourced)
function Test-ToolInstalled {
    param([string]$ToolId)

    # Returns the installed version string, or $null when the tool is absent.
    # Matching the first column exactly avoids matching a tool whose id merely
    # contains this one as a prefix.
    $list = & dotnet tool list --global 2>$null
    foreach ($line in $list) {
        $columns = $line -split '\s{2,}', 3
        if ($columns.Count -ge 2 -and $columns[0].Trim() -eq $ToolId) {
            return $columns[1].Trim()
        }
    }
    return $null
}

function Assert-ReportGeneratorPin {
    param([AllowNull()][object]$Settings)
    # The single guard that keeps an unpinned tool out of the coverage step.
    # Extracted so a Pester test can feed it a missing pin and assert it throws
    # rather than only matching the throw text in this source.
    if (-not $Settings -or -not $Settings.Version) {
        throw 'tool-versions.psd1 has no dotnet-reportgenerator-globaltool pin. The coverage merge tool must be pinned (W11).'
    }
}

function Get-PackageCoverageRate {
    <#
    .SYNOPSIS
        Reads one package's line and branch rate from a merged Cobertura
        report, or returns $null when the package is absent.
    .DESCRIPTION
        Kept pure (report in, object out) so the merge result can be tested
        against a fixture instead of needing a full test run.
    #>
    param(
        [Parameter(Mandatory)][xml]$Report,
        [Parameter(Mandatory)][string]$PackageName
    )

    $package = @($Report.coverage.packages.package) |
        Where-Object { [string]$_.name -eq $PackageName } |
        Select-Object -First 1
    if (-not $package) { return $null }

    return [pscustomobject]@{
        Package    = $PackageName
        LineRate   = [double]$package.'line-rate'
        BranchRate = [double]$package.'branch-rate'
    }
}

function Get-MergedPackageCoverage {
    <#
    .SYNOPSIS
        Returns one object per production package in a merged Cobertura
        report, excluding test assemblies.
    #>
    param([Parameter(Mandatory)][xml]$Report)

    foreach ($package in @($Report.coverage.packages.package)) {
        $name = [string]$package.name
        if ($name -like '*.Tests' -or $name -like '*.TestUtilities') { continue }
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        [pscustomobject]@{
            Package    = $name
            LineRate   = [double]$package.'line-rate'
            BranchRate = [double]$package.'branch-rate'
        }
    }
}

# ---------------------------------------------------------------- main body
# Dot-sourcing (the Pester tests) must only define the helpers above.
if ($MyInvocation.InvocationName -eq '.') { return }

$settings = $script:tools.'dotnet-reportgenerator-globaltool'
Assert-ReportGeneratorPin $settings

$reports = @(Get-ChildItem -LiteralPath $ReportsRoot -Recurse -Filter 'coverage.cobertura.xml' -ErrorAction SilentlyContinue)
if ($reports.Count -eq 0) {
    throw ("No coverage.cobertura.xml under '{0}'. Run the test step with --collect:'XPlat Code Coverage' first: a merge over nothing would report a rate the suite never produced." -f $ReportsRoot)
}

# ---------------------------------------------------------------- install
# The tool is installed by this script, not by a csproj PackageReference, so it
# is test-only and can never reach a production project.
$installedVersion = Test-ToolInstalled $settings.ToolId
if ($installedVersion -ne $settings.Version) {
    if ($null -eq $installedVersion) {
        Write-Host "coverage-report: installing $($settings.ToolId) $($settings.Version) as a global tool..."
        & dotnet tool install --global $settings.ToolId --version $settings.Version
    } else {
        Write-Host "coverage-report: $($settings.ToolId) $installedVersion is installed but $($settings.Version) is required; updating..."
        & dotnet tool update --global $settings.ToolId --version $settings.Version
    }
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to bring $($settings.ToolId) to version $($settings.Version) (installed: $installedVersion)."
    }
} else {
    Write-Host "coverage-report: $($settings.ToolId) $($settings.Version) already installed."
}

# ---------------------------------------------------------------- merge
New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
$reportGlob = Join-Path $ReportsRoot '**\coverage.cobertura.xml'
Write-Host "coverage-report: merging $($reports.Count) report(s) from $ReportsRoot"
# Generated compiler sources are excluded on purpose: RegexGenerator output is
# machine-written and no test can reach it, so including it would count the
# compiler's code as untested repository code.
& reportgenerator "-reports:$reportGlob" "-targetdir:$TargetDir" '-reporttypes:Cobertura' '-filefilters:-**/obj/**' '-verbosity:Warning'
if ($LASTEXITCODE -ne 0) {
    throw "reportgenerator failed with exit code $LASTEXITCODE."
}

$mergedPath = Join-Path $TargetDir 'Cobertura.xml'
if (-not (Test-Path -LiteralPath $mergedPath)) {
    throw "reportgenerator produced no merged report at $mergedPath. Treating this as a FAILURE rather than a pass."
}

[xml]$merged = Get-Content -LiteralPath $mergedPath -Raw
$coverage = @(Get-MergedPackageCoverage -Report $merged)
if ($coverage.Count -eq 0) {
    throw "The merged report at $mergedPath contains no packages. Treating this as a FAILURE rather than a pass."
}

foreach ($package in $coverage) {
    Write-Host ("  - {0}: merged line {1}%, branch {2}%" -f `
            $package.Package,
        [math]::Round($package.LineRate * 100, 2),
        [math]::Round($package.BranchRate * 100, 2))
}

return $coverage
