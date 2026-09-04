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
      2. dotnet restore --locked-mode (after a lock file is produced)
      3. dotnet build -c Release -warnaserror
      4. dotnet publish the AddIn in Release configuration
      5. dotnet test on Core, Raster, Office contract, AddIn, Architecture
         (OfficeIntegration trait excluded)
#>

[CmdletBinding()]
param(
    [string]$Solution = 'GanttCreator.slnx',
    [string]$Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $PSCommandPath
$artifacts  = Join-Path $scriptRoot '_artifacts'
if (-not (Test-Path $artifacts)) { New-Item -ItemType Directory -Path $artifacts | Out-Null }
$report = Join-Path $artifacts 'verify-quick.txt'
$start  = Get-Date
"" | Set-Content -LiteralPath $report

function Run-Step {
    param([string]$Name, [scriptblock]$Block)
    $line = "[{0:HH:mm:ss}] {1}" -f (Get-Date), $Name
    Write-Host $line
    Add-Content -LiteralPath $report -Value $line
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
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

Run-Step 'dotnet --version' { dotnet --version }

Run-Step 'clinerules skill tree in sync' {
    # Drift gate: re-run the sync and fail if any of the three views
    # (.clinerules/, .cline/skills/, docs/) is out of date. Wired in so
    # a stale rule or skill is caught at every commit.
    pwsh -NoProfile -File (Join-Path $PSScriptRoot 'check-cline-skills.ps1')
}

Run-Step 'restore (locked mode if lock exists)' {
    # The .NET SDK with RestorePackagesWithLockFile produces
    # packages.lock.json files next to each project that has
    # dependencies. Detect any of them to decide whether --locked-mode
    # is available.
    $lockFiles = Get-ChildItem -Path $PSScriptRoot\..\src -Recurse -Filter 'packages.lock.json' -ErrorAction SilentlyContinue
    if ($lockFiles) {
        dotnet restore --locked-mode
    } else {
        Write-Host 'NOTE: no packages.lock.json found; using plain dotnet restore (R0.3 will add the lock file).'
        dotnet restore $Solution
    }
}

Run-Step 'build Release -warnaserror' {
    dotnet build $Solution -c $Configuration --no-restore -warnaserror
}

Run-Step 'publish AddIn Release' {
    dotnet publish (Join-Path $scriptRoot '..\src\GanttCreator.AddIn\GanttCreator.AddIn.csproj') `
        -c Release --no-restore -warnaserror
}

Run-Step 'test (OfficeIntegration excluded)' {
    dotnet test $Solution -c $Configuration --no-build --no-restore `
        --filter 'Category!=OfficeIntegration'
}

$end = Get-Date
$elapsed = ($end - $start).TotalSeconds
$line = "`nverify-quick: PASS in {0:N1}s. Report: {1}" -f $elapsed, $report
Write-Host $line -ForegroundColor Green
Add-Content -LiteralPath $report -Value $line
exit 0
