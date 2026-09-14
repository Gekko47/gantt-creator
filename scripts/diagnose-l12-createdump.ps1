#requires -Version 7
<#
.SYNOPSIS
    Diagnose the root cause of the ~0.2s createdump crash seen with
    --blame-hang-timeout (L12). Run on the self-hosted runner that exhibited
    the failure.

.DESCRIPTION
    Two diagnostic steps, in order of cost:

    1. Cheaper native-flag alternative: run the OfficeIntegration tests with
       --blame-hang-timeout re-enabled but with --blame-hang-dump-type none.
       If the testhost survives, the original failure was the dump-writing path
       (createdump), not the hang signalling itself.

    2. Antivirus / EDR debug-privilege check: run createdump against a
       long-lived test process and observe whether it can obtain the debug
       privilege required to write a dump. If the dump fails with an
       access-denied pattern while the same binary works on a clean host, the
       failure is EDR blocking.

    The script exits 0 when both steps complete without throwing. Step results
    are written to the report file and to the console; a human must interpret
    them -- the script does not declare the root cause solved.

    Requires: a real Microsoft 365 x64 install (per docs/adr/0001),
    dotnet on PATH, and the GanttCreator.sln at the repo root.
#>

[CmdletBinding()]
param(
    [string]$Solution = 'GanttCreator.slnx',
    [string]$Configuration = 'Release',
    [int]$Step1TimeoutSeconds = 60,
    [int]$Step1DeadlineSeconds = 120
)

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $PSCommandPath
$artifacts  = Join-Path $scriptRoot '_artifacts'
$evidence   = Join-Path $artifacts 'l12-diagnostic'
if (-not (Test-Path -LiteralPath $artifacts)) { New-Item -ItemType Directory -Path $artifacts -Force | Out-Null }
if (-not (Test-Path -LiteralPath $evidence))  { New-Item -ItemType Directory -Path $evidence  -Force | Out-Null }
$report = Join-Path $evidence 'l12-diagnostic-report.txt'
"" | Set-Content -LiteralPath $report -Encoding utf8NoBOM

function Log { param($s) Write-Host $s; Add-Content -LiteralPath $report -Value $s }

Log "L12 createdump diagnostic started $(Get-Date -Format 'o')"
Log "Solution:      $Solution"
Log "Configuration: $Configuration"
Log "Step1 timeout: $Step1TimeoutSeconds s"
Log "Step1 deadline:$Step1DeadlineSeconds s"

# ------------------------------------------------------------------
# Shared pre-flight: confirm Office and dotnet are present, as the real
# OfficeIntegration suite requires both. This is the same pre-flight shape
# verify-office.ps1 uses, so the diagnostic is runnable on the same host.
# ------------------------------------------------------------------
$cfg = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
if (-not $cfg) {
    Log 'FAIL: Office ClickToRun registry key not found. Install Microsoft 365 first.'
    exit 2
}
Log "Office VersionToReport: $($cfg.VersionToReport)"
Log "Office Platform:        $($cfg.Platform)"
Log "Office ClientCulture:   $($cfg.ClientCulture)"

$dotnet = Get-Command dotnet -ErrorAction SilentlyContinue
if (-not $dotnet) {
    Log 'FAIL: dotnet not found on PATH. Add the .NET SDK to PATH first.'
    exit 2
}
Log "dotnet: $($dotnet.Source) $($dotnet.Version)"

# ------------------------------------------------------------------
# Step 0: locate createdump so Step 2 can invoke it by absolute path.
# ------------------------------------------------------------------
$dotnetRoot = Split-Path -Parent $dotnet.Source
$createdump = Get-ChildItem -Path $dotnetRoot -Recurse -Filter 'createdump.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $createdump) {
    Log 'WARN: createdump.exe not found under the dotnet root; Step 2 skipped.'
} else {
    Log "createdump: $($createdump.FullName)"
}

# ------------------------------------------------------------------
# Step 1: --blame-hang-dump-type none as a cheaper native-flag probe.
#
# The original L12 failure used --blame-hang-timeout with the default dump
# type (full), which invokes createdump ~0.2s in and crashes the testhost.
# Passing --blame-hang-dump-type none should make the blame collector skip
# the dump write while still exercising the hang signalling path. If the
# testhost survives this run, the failure was specific to the dump path.
# ------------------------------------------------------------------
Log ''
Log '=== Step 1: --blame-hang-dump-type none probe ==='
Log 'Running the OfficeIntegration suite with blame hangs enabled but dump type none.'
Log "This is the same command shape that crashed ~0.2s in with the default dump type."

$step1Args = @(
    'test', $Solution, '-c', $Configuration, '--no-build', '--no-restore',
    '--blame-hang',
    '--blame-hang-timeout', "${Step1TimeoutSeconds}s",
    '--blame-hang-dump-type', 'none',
    '--filter', 'Category=OfficeIntegration',
    '--logger', 'trx;LogFileName=office-step1.trx'
)

$step1Proc = Start-Process -FilePath 'dotnet' -ArgumentList $step1Args -NoNewWindow -PassThru
$step1Watchdog = [System.Diagnostics.Stopwatch]::StartNew()
$step1ExitCode = $null
$step1Crashed = $false

# Poll the process. If it dies in under 1s we treat that as the crash pattern
# rather than a normal early exit, because a healthy OfficeIntegration run
# does not complete that fast on this host.
while (-not $step1Proc.HasExited -and $step1Watchdog.Elapsed.TotalSeconds -lt $Step1DeadlineSeconds) {
    Start-Sleep -Seconds 1
}

if ($step1Proc.HasExited) {
    $step1ExitCode = $step1Proc.ExitCode
    $age = $step1Watchdog.Elapsed.TotalSeconds
    if ($age -lt 1.0) {
        $step1Crashed = $true
        Log "Step 1: testhost exited ${age}s after start with exit code $step1ExitCode."
        Log '  Interpretation: crash pattern (under 1s). The dump-type-none probe did NOT avoid the failure.'
    } else {
        Log "Step 1: testhost exited after ${age}s with exit code $step1ExitCode."
        if ($step1ExitCode -eq 0) {
            Log '  Interpretation: PASS. The blame hang signalling survived, only the dump write was failing.'
        } else {
            Log '  Interpretation: testhost survived the dump path but the test run itself failed (exit non-zero).'
            Log '  Distinguish from the original L12 failure: this is a test result, not a testhost abort.'
        }
    }
} else {
    Log "Step 1: deadline (${Step1DeadlineSeconds}s) reached; testhost still running."
    Log '  Killing the step-1 process tree and continuing to Step 2.'
    & taskkill /PID $step1Proc.Id /T /F 2>$null | Out-Null
    $step1ExitCode = 124
    $step1Crashed = $true
    Log '  Interpretation: timed out rather than crashed. The dump-type-none probe may have changed behaviour (the process is still alive past 0.2s).'
}

Log "Step 1 result: exit=$step1ExitCode crashed=$step1Crashed"

# ------------------------------------------------------------------
# Step 2: createdump debug-privilege check.
#
# createdump needs SeDebugPrivilege to open a remote process and write its
# memory. If an antivirus / EDR block is in force, createdump will fail fast
# with an access-denied pattern even against a process it owns. We spawn a
# trivial long-lived process we control, point createdump at it, and inspect
# the outcome.
#
# This is a host-configuration probe, not a code gate. A host where Step 1
# passes and Step 2 also passes is not the L12 host. A host where Step 1
# crashes and Step 2 fails with an access-denied pattern is consistent with
# EDR blocking createdump's debug privilege.
# ------------------------------------------------------------------
Log ''
Log '=== Step 2: createdump debug-privilege probe ==='

$step2Result = @{
    CreatedumpFound = [bool]$createdump
    Outcome         = ''
    ExitCode        = $null
    StdOut          = ''
    StdErr          = ''
    ElapsedSeconds  = 0
}

if (-not $createdump) {
    Log 'Step 2: skipped (createdump not found).'
} else {
    # Spawn a trivial long-lived helper we can point createdump at.
    $helperDir = Join-Path ([System.IO.Path]::GetTempPath()) ('l12-step2-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $helperDir -Force | Out-Null
    $helperScript = Join-Path $helperDir 'sleep.ps1'
    @'
Start-Sleep -Seconds 600
'@
    Set-Content -LiteralPath $helperScript -Value @'
Start-Sleep -Seconds 600
'@ -Encoding utf8NoBOM

    $helperProc = Start-Process -FilePath 'pwsh' -ArgumentList @('-NoProfile', '-File', $helperScript) -NoNewWindow -PassThru
    Start-Sleep -Milliseconds 500

    if (-not $helperProc.HasExited) {
        $helperPid = $helperProc.Id
        Log "Step 2: spawned helper PID $helperPid (pwsh sleeping 600s)."

        $dumpDir = Join-Path $evidence 'step2-dump'
        if (-not (Test-Path -LiteralPath $dumpDir)) { New-Item -ItemType Directory -Path $dumpDir -Force | Out-Null }
        $dumpPath = Join-Path $dumpDir "pid-${helperPid}.dmp"

        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $proc = Start-Process -FilePath $createdump.FullName -ArgumentList @($helperPid, '-o', $dumpPath) -NoNewWindow -PassThru -Wait -RedirectStandardOutput $env:TEMP Step2-out.tmp -RedirectStandardError $env:TEMP Step2-err.tmp
        $sw.Stop()

        $step2Result.ExitCode = $proc.ExitCode
        $step2Result.StdOut  = (Get-Content -LiteralPath "$env:TEMP Step2-out.tmp" -Raw -ErrorAction SilentlyContinue)
        $step2Result.StdErr  = (Get-Content -LiteralPath "$env:TEMP Step2-err.tmp" -Raw -ErrorAction SilentlyContinue)
        $step2Result.ElapsedSeconds = $sw.Elapsed.TotalSeconds
        Remove-Item -LiteralPath "$env:TEMP Step2-out.tmp" -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath "$env:TEMP Step2-err.tmp" -Force -ErrorAction SilentlyContinue

        Log "Step 2: createdump exited ${sw.Elapsed.TotalSeconds}s with code $($proc.ExitCode)."

        if (Test-Path -LiteralPath $dumpPath) {
            $dumpBytes = (Get-Item -LiteralPath $dumpPath).Length
            Log "Step 2: dump written: $dumpPath ($([math]::Round($dumpBytes / 1MB, 2)) MB)."
            $step2Result.Outcome = 'dump-written'
        } elseif ($proc.ExitCode -ne 0) {
            $err = $step2Result.StdErr
            if ($err -match 'access denied|AccessDenied|STATUS_ACCESS_DENIED|debug privilege|SeDebug') {
                Log 'Step 2: createdump failed with an access-denied / debug-privilege pattern.'
                Log "  stderr: $err"
                $step2Result.Outcome = 'access-denied'
            } else {
                Log "Step 2: createdump failed for another reason (exit $($proc.ExitCode))."
                Log "  stderr: $err"
                $step2Result.Outcome = 'other-failure'
            }
        } else {
            Log 'Step 2: createdump exited 0 but no dump appeared.'
            $step2Result.Outcome = 'no-dump-exit-zero'
        }
    } else {
        Log 'Step 2: helper process exited before createdump could run; skipped.'
        $step2Result.Outcome = 'helper-exited-early'
    }

    # Tear down the helper. taskkill /T is the safe choice here because the
    # helper may itself have spawned child processes.
    if ($helperProc -and -not $helperProc.HasExited) {
        & taskkill /PID $helperProc.Id /T /F 2>$null | Out-Null
    }
    Remove-Item -LiteralPath $helperDir -Recurse -Force -ErrorAction SilentlyContinue
}

# ------------------------------------------------------------------
# Summary
# ------------------------------------------------------------------
Log ''
Log '=== L12 createdump diagnostic summary ==='
Log "Step 1 (--blame-hang-dump-type none): exited=$step1ExitCode crashed=$step1Crashed"
Log "Step 2 (createdump debug-privilege): $(if ($createdump) { $step2Result.Outcome } else { 'skipped (createdump not found)' })"
Log ''
Log 'Human interpretation:'
Log '  - If Step 1 PASS (exit 0, survived past 0.2s) and Step 2 PASS (dump written):'
Log '      the original L12 failure was the createdump dump-write path, not hang signalling.'
Log '      The --blame-hang-dump-type none alternative is viable on this host.'
Log '  - If Step 1 crashed (under 1s) AND Step 2 access-denied:'
Log '      consistent with antivirus / EDR blocking createdump''s SeDebugPrivilege.'
Log '      The cheaper native-flag alternative did not help; the next hang-detection'
Log '      implementation should treat the native blame collector as unsupported here.'
Log '  - If Step 1 crashed but Step 2 passed:'
Log '      the failure is not explained by debug-privilege blocking alone; the referenced'
Log '      upstream issues (dotnet/sdk#15555, dotnet/diagnostics#5196) remain the best'
Log '      public record. Definitive root cause may not be knowable at this tier.'
Log '  - If createdump was not found: Step 2 is inconclusive; rely on Step 1 alone.'

Log "L12 createdump diagnostic finished $(Get-Date -Format 'o')"
exit 0
