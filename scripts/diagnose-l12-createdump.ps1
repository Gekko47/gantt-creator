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

    2. Parent-process dump smoke test: run createdump directly (it dumps its
       parent process) and observe whether a dump is written. The helper
       process below is an owned long-lived companion whose teardown
       discipline must hold on every path; it is not the dump target.
       If validating EDR blocking of privileged access is required,
       reproduce the actual runtime-triggered createdump path targeting a
       different process such as $helperProc instead of relying on this
       smoke test.

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
    [int]$Step1DeadlineSeconds = 120,
    [int]$Step2DeadlineSeconds = 120
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
Log "Step2 deadline:$Step2DeadlineSeconds s"

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

# Resolve repo root early: Step 0 needs it to locate the testhost runtime
# config, and Step 1's pre-flight uses it for solution and artifact paths.
$scriptRoot = Split-Path -Parent $PSCommandPath
$repoRoot = Split-Path -Parent $scriptRoot

# ------------------------------------------------------------------
# Step 0: locate createdump so Step 2 can invoke it by absolute path.
#
# Resolve the .NET runtime used by the OfficeIntegration testhost first,
# then select the createdump.exe belonging to that runtime. Taking an
# arbitrary recursive match under the dotnet root could pick a createdump
# from a different runtime than the one the testhost actually runs on.
# ------------------------------------------------------------------
$dotnetRoot = Split-Path -Parent $dotnet.Source
$createdump = $null
$step1TestHostCandidate = Get-ChildItem -Path (Join-Path $repoRoot "tests\GanttCreator.Office.IntegrationTests\bin") -Recurse -Filter "GanttCreator.Office.IntegrationTests.dll" -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -match [regex]::Escape($Configuration) } | Select-Object -First 1
if ($step1TestHostCandidate) {
    $runtimeConfig = Join-Path (Split-Path -Parent $step1TestHostCandidate.FullName) 'GanttCreator.Office.IntegrationTests.runtimeconfig.json'
    if (Test-Path -LiteralPath $runtimeConfig) {
        $rc = Get-Content -LiteralPath $runtimeConfig -Raw | ConvertFrom-Json
        # Requested runtime version: a framework-dependent runtimeconfig.json
        # carries it in runtimeOptions.framework (single framework) or
        # runtimeOptions.frameworks (multi-framework list). The previously
        # read 'Microsoft.NETCore.App.RuntimeVersion' property is not part of
        # the standard schema, so the old code silently fell through to an
        # arbitrary recursive createdump match -- exactly what this step
        # exists to avoid.
        $frameworkEntry = $null
        if ($rc.runtimeOptions.framework -and $rc.runtimeOptions.framework.version) {
            $frameworkEntry = $rc.runtimeOptions.framework
        } elseif ($rc.runtimeOptions.frameworks) {
            $frameworkEntry = @($rc.runtimeOptions.frameworks) |
                Where-Object { $_.version -and $_.name -eq 'Microsoft.NETCore.App' } |
                Select-Object -First 1
        }
        if (-not $frameworkEntry) {
            Log 'Step 0: runtimeconfig.json carries no framework/frameworks version entry; Step 2 skipped.'
        } else {
            $frameworkName = if ($frameworkEntry.name) { $frameworkEntry.name } else { 'Microsoft.NETCore.App' }
            $requestedVersion = $null
            if (-not [version]::TryParse([string]$frameworkEntry.version, [ref]$requestedVersion)) {
                Log "Step 0: cannot parse requested framework version '$($frameworkEntry.version)'; Step 2 skipped."
            } else {
                Log "Step 0: testhost targets $frameworkName $($frameworkEntry.version)"
                # Roll-forward resolution: the runtime selects the latest patch
                # of the requested minor band when one is installed, otherwise
                # the latest patch of the lowest higher minor band, within the
                # same major version. Resolving before locating createdump
                # keeps the probe pointed at the runtime the testhost actually
                # uses.
                $sharedRoot = Join-Path $dotnetRoot "shared\$frameworkName"
                $resolvedVersion = $null
                if (Test-Path -LiteralPath $sharedRoot) {
                    $installed = @(Get-ChildItem -Path $sharedRoot -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                        $v = $null
                        if ([version]::TryParse($_.Name, [ref]$v)) { $v }
                    })
                    $sameBand = @($installed | Where-Object { $_.Major -eq $requestedVersion.Major -and $_.Minor -eq $requestedVersion.Minor } | Sort-Object)
                    if ($sameBand.Count -gt 0) {
                        $resolvedVersion = $sameBand[-1]
                    } else {
                        $higherBands = @($installed |
                            Where-Object { $_.Major -gt $requestedVersion.Major -or ($_.Major -eq $requestedVersion.Major -and $_.Minor -gt $requestedVersion.Minor) } |
                            Sort-Object)
                        $lowestHigher = $higherBands | Select-Object -First 1
                        if ($lowestHigher) {
                            $resolvedVersion = $higherBands |
                                Where-Object { $_.Major -eq $lowestHigher.Major -and $_.Minor -eq $lowestHigher.Minor } |
                                Sort-Object | Select-Object -Last 1
                        }
                    }
                }
                if (-not $resolvedVersion) {
                    Log "Step 0: no installed $frameworkName runtime roll-forward-matches $($frameworkEntry.version); Step 2 skipped."
                } else {
                    $runtimeVersion = $resolvedVersion.ToString()
                    Log "Step 0: resolved installed runtime version $runtimeVersion"
                    # Find the dotnet root that hosts this runtime version
                    $sharedFramework = Join-Path $dotnetRoot "shared\$frameworkName\$runtimeVersion"
                    if (Test-Path -LiteralPath $sharedFramework) {
                        $createdump = Get-ChildItem -Path $sharedFramework -Filter 'createdump.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
                    }
                    if (-not $createdump) {
                        # Fallback: search under the dotnet root for a createdump
                        # whose runtime directory matches the resolved version
                        $createdump = Get-ChildItem -Path $dotnetRoot -Recurse -Filter 'createdump.exe' -ErrorAction SilentlyContinue |
                            Where-Object { $_.FullName -match [regex]::Escape("shared\$frameworkName\$runtimeVersion") } |
                            Select-Object -First 1
                    }
                }
            }
        }
    }
}
if (-not $createdump) {
    # No unconditional recursive pick: a createdump from a different runtime
    # than the one the testhost actually runs on would make Step 2 measure
    # the wrong binary. Framework parsing or version resolution failure
    # means Step 2 is skipped.
    Log 'WARN: createdump.exe for the testhost runtime not found; Step 2 skipped.'
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

$solutionPath = if ([System.IO.Path]::IsPathRooted($Solution)) { $Solution } else { Join-Path $repoRoot $Solution }
$step1Args = @(
    'test', $solutionPath, '-c', $Configuration, '--no-build', '--no-restore',
    '--blame-hang',
    '--blame-hang-timeout', "${Step1TimeoutSeconds}s",
    '--blame-hang-dump-type', 'none',
    '--filter', 'Category=OfficeIntegration',
    '--logger', 'trx;LogFileName=office-step1.trx'
)

# Step 1 pre-flight: validate inputs and required build artifacts before
# launching. Step 1 runs with --no-build --no-restore, so a missing solution,
# non-positive timeout, or absent built test assembly would exit fast with a
# non-zero code that must not be misclassified as a testhost abort below.
# Step 2 deadline is validated here too so a zero/negative value exits through
# this validation path instead of letting Step 2 produce a misleading
# timeout-killed result.
if (-not (Test-Path -LiteralPath $solutionPath)) {
    Log "FAIL: Step 1 solution not found: $solutionPath. Pass -Solution with the path to GanttCreator.slnx."
    exit 2
}
if ($Step1TimeoutSeconds -le 0 -or $Step1DeadlineSeconds -le 0 -or $Step2DeadlineSeconds -le 0) {
    Log 'FAIL: Step 1 timeout, Step 1 deadline, and Step 2 deadline must be positive integers.'
    exit 2
}
$step1TestDll = Get-ChildItem -Path (Join-Path $repoRoot 'tests\GanttCreator.Office.IntegrationTests\bin') -Recurse -Filter 'GanttCreator.Office.IntegrationTests.dll' -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName -match [regex]::Escape($Configuration) } | Select-Object -First 1
if (-not $step1TestDll) {
    Log "FAIL: Step 1 requires built $Configuration test artifacts (--no-build --no-restore). Build the solution first."
    exit 2
}
Log "Step 1 test assembly: $($step1TestDll.FullName)"

$step1OutTmp = Join-Path $env:TEMP 'l12-step1-out.tmp'
$step1ErrTmp = Join-Path $env:TEMP 'l12-step1-err.tmp'
# The launch, poll, and classification below run inside try/finally: on any
# terminating error or interruption after launch, the finally block still
# terminates the owned testhost process tree, instead of leaking it (the
# normal path handles cleanup via the explicit deadline branch below).
$step1Proc = $null
try {
$step1Proc = Start-Process -FilePath 'dotnet' -ArgumentList $step1Args -NoNewWindow -PassThru -RedirectStandardOutput $step1OutTmp -RedirectStandardError $step1ErrTmp
$step1Watchdog = [System.Diagnostics.Stopwatch]::StartNew()
$step1ExitCode = $null
$step1Crashed = $false
$step1Outcome = ''

# Poll the process. A fast exit alone is not a crash: missing inputs,
# invalid arguments, and no-build/no-restore failures also exit fast. The
# classification below captures command output and the exit status and only
# sets $step1Crashed when those results identify a testhost abort; missing
# inputs and invalid arguments keep distinct handling via the pre-flight
# exits above and the invalid-args outcome below.
while (-not $step1Proc.HasExited -and $step1Watchdog.Elapsed.TotalSeconds -lt $Step1DeadlineSeconds) {
    Start-Sleep -Seconds 1
}

if (-not $step1Proc.HasExited) {
    # Deadline reached: terminate the owned tree and WAIT for it before
    # touching the redirected streams. The child keeps its stdout/stderr
    # temp files open until it exits, so reading or deleting before the
    # kill can capture a partial stream and can silently fail the deletes
    # (leaking the temp files in TEMP).
    #
    # WaitForExit(5000) returns $true only when the process has actually
    # terminated. If it returns $false the tree is still alive and still
    # holds the redirected streams open; in that case the diagnostic is
    # stopped before reading or deleting those streams, so stream access
    # is preserved only after confirmed process termination. The streams
    # are left untouched on the abort path rather than risked to a partial
    # read or a silently-failed delete.
    $step1Outcome = 'timeout'
    Log "Step 1: deadline (${Step1DeadlineSeconds}s) reached; testhost still running."
    Log '  Killing the step-1 process tree and waiting for it to exit.'
    # Capture taskkill's own result rather than discarding it: taskkill /T /F
    # returns 0 only when it successfully signaled the whole tree, so a
    # non-zero exit here (process not found, access denied, a child refused
    # the signal) is itself evidence that termination was not confirmed.
    # WaitForExit(5000) alone would only tell us whether the root process
    # eventually went away, not whether the kill signal reached the tree.
    & taskkill /PID $step1Proc.Id /T /F
    $step1TaskkillExit = $LASTEXITCODE
    $step1TreeExited = $step1Proc.WaitForExit(5000)
    if ($step1TaskkillExit -ne 0 -or -not $step1TreeExited) {
        Log "FAIL: Step 1: the owned testhost process tree termination was not confirmed (taskkill exit=$step1TaskkillExit, WaitForExit=$step1TreeExited)."
        Log '  The redirected streams are left untouched (the process still has them open);'
        Log '  the diagnostic is stopping here rather than reading or deleting them.'
        exit 3
    }
    $step1ExitCode = 124
    Log '  Interpretation: timed out rather than crashed. The dump-type-none probe may have changed behaviour (the process is still alive past 0.2s).'
} else {
    $step1ExitCode = $step1Proc.ExitCode
}

# Read and preserve both redirected streams BEFORE deleting their temp
# files: once the files are gone this content is the only evidence of what
# the process printed, and both the classification below and the timeout
# failure report include it.
$step1StdOut = Get-Content -LiteralPath $step1OutTmp -Raw -ErrorAction SilentlyContinue
$step1StdErr = Get-Content -LiteralPath $step1ErrTmp -Raw -ErrorAction SilentlyContinue
$step1Output = "$step1StdOut`n$step1StdErr"
Remove-Item -LiteralPath $step1OutTmp -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $step1ErrTmp -Force -ErrorAction SilentlyContinue

if ($step1Outcome -eq 'timeout') {
    Log "  Output excerpt: $($step1Output.Substring(0, [Math]::Min(500, $step1Output.Length)))"
} elseif ($step1Proc.HasExited) {
    # Measure the real process lifetime from the process's own start/exit
    # timestamps, not the watchdog stopwatch: the stopwatch includes this
    # script's polling latency (up to the 1s sleep quantum), which could
    # misclassify a rapid crash as a later exit.
    $age = ($step1Proc.ExitTime - $step1Proc.StartTime).TotalSeconds
    # The abort signature must be the unambiguous native abort signal only.
    # The .NET CRT prints exactly "Aborted." to stderr on SIGABRT, and
    # Environment.FailFast emits "The process was aborted."; both terminate
    # the testhost and both contain the token "Aborted." (with the trailing
    # period). Matching that single token -- rather than broad terms such as
    # testhost, dump, createdump, access denied, or 0x800 -- keeps normal
    # failure handling (test results, invalid arguments, missing inputs) on
    # its own branches and sets $step1Crashed only for a genuine native
    # blame-collector abort. A bare "aborted" without the period is not used:
    # it would also match ordinary words in test names and log lines.
    # Anchor at line start/end so messages such as "Operation was aborted."
    # do not match: the accepted signatures are "Aborted." and "was aborted."
    # as complete output lines.
    $abortPattern = 'Aborted\.'
    $invalidArgsPattern = 'MSB1008|invalid argument|unrecognized|MSB1009|missing|not found|could not find'
    if ($age -lt 1.0 -and $step1ExitCode -ne 0 -and $step1Output -match "(?i)$invalidArgsPattern") {
        $step1Outcome = 'invalid-args'
        Log "Step 1: exited ${age}s after start with exit code $step1ExitCode and invalid-argument output."
        Log '  Interpretation: invalid arguments or missing inputs; not classified as a testhost crash. Fix the invocation and re-run.'
        Log "  Output excerpt: $($step1Output.Substring(0, [Math]::Min(500, $step1Output.Length)))"
    } elseif ($age -lt 1.0 -and $step1ExitCode -ne 0 -and $step1Output -match "(?i)$abortPattern") {
        $step1Crashed = $true
        $step1Outcome = 'testhost-abort'
        Log "Step 1: testhost exited ${age}s after start with exit code $step1ExitCode."
        Log '  Interpretation: crash pattern (under 1s with testhost-abort output). The dump-type-none probe did NOT avoid the failure.'
        Log "  Output excerpt: $($step1Output.Substring(0, [Math]::Min(500, $step1Output.Length)))"
    } elseif ($step1ExitCode -ne 0 -and $step1Output -match "(?i)$abortPattern") {
        # Abort signature at or after one second: the canonical ~0.2s L12
        # crash is the dominant form, but a slower abort must not be read as
        # a mere test-run failure just because it survived the first second.
        $step1Crashed = $true
        $step1Outcome = 'testhost-abort'
        Log "Step 1: testhost exited ${age}s after start with exit code $step1ExitCode."
        Log '  Interpretation: crash pattern (testhost-abort output at/after 1s). The dump-type-none probe did NOT avoid the failure.'
        Log "  Output excerpt: $($step1Output.Substring(0, [Math]::Min(500, $step1Output.Length)))"
    } elseif ($age -lt 1.0 -and $step1ExitCode -ne 0) {
        $step1Outcome = 'fast-exit-unclassified'
        Log "Step 1: exited ${age}s after start with exit code $step1ExitCode but no testhost-abort signature in output."
        Log '  Interpretation: fast exit without abort evidence; not classified as a testhost crash. Inspect the output excerpt below.'
        Log "  Output excerpt: $($step1Output.Substring(0, [Math]::Min(500, $step1Output.Length)))"
    } elseif ($age -lt 1.0) {
        $step1Outcome = 'fast-exit-zero'
        Log "Step 1: exited ${age}s after start with exit code $step1ExitCode."
        Log '  Interpretation: fast zero exit; not a testhost abort pattern.'
    } else {
        $step1Outcome = 'survived'
        Log "Step 1: testhost exited after ${age}s with exit code $step1ExitCode."
        if ($step1ExitCode -eq 0) {
            Log '  Interpretation: PASS. The blame hang signalling survived, only the dump write was failing.'
        } else {
            Log '  Interpretation: testhost survived the dump path and the test run itself failed (exit non-zero, no abort signature).'
            Log '  Distinguish from the original L12 failure: this is a test result, not a testhost abort.'
        }
    }
}
}
finally {
    # Owned-process teardown on every path: normal exit, the timeout branch
    # above, and any terminating error or interruption after launch. The
    # timeout branch usually leaves the tree dead already; the HasExited
    # guard makes a repeated kill a harmless no-op race.
    if ($step1Proc -and -not $step1Proc.HasExited) {
        Log 'Step 1: terminating the owned testhost process tree (error/interruption cleanup path).'
        & taskkill /PID $step1Proc.Id /T /F 2>$null | Out-Null
        $null = $step1Proc.WaitForExit(5000)
    }
}

Log "Step 1 result: exit=$step1ExitCode crashed=$step1Crashed outcome=$step1Outcome"

# ------------------------------------------------------------------
# Step 2: createdump parent-process dump smoke test.
#
# The installed createdump writes a dump of its parent process (it does not
# accept an arbitrary PID target), so this step only establishes whether a
# direct createdump invocation can write a dump on this host. It is not a
# debug-privilege probe against a remote process: the spawned helper below
# is an owned long-lived companion whose teardown discipline must hold on
# every path, not the dump target.
#
# This is a host-configuration probe, not a code gate. A host where Step 1
# passes and Step 2 also passes is not the L12 host. A host where Step 1
# crashes and Step 2 also fails still needs the runtime-triggered
# createdump path (targeting a different process such as $helperProc) to
# validate any EDR-blocking hypothesis.
# ------------------------------------------------------------------
Log ''
Log '=== Step 2: createdump parent-process dump smoke test ==='

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
    # Spawn a trivial long-lived helper we control. The installed createdump
    # cannot be pointed at an arbitrary PID (verified locally and in
    # dotnet/runtime src/coreclr/debug/createdump/createdumpmain.cpp:
    # "The pid argument is no longer supported"; createdump writes a dump of
    # its parent process), so this step is a diagnostic parent-process dump
    # smoke test, not a debug-privilege probe. The helper stays as an owned
    # long-lived companion process whose teardown discipline must hold on
    # every path.
    $helperDir = Join-Path ([System.IO.Path]::GetTempPath()) ('l12-step2-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $helperDir -Force | Out-Null
    $helperScript = Join-Path $helperDir 'sleep.ps1'
    Set-Content -LiteralPath $helperScript -Value @'
Start-Sleep -Seconds 600
'@ -Encoding utf8NoBOM

    try {
        $helperProc = Start-Process -FilePath 'pwsh' -ArgumentList @('-NoProfile', '-File', $helperScript) -NoNewWindow -PassThru
        Start-Sleep -Milliseconds 500

        if (-not $helperProc.HasExited) {
            $helperPid = $helperProc.Id
            Log "Step 2: spawned helper PID $helperPid (pwsh sleeping 600s)."

            $dumpDir = Join-Path $evidence 'step2-dump'
            if (-not (Test-Path -LiteralPath $dumpDir)) { New-Item -ItemType Directory -Path $dumpDir -Force | Out-Null }
            $dumpPath = Join-Path $dumpDir 'step2-parent.dmp'
            # Never reuse a stale dump: a leftover step2-parent.dmp from a
            # previous run would make a fresh failure report dump-written.
            # Remove it up front so Test-Path below only sees this run's dump.
            Remove-Item -LiteralPath $dumpPath -Force -ErrorAction SilentlyContinue

            # One valid path value per redirect parameter: the previous form
            # ('-RedirectStandardOutput $env:TEMP Step2-out.tmp') passed two
            # tokens and aborted Start-Process parameter binding before
            # createdump ever ran.
            $outTmp = Join-Path $env:TEMP 'l12-step2-out.tmp'
            $errTmp = Join-Path $env:TEMP 'l12-step2-err.tmp'

            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            # No -Wait: poll the returned process against a deadline so a hung
            # createdump cannot wedge the script. -f/--name is createdump's
            # supported dump-path option (there is no -o option); no PID
            # argument (rejected with exit -1 by the installed tool).
            $proc = Start-Process -FilePath $createdump.FullName -ArgumentList @('-f', $dumpPath) -NoNewWindow -PassThru -RedirectStandardOutput $outTmp -RedirectStandardError $errTmp

            while (-not $proc.HasExited -and $sw.Elapsed.TotalSeconds -lt $Step2DeadlineSeconds) {
                Start-Sleep -Seconds 1
            }

            $timedOut = -not $proc.HasExited
            if ($timedOut) {
                Log "Step 2: createdump did not exit within ${Step2DeadlineSeconds}s; terminating the owned process and continuing."
                Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
                $null = $proc.WaitForExit(5000)
            }
            $sw.Stop()

            $step2Result.ExitCode = if ($proc.HasExited) { $proc.ExitCode } else { $null }
            $step2Result.StdOut = Get-Content -LiteralPath $outTmp -Raw -ErrorAction SilentlyContinue
            $step2Result.StdErr = Get-Content -LiteralPath $errTmp -Raw -ErrorAction SilentlyContinue
            $step2Result.ElapsedSeconds = $sw.Elapsed.TotalSeconds
            Remove-Item -LiteralPath $outTmp -Force -ErrorAction SilentlyContinue
            Remove-Item -LiteralPath $errTmp -Force -ErrorAction SilentlyContinue

            Log "Step 2: createdump exited $($sw.Elapsed.TotalSeconds)s with code $($step2Result.ExitCode)."

            if ($timedOut) {
                Log 'Step 2: createdump hit the deadline and was terminated; the probe is inconclusive.'
                $step2Result.Outcome = 'timeout-killed'
            } elseif ($step2Result.ExitCode -eq 0 -and (Test-Path -LiteralPath $dumpPath) -and (Get-Item -LiteralPath $dumpPath).Length -gt 0) {
                # dump-written requires a zero createdump exit code AND a
                # non-empty dump file: a failed createdump can leave a
                # zero-byte or partial artifact that must not read as success.
                $dumpBytes = (Get-Item -LiteralPath $dumpPath).Length
                Log "Step 2: dump written: $dumpPath ($([math]::Round($dumpBytes / 1MB, 2)) MB)."
                $step2Result.Outcome = 'dump-written'
            } elseif ($step2Result.ExitCode -ne 0) {
                $err = $step2Result.StdErr
                if (Test-Path -LiteralPath $dumpPath) {
                    Log "Step 2: a partial dump artifact remains at $dumpPath; it is not counted as a dump."
                }
                if ($err -match 'access denied|AccessDenied|STATUS_ACCESS_DENIED|debug privilege|SeDebug') {
                    Log 'Step 2: createdump failed with an access-denied / debug-privilege pattern.'
                    Log "  stderr: $err"
                    $step2Result.Outcome = 'access-denied'
                } else {
                    Log "Step 2: createdump failed for another reason (exit $($step2Result.ExitCode))."
                    Log "  stderr: $err"
                    $step2Result.Outcome = 'other-failure'
                }
            } else {
                if (Test-Path -LiteralPath $dumpPath) {
                    Log 'Step 2: createdump exited 0 but only a zero-byte dump artifact appeared; not counted as a dump.'
                } else {
                    Log 'Step 2: createdump exited 0 but no dump appeared.'
                }
                $step2Result.Outcome = 'no-dump-exit-zero'
            }
        } else {
            Log 'Step 2: helper process exited before createdump could run; skipped.'
            $step2Result.Outcome = 'helper-exited-early'
        }
    }
    finally {
        # Tear down the helper. taskkill /T is the safe choice here because the
        # helper may itself have spawned child processes. The finally block
        # runs on terminating failures from the dump probe too, so a throw
        # from Start-Process or the classification never leaks the helper.
        if ($helperProc -and -not $helperProc.HasExited) {
            & taskkill /PID $helperProc.Id /T /F 2>$null | Out-Null
        }
        Remove-Item -LiteralPath $helperDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# ------------------------------------------------------------------
# Summary
# ------------------------------------------------------------------
Log ''
Log '=== L12 createdump diagnostic summary ==='
Log "Step 1 (--blame-hang-dump-type none): exited=$step1ExitCode crashed=$step1Crashed"
Log "Step 2 (createdump parent-process dump smoke test): $(if ($createdump) { $step2Result.Outcome } else { 'skipped (createdump not found)' })"
Log ''
Log 'Human interpretation:'
Log '  - If Step 1 PASS (exit 0, survived past 0.2s) and Step 2 PASS (dump written):'
Log '      the original L12 failure was the createdump dump-write path, not hang signalling.'
Log '      The --blame-hang-dump-type none alternative is viable on this host.'
Log '  - If Step 1 crashed (abort signature at any age) AND Step 2 failed:'
Log '      the smoke test alone cannot identify EDR blocking of privileged access;'
Log '      reproduce the runtime-triggered createdump path targeting a different'
Log '      process such as $helperProc before attributing the failure to SeDebugPrivilege.'
Log '      The cheaper native-flag alternative did not help; the next hang-detection'
Log '      implementation should treat the native blame collector as unsupported here.'
Log '  - If Step 1 crashed but Step 2 passed:'
Log '      the failure is not explained by debug-privilege blocking alone; the referenced'
Log '      upstream issues (dotnet/sdk#15555, dotnet/diagnostics#5196) remain the best'
Log '      public record. Definitive root cause may not be knowable at this tier.'
Log '  - If createdump was not found: Step 2 is inconclusive; rely on Step 1 alone.'

Log "L12 createdump diagnostic finished $(Get-Date -Format 'o')"
exit 0
