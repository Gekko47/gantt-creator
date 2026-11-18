#requires -Version 7
<#
.SYNOPSIS
    Office-integration verification. Run on a self-hosted runner with a
    real Microsoft 365 x64 install (per docs/adr/0001).

.DESCRIPTION
    Only the OfficeIntegration xUnit trait is selected. This is the suite
    that requires Excel / PowerPoint / clipboard and is run at phase
    exit and before release. The script:
      1. records the Office version, channel, bitness, locale, and display scale
      2. captures any existing EXCEL.EXE / POWERPNT.EXE processes (user-owned, not killed)
      3. runs the OfficeIntegration tests with a deadline
      4. on timeout, kills the owned dotnet-test tree and any harness-owned Office
         processes: primarily those whose PID OfficeFixture recorded in the
         owned-PID manifest, with Windows parentage as a fallback; never every
         Office process that appeared during the run
      5. on failure, retains logs and a screenshot manifest under
         scripts/_artifacts/office-evidence/
      6. reports the forced-exit signal per teardown -- how many owned Excel
         processes a teardown had to kill after Quit() returned -- fails with
         exit 4 when a teardown forced more than one exit, and fails closed with
         exit 5 when no teardown record was written at all
#>

[CmdletBinding()]
param(
    [string]$Solution = 'GanttCreator.slnx',
    [string]$Configuration = 'Release',
    # 660 rather than 600: the suite completed in 9 m 59 s on 2026-11-09, i.e. one second
    # inside the old default, so the gate was failing on its own deadline rather than on a
    # defect. 660 keeps a real hang bounded and gives the suite room to finish.
    #
    # RAISED 660 -> 780 ON 2026-11-14, again on a measurement rather than a guess. The suite
    # measured 10 m 55 s on the same tree the day before (one second inside 660), and the ONE
    # test added for the column-visibility round trip pushed it past the deadline: the run
    # timed out at 660 s with no trx produced, which is the gate failing on its own clock and
    # not on a defect. 780 leaves just over a minute of headroom over 11 m 10 s. The cost of
    # the larger bound is stated plainly: a real hang now takes two minutes longer to surface.
    # Raise it again only with a fresh measurement, never to make a red run green.
    [int]$DeadlineSeconds = 780,
    # Ratchet on the forced-exit signal. What the number MEANS was corrected on
    # 2026-11-18 by a body bisect, and what is GATED was corrected earlier the same
    # day from the measurement rather than from the total:
    #
    # 0. A forced exit is NOT evidence of an unreleased proxy. A probe that varied
    #    one operation per body (one Excel per fixture, the same record this gate
    #    reads) found that bodies which release nothing exit cleanly, that a body
    #    which only walks the Sheets -> ListObjects indexer chain still needs a
    #    kill, and that the surviving Excel is idle and RESPONDING (windowless,
    #    flat CPU) yet never exits after Quit() returns -- it was still alive at
    #    30 s with the grace widened. There is also a body-independent flap: a body
    #    that releases nothing needed a kill in 2 of its 7 runs. So this measures
    #    an Excel SHUTDOWN RACE, with an explicit Marshal.ReleaseComObject from the
    #    test body as a deterministic trigger, and it is named accordingly. R5.18n
    #    D4 is closed on that finding: no adapter change is justified by it, and
    #    the received "unreleased proxies keep Excel alive" story is not what these
    #    breaks are.
    #
    # 1. The verdict is the PER-TEARDOWN invariant, not an absolute run total.
    #    OfficeFixture owns one Excel per fixture and force-kills it once when
    #    Quit() does not terminate, so a teardown contributes at most one forced
    #    exit and the run total counts TEST BODIES that hit the race, not
    #    conditions. Two live runs give the measurement: 52 teardown records on
    #    2026-11-18 (12 at 0, 40 at 1) and 51 on 2026-11-16 (12 at 0, 39 at 1),
    #    with NONE above 1 in either run. Totals track the test count -- 38 at 63
    #    bodies, 40 at 65, 67 and 69, 41 on the R5.22 tree -- which is why the
    #    retired ceiling had to be raised six times (24 -> 26 -> 28 -> 30 -> 35 ->
    #    37 -> 38), always to the highest value observed and therefore always with
    #    no margin. It failed for ADDING A TEST, not for adding a leak.
    #
    #    Two candidate replacement scalars were tested against the same measurement
    #    and both drift with suite composition: the total rises by one per new
    #    workbook-operating body, and the escalating fraction (40/52 = 77% on
    #    2026-11-18, 39/51 = 76% the run before) rises
    #    toward 100% as such bodies are added. The per-teardown count is the one
    #    quantity here that does not scale with the suite: one fixture owns one
    #    process, so TWO kills from one teardown means two owned processes or a
    #    second escalation path, whatever the suite holds.
    #
    # 2. The signal FAILS CLOSED. A missing or empty forced-kill log used to read as
    #    zero kills and pass; an unmeasured signal is not a zero signal, so that is
    #    now exit 5 with the reason stated.
    #
    # The run total, the recorded-teardown count and the escalating fraction are
    # still logged: they are the numbers R5.18n D4 recorded run over run, and the
    # end state is still 0. They are reported, not gated, because they cannot
    # separate a new body that hits the race from a new test. The per-test
    # attribution line is the mitigation, and the stray sweep (exit 3) is the only
    # suite-size-independent verdict here: a process still alive after the run is a
    # genuine orphan, which the shutdown race is not.
    #
    # A deliberately forced kill is excluded by OfficeFixture.SuppressLeakSignal,
    # which twelve fixtures across ten test classes set today (search
    # `SuppressLeakSignal = true`); a suppressed fixture writes no record at all.
    #
    # Two attribution caveats belong beside the number:
    #   * The per-test line SUMS a [Theory]'s cases under its method name, so
    #     `Initialise_preserves_a_cell_empty_sheet_with_live_non_cell_state` reads 3
    #     because it has three InlineData rows, one kill each. The "outlier at 3"
    #     the retired 35 -> 37 comment treated as a possible extra leaked chain was
    #     this aggregation, not a three-chain teardown, and the measurement above
    #     (no record above 1) is the proof.
    #   * The retired ceiling and its history are the evidence for why an absolute
    #     total cannot be a gate: L19 keeps that record.
    [int]$MaxForcedExitsPerTeardown = 1
)

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $PSCommandPath
$artifacts  = Join-Path $scriptRoot '_artifacts'
$evidence   = Join-Path $artifacts 'office-evidence'
if (-not (Test-Path -LiteralPath $artifacts)) { New-Item -ItemType Directory -Path $artifacts | Out-Null }
if (-not (Test-Path -LiteralPath $evidence))  { New-Item -ItemType Directory -Path $evidence  | Out-Null }
$report = Join-Path $artifacts 'verify-office.txt'
"" | Set-Content -LiteralPath $report

function Log { param($s) Write-Host $s; Add-Content -LiteralPath $report -Value $s }

# Per-run trx retention: dotnet test overwrites office.trx on every run, so a
# failing run's results were destroyed by the next invocation and failures
# could not be attributed after the fact. Archive a timestamped copy into the
# ignored evidence directory on every verdict path (timeout, fail, pass).
function Save-Trx {
    $trxSource = Join-Path (Split-Path -Parent $PSScriptRoot) 'tests/GanttCreator.Office.IntegrationTests/TestResults/office.trx'
    if (-not (Test-Path -LiteralPath $trxSource)) {
        # Expected on a timeout: the killed run never flushed its results. Saying so is the
        # point - the alternative was archiving whatever an earlier run had left behind.
        Log "WARN: no trx was produced by this run ($trxSource); nothing archived."
        return
    }
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmssfff'
    $archivePath = Join-Path $evidence "office-$stamp.trx"
    try {
        Copy-Item -LiteralPath $trxSource -Destination $archivePath -Force
        Log "Archived trx: $archivePath"
    } catch {
        Log "WARN: trx archive failed: $($_.Exception.Message)"
    }
}

function Test-HarnessProcessTreeActive {
    param(
        [Parameter(Mandatory)][int]$RootProcessId,
        [int[]]$KnownChildPids
    )

    $processes = @(Get-CimInstance -ClassName Win32_Process -ErrorAction SilentlyContinue |
        ForEach-Object {
            [pscustomobject]@{
                ProcessId     = [int]$_.ProcessId
                ParentProcessId = [int]$_.ParentProcessId
            }
        })
    if ($processes.Count -eq 0) { return $false }

    # The launcher can exit before a child, and children can have grandchildren.
    # Walk every known root and its descendants rather than checking only one
    # generation or relying on the launcher object's HasExited property.
    $activePids = @{}
    $roots = @($RootProcessId) + @($KnownChildPids | Where-Object { $_ })
    foreach ($rootPid in $roots) {
        if ($activePids.ContainsKey($rootPid)) { continue }

        $pending = @($rootPid)
        $visited = @{}
        $index = 0
        while ($index -lt $pending.Count) {
            $processId = $pending[$index]
            $index++
            if ($visited.ContainsKey($processId)) { continue }
            $visited[$processId] = $true

            $process = $processes | Where-Object { $_.ProcessId -eq $processId } | Select-Object -First 1
            if (-not $process) { continue }

            $activePids[$processId] = $true
            $children = @($processes | Where-Object { $_.ParentProcessId -eq $processId } |
                ForEach-Object { $_.ProcessId })
            $pending += $children
        }
    }

    return $activePids.Count -gt 0
}


# Reads the fixture's owned-PID manifest into a hashtable of
#   ProcessId -> StartTimeUtcTicks.
# Each record is one complete JSON object on its own line, so every line is
# parsed independently. Parsing the whole file as one JSON document fails on
# concatenated objects ("Additional text encountered after finished reading
# JSON content"), which previously emptied this signal and silently degraded
# the sweep to the parentage fallback.
function Read-OwnedPidManifest {
    param([string]$Path)

    $result = @{}
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return $result }

    foreach ($line in @(Get-Content -LiteralPath $Path -ErrorAction SilentlyContinue)) {
        if (-not $line -or $line.Trim().Length -eq 0) { continue }
        try {
            $record = $line | ConvertFrom-Json -ErrorAction Stop
            if ($null -eq $record -or $null -eq $record.ProcessId) { continue }
            $result[[int]$record.ProcessId] = [long]$record.StartTimeUtcTicks
        } catch {
            # A single unreadable line must not discard the rest of the manifest.
            Log "WARN: could not parse owned-PID manifest line: $($_.Exception.Message)"
        }
    }

    return $result
}

# Returns whether the manifest's record still identifies the live process.
# A PID alone is not identity: Windows recycles process IDs between runs, so a
# stale manifest entry can name an unrelated process. The start time recorded
# alongside the PID is the second half of the proof. A record with no start time
# (an older manifest, or an unreadable process) fails safe and is NOT owned.
function Test-ManifestProcessIdentity {
    param(
        [Parameter(Mandatory)][hashtable]$Manifest,
        [Parameter(Mandatory)][int]$ProcessId
    )

    if (-not $Manifest.ContainsKey($ProcessId)) { return $false }

    $recordedTicks = $Manifest[$ProcessId]
    if ($recordedTicks -le 0) {
        Log "Skipping Office process PID $ProcessId (manifest carries no start time, so it cannot be identified safely)"
        return $false
    }

    try {
        $live = Get-Process -Id $ProcessId -ErrorAction Stop
        $liveTicks = $live.StartTime.ToUniversalTime().Ticks
    } catch {
        Log "Skipping Office process PID $ProcessId (live start time unreadable: $($_.Exception.Message))"
        return $false
    }

    if ($liveTicks -ne $recordedTicks) {
        Log "Skipping Office process PID $ProcessId (PID was recycled: manifest start $recordedTicks, live start $liveTicks)"
        return $false
    }

    return $true
}

function Remove-HarnessOwnedOfficeProcesses {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [int[]]$BeforeSnapshot,
        [int[]]$OwnedTreePids,
        [Parameter(Mandatory)][hashtable]$FixtureManifest
    )

    $beforeSet = @{}
    foreach ($processId in $BeforeSnapshot) { $beforeSet[$processId] = $true }

    $ownedSet = @{}
    foreach ($processId in $OwnedTreePids) { $ownedSet[$processId] = $true }

    $currentPids = @()
    try {
        $currentPids = @(Get-Process -Name 'EXCEL','POWERPNT' -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) | Where-Object { $_ } | Sort-Object -Unique
    } catch {
        Log "WARN: Could not enumerate Office processes for sweep: $($_.Exception.Message)"
        return 0
    }

    $killed = 0
    foreach ($processId in $currentPids) {
        if ($beforeSet.ContainsKey($processId)) { continue }

        # Primary ownership test: the fixture's own recorded PID *and* the start
        # time it recorded with it. This is the signal OfficeFixture wrote to the
        # owned-PID manifest, so it does not depend on Excel being a child of the
        # test host. The start-time check is what makes a stale manifest safe:
        # a recycled PID no longer matches and is skipped.
        if (Test-ManifestProcessIdentity -Manifest $FixtureManifest -ProcessId $processId) {
            Log "Killing harness-owned Office process PID $processId (recorded by OfficeFixture)"
            try {
                if ($PSCmdlet.ShouldProcess($processId, 'Kill harness-owned Office process', 'Office process sweep')) {
                    & taskkill /PID $processId /T /F 2>$null | Out-Null
                    if ($LASTEXITCODE -ne 0) {
                        Log "WARN: taskkill failed for PID $processId with exit $LASTEXITCODE"
                    } else {
                        $killed++
                    }
                }
            } catch {
                Log "WARN: taskkill threw for PID ${processId}: $($_.Exception.Message)"
            }
            continue
        }

        # Fallback ownership test: Windows parentage. Only used when the
        # fixture's own PID was not recorded; if the parent cannot be
        # determined we refuse to assume ownership -- the safe direction is
        # to leave an unrelated user-owned Office process alone.
        if (-not $FixtureManifest.ContainsKey($processId)) {
            $parentPid = $null
            try {
                $owner = Get-CimInstance -ClassName Win32_Process -Filter "ProcessId = $processId" -ErrorAction SilentlyContinue
                if ($owner) { $parentPid = $owner.ParentProcessId }
            } catch {
                $err = $_.Exception.Message
                Log "WARN: Could not determine parent of Office process PID ${processId}: $err"
            }
            if ($null -eq $parentPid -or -not $ownedSet.ContainsKey($parentPid)) {
                Log "Skipping Office process PID $processId (parent $parentPid is not in the owned dotnet-test tree)"
                continue
            }

            Log "Killing harness-owned Office process PID $processId (child of owned PID $parentPid)"
            try {
                if ($PSCmdlet.ShouldProcess($processId, 'Kill harness-owned Office process', 'Office process sweep')) {
                    & taskkill /PID $processId /T /F 2>$null | Out-Null
                    if ($LASTEXITCODE -ne 0) {
                        Log "WARN: taskkill failed for PID $processId with exit $LASTEXITCODE"
                    } else {
                        $killed++
                    }
                }
            } catch {
                Log "WARN: taskkill threw for PID ${processId}: $($_.Exception.Message)"
            }
        }
    }

    return $killed
}

Log "verify-office: started $(Get-Date -Format 'o')"
Log "Solution: $Solution"
Log "Configuration: $Configuration"
Log "Deadline: $DeadlineSeconds s"

# Record the actual Office build. Per docs/adr/0001 this is the
# self-hosted runner's Microsoft 365 install.
$cfg = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
if ($cfg) {
    Log "Office ProductReleaseIds: $($cfg.ProductReleaseIds)"
    Log "Office Platform:         $($cfg.Platform)"
    Log "Office VersionToReport:  $($cfg.VersionToReport)"
    Log "Office ClientCulture:    $($cfg.ClientCulture)"
} else {
    Log 'WARN: Office ClickToRun registry key not found. Install Microsoft 365 first.'
    exit 2
}

# The owned-PID manifest. OfficeFixture records the EXCEL.EXE process IDs it
# created, each with the start time that identifies it, so the sweep can
# positively identify the fixture's Excel even when it is not a child of the
# test host, and can tell a live owned process from a recycled PID. The file
# must not be committed; it lives under the ignored scripts/_artifacts/ tree.
$ownedPidsPath = Join-Path $evidence 'owned-office-pids.json'

# Preflight sweep, BEFORE the build. A harness-owned Excel left over from a
# previous run holds the packed XLL open, and the build's ExcelDnaPack step
# fails first ("could not be deleted. Perhaps loaded in Excel?") -- before any
# post-test sweep could run. So the previous run's manifest is consumed here to
# clear the lock, then the file is reset for this run.
#
# Identity is PID plus start time, so this can never kill a recycled PID or a
# user-owned Excel: a process this script did not record, or whose start time no
# longer matches, is skipped.
$staleManifest = Read-OwnedPidManifest $ownedPidsPath
if ($staleManifest.Count -gt 0) {
    Log "preflight: sweeping $(@($staleManifest.Keys).Count) process(es) recorded by a previous run"
    $preflightKilled = Remove-HarnessOwnedOfficeProcesses -BeforeSnapshot @() -OwnedTreePids @() -FixtureManifest $staleManifest
    if ($preflightKilled -gt 0) {
        Log "preflight: cleared $preflightKilled harness-owned Office process(es) left by a previous run"
    }
}

# Capture the Office processes that existed before this script started, so the
# sweep can distinguish harness-owned processes from user-owned ones.
$officeBeforeSnapshot = @()
try {
    $officeBeforeSnapshot = @(Get-Process -Name 'EXCEL','POWERPNT' -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) | Where-Object { $_ } | Sort-Object -Unique
} catch {
    $officeBeforeSnapshot = @()
}

if (Test-Path -LiteralPath $ownedPidsPath) { Remove-Item -LiteralPath $ownedPidsPath -Force -ErrorAction SilentlyContinue }
$env:GANTTCREATOR_OWNED_PIDS_PATH = $ownedPidsPath

# Escalation log. The fixture records how many owned processes it had to force
# to the kill path; the target is zero, and a non-zero total is the visible
# signal that a test body left a COM proxy alive. Reported, never used as the
# pass/fail verdict -- that is the stray sweep's job.
$forcedKillsPath = Join-Path $evidence 'office-forced-kills.log'
if (Test-Path -LiteralPath $forcedKillsPath) { Remove-Item -LiteralPath $forcedKillsPath -Force -ErrorAction SilentlyContinue }
$env:GANTTCREATOR_FORCED_KILLS_PATH = $forcedKillsPath

# Build, then test only the OfficeIntegration trait.
Log 'build Release -warnaserror'
dotnet build $Solution -c $Configuration --no-restore -warnaserror
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Log 'test OfficeIntegration (external watchdog deadline; blame collector omitted per L12)'

# Remove any trx from a previous run BEFORE starting, so a take-over archive can never
# present an old run's results as this run's. Measured 2026-11-09: a timed-out gate
# archived a trx dated four days earlier, and reading it produced a false diagnosis.
$staleTrx = Join-Path (Split-Path -Parent $PSScriptRoot) 'tests/GanttCreator.Office.IntegrationTests/TestResults/office.trx'
if (Test-Path -LiteralPath $staleTrx) { Remove-Item -LiteralPath $staleTrx -Force -ErrorAction SilentlyContinue }

$testArgs = @(
    'test', $Solution, '-c', $Configuration, '--no-build', '--no-restore',
    '--filter', 'Category=OfficeIntegration',
    '--logger', 'trx;LogFileName=office.trx'
)
$watchdogIntervalSeconds = 5
$dotnetExe = 'dotnet'
$testProc = Start-Process -FilePath $dotnetExe -ArgumentList $testArgs -NoNewWindow -PassThru
$watchdog = [System.Diagnostics.Stopwatch]::StartNew()
while (-not $testProc.HasExited -and $watchdog.Elapsed.TotalSeconds -lt $DeadlineSeconds)
{
    Start-Sleep -Seconds $watchdogIntervalSeconds
}
if (-not $testProc.HasExited)
{
    # Soft timeout: kill the whole dotnet-test process tree (not just the
    # launcher -- Stop-Process alone leaves the actual test-host child,
    # and any Excel it owns, running as an orphan). taskkill /T reaches
    # the tree; by PID, never by Office process name.
    # Capture the owned tree first: the sweep below needs the parent PIDs
    # that were children of this launcher, and once taskkill runs they are
    # gone. OfficeFixture launches Excel in-process, so the fixture's Excel
    # is a child of the test host, which is a child of this launcher.
    $ownedTreePids = @(Get-CimInstance -ClassName Win32_Process -ErrorAction SilentlyContinue |
        Where-Object { $_.ParentProcessId -eq $testProc.Id } | ForEach-Object { $_.ProcessId })
    # The owned-PID manifest outlives the killed test host, so read it now.
    $fixtureManifest = Read-OwnedPidManifest $ownedPidsPath

    $cleanupDeadlineSeconds = 30
    $cleanupWatchdog = [System.Diagnostics.Stopwatch]::StartNew()
    & taskkill /PID $testProc.Id /T /F 2>$null
    $taskkillExitCode = $LASTEXITCODE
    if ($taskkillExitCode -ne 0) {
        Log "WARN: taskkill failed to stop the dotnet-test process tree (exit $taskkillExitCode)"
    }

    # Give the forced tree kill time to propagate to grandchildren. If the
    # launcher or any known descendant is still present at the deadline, the
    # timeout is not a clean shutdown and must be reported as a cleanup failure.
    while ($cleanupWatchdog.Elapsed.TotalSeconds -lt $cleanupDeadlineSeconds) {
        if (-not (Test-HarnessProcessTreeActive -RootProcessId $testProc.Id -KnownChildPids $ownedTreePids)) {
            break
        }
        Start-Sleep -Seconds 1
    }
    if (Test-HarnessProcessTreeActive -RootProcessId $testProc.Id -KnownChildPids $ownedTreePids) {
        Log "WARN: Cleanup failed: dotnet-test process tree for PID $testProc.Id is still active after ${cleanupDeadlineSeconds}s"
    }

    # A genuine timeout means OfficeFixture.DisposeAsync never ran, so its
    # owned Excel can be orphaned outside the killed tree too -- re-run
    # this script's own sweep. Ownership is primary from the fixture's own
    # recorded PID, with parentage as a fallback; user-owned Office is never
    # killed.
    Remove-HarnessOwnedOfficeProcesses -BeforeSnapshot $officeBeforeSnapshot -OwnedTreePids $ownedTreePids -FixtureManifest $fixtureManifest | Out-Null
    Save-Trx
    Log "TIMEOUT: OfficeIntegration tests exceeded the $DeadlineSeconds s deadline and were stopped. Evidence preserved under $evidence."
    exit 124
}
Save-Trx

# Post-run sweep on the non-timeout paths too. A fixture-owned Excel that
# outlived its test holds the packed XLL and breaks the NEXT run's publish step,
# so the sweep is not reserved for the timeout path. Ownership is PID plus the
# recorded start time, with parentage as a fallback; user-owned Office is never
# killed.
$finalManifest = Read-OwnedPidManifest $ownedPidsPath
$straysKilled = Remove-HarnessOwnedOfficeProcesses -BeforeSnapshot $officeBeforeSnapshot -OwnedTreePids @() -FixtureManifest $finalManifest

# Report the forced-exit signal. Every record is ONE fixture teardown, so its
# Count is how many owned processes that teardown had to kill after Quit()
# returned -- one, normally. The count is NOT read as "a proxy chain leaked": the
# 2026-11-18 bisect measured idle, responding Excel processes that ignore a
# returned Quit(), including in bodies that release nothing at all.
$forcedKills = 0
$teardowns = 0
$killAttribution = [ordered]@{}
$cleanTeardowns = [System.Collections.Generic.List[string]]::new()
$quitOutcomes = [ordered]@{}
$overTeardownCeiling = @()
try {
    if (Test-Path -LiteralPath $forcedKillsPath) {
        foreach ($line in @(Get-Content -LiteralPath $forcedKillsPath -ErrorAction SilentlyContinue)) {
            $trimmed = $line.Trim()
            if ($trimmed.Length -eq 0) { continue }
            $count = 0
            $test = $null
            $quit = $null
            if ($trimmed.StartsWith('{')) {
                # One fixture teardown's contribution, with the test that caused it
                # (ADR-0035's determinism work). The total is unchanged by the
                # richer record -- attribution is added ON TOP of the sum, never
                # instead of it.
                try {
                    $rec = $trimmed | ConvertFrom-Json -ErrorAction Stop
                    $count = [int]$rec.Count
                    $test = [string]$rec.Test
                    if ($rec.PSObject.Properties['Quit']) { $quit = [string]$rec.Quit }
                } catch { continue }
            } elseif ([int]::TryParse($trimmed, [ref]$count)) {
                # A bare integer from an older evidence file: still summed, so a
                # previous run's log cannot silently reduce the reported total.
                $test = '(unattributed legacy record)'
            } else {
                continue
            }
            $teardowns++
            $forcedKills += $count
            if (-not [string]::IsNullOrEmpty($quit)) {
                if (-not $quitOutcomes.Contains($quit)) { $quitOutcomes[$quit] = 0 }
                $quitOutcomes[$quit]++
            }
            if ($count -gt $MaxForcedExitsPerTeardown) {
                $overTeardownCeiling += [pscustomobject]@{ Test = $test; Count = $count }
            }
            if ($count -eq 0) {
                # A clean teardown is part of the measurement, not noise: it is the
                # side of the race the gate has to keep, so printing both sides
                # makes a shift visible in every run without a probe being written.
                $cleanTeardowns.Add(
                    [string]::IsNullOrEmpty($test) ? '(unattributed teardown)' : $test)
            }
            if ($count -gt 0) {
                if (-not $killAttribution.Contains($test)) { $killAttribution[$test] = 0 }
                $killAttribution[$test] += $count
            }
        }
    }
} catch {
    Log "WARN: could not read the forced-kill log: $($_.Exception.Message)"
}
# The escalating fraction is computed as a whole percent with no culture
# formatting, so the report text is identical on every host locale.
$escalatingPercent = if ($teardowns -gt 0) { [int][math]::Round(100.0 * $forcedKills / $teardowns) } else { 0 }
Log "Forced-exit signal: $forcedKills forced exit(s) across $teardowns recorded teardown(s) ($escalatingPercent% escalated; per-teardown ceiling $MaxForcedExitsPerTeardown, end state 0). This measures an Excel shutdown race, not a proxy chain: a body that releases nothing can still hit it, and releasing the Sheets collection has not produced one in five runs. The total is REPORTED, not gated: it tracks how many Office test bodies exist."

# Per-test attribution. The total above is the reported number; this is WHERE it
# comes from, and it is what makes the signal actionable rather than merely
# measurable. Without it a count can only be bounded, never pinned to the body
# responsible for hitting the race.
if ($killAttribution.Count -gt 0) {
    # A [Theory]'s data rows are summed under its method name here, so a method
    # with three InlineData rows reads 3 for three one-kill teardowns.
    Log "Forced-exit signal by test body (descending; a [Theory]'s cases are summed under its method name):"
    foreach ($entry in ($killAttribution.GetEnumerator() | Sort-Object -Property Value -Descending)) {
        Log ("  {0,3}  {1}" -f $entry.Value, $entry.Key)
    }
}

# The clean side of the same measurement. Ordinal sort, so the list is identical
# on every host locale and two runs can be diffed.
if ($cleanTeardowns.Count -gt 0) {
    $cleanTeardowns.Sort([System.StringComparer]::Ordinal)
    Log "Forced-exit signal, clean teardowns (the fixture's process exited without a kill):"
    foreach ($name in $cleanTeardowns) {
        Log ("  {0,3}  {1}" -f 0, $name)
    }
}

# What Quit() did in each recorded teardown. A Quit that RETURNS while the process
# survives is the shutdown race itself, and it is a different defect from one the
# host REFUSES; that distinction lives in the run log because a probe cannot
# reproduce it on demand.
if ($quitOutcomes.Count -gt 0) {
    Log "Forced-exit signal by Quit() outcome:"
    foreach ($entry in ($quitOutcomes.GetEnumerator() | Sort-Object -Property Key)) {
        Log ("  {0,3}  {1}" -f $entry.Value, $entry.Key)
    }
}

if ($testProc.ExitCode -ne 0) {
    Log "FAIL: OfficeIntegration tests exited $($testProc.ExitCode). Evidence preserved under $evidence."
    exit $testProc.ExitCode
}

# A stray that had to be killed is a real leak, not a warning: it means a test
# left an Excel process holding the packed XLL, which is exactly the condition
# that breaks the next run. It fails the gate so the leak cannot stay invisible.
if ($straysKilled -gt 0) {
    Log "FAIL: $straysKilled harness-owned Office process(es) survived the test run and had to be killed. Evidence preserved under $evidence."
    exit 3
}

# Fail closed. An absent log is not a zero signal: OfficeFixture writes one record
# per teardown, so a successful OfficeIntegration run always produces records. None
# means the handoff broke (env var, path, or the fixture write), and the gate must
# say the signal was never measured instead of passing on a number it did not read.
if ($teardowns -eq 0) {
    Log "FAIL: no teardown record was found in $forcedKillsPath, so the forced-exit signal was NOT measured. An unmeasured signal is not a zero signal. Evidence preserved under $evidence."
    exit 5
}

# Ratchet on the forced-exit signal, per teardown. One fixture owns one Excel
# process, so a teardown that needed more than one kill owned more than one process
# or escalated twice -- a real anomaly that does not scale with how many tests the
# suite holds, which is what the retired absolute total could never be.
if ($overTeardownCeiling.Count -gt 0) {
    foreach ($entry in $overTeardownCeiling) {
        Log ("  teardown {0} forced {1} exit(s)" -f $entry.Test, $entry.Count)
    }
    Log "FAIL: $($overTeardownCeiling.Count) teardown(s) forced more than $MaxForcedExitsPerTeardown exit(s). One fixture owns one Excel process, so a teardown above the ceiling owned more than one process or escalated twice. Evidence preserved under $evidence."
    exit 4
}

Log "verify-office: PASS. Report: $report"
exit 0
