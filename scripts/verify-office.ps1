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
      6. reports two per-teardown signals: how many owned Excel processes a
         teardown had to kill after Quit() returned, and how many application
         references it left behind. The reference count is the leak gate (exit 7
         when one is not zero); a forced exit is Excel's residency, gated only by
         the per-teardown ceiling (exit 4). Fails closed with exit 5 when no
         teardown record was written at all
#>

[CmdletBinding()]
param(
    [string]$Solution = 'GanttCreator.slnx',
    [string]$Configuration = 'Release',
    # Shards. The gate CAN run the OfficeIntegration set as N sequential processes, so a
    # hang costs one shard's remaining classes rather than the whole run and every shard
    # that finishes keeps its own trx. Each shard is a fresh test host, which is also what
    # lets a finished shard close cleanly instead of being taken down with the run.
    #
    # DEFAULT 1, and the reason is measured. Shards are now selected with self-contained
    # `Category=OfficeIntegration&FullyQualifiedName~Class` terms, because VSTest's filter
    # grammar has no grouping (`Category=...&(A|B|C)` ran 82 of 106 while every shard exited 0,
    # 2026-11-25) and a bare OR list matches by substring (a rehearsal split ran 143 tests
    # against the gate's 103, including the excluded TransferSoak body). The coverage check
    # now fails on either direction, so an under- or over-wide partition cannot pass silently.
    # The default still stays 1 until a multi-shard run is MEASURED to be faster: the 68-body
    # run shows a position effect (mean 2.7 s over the first half, 6.7 s over the second), but
    # nothing yet shows that splitting removes it.
    [ValidateRange(1, 8)]
    [int]$ShardCount = 1,
    # Discovery normally proves the gate ran what it claims: every listed test must run, and nothing
    # else may. When the host cannot list its tests -- a wedged Excel, a test host that refuses to
    # start, a changed `--list-tests` banner -- that proof cannot run, and a gate that cannot prove
    # what it ran is not a gate: it exits 6. This switch exists for exactly one caller, the isolated
    # Pester harness, whose fake `dotnet` prints no test names; passing it records a warning instead
    # of a proof, so the loss is visible in the log rather than silent.
    [switch]$AllowUnprovenDiscovery,
    # Run the smoke subset instead of the full OfficeIntegration set: the same tests, selected by
    # the `OfficeSmoke` trait they also carry. This is the per-branch run the full gate cannot be;
    # the full gate stays the release gate. The category variable below is the ONLY difference
    # between the two runs, which is why every filter site reads it instead of a literal.
    [switch]$Smoke,
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
    # 3. The LEAK gate is the application reference count, not the forced exit.
    #    OfficeFixture records what Marshal.ReleaseComObject returned, so 0 proves
    #    the fixture was the last managed holder. Measured 2026-11-25 over 68 live
    #    teardowns: every record read 0 (22 clean, 46 with one forced kill) while
    #    46 of them still needed the kill. A forced exit is therefore Excel's own
    #    residency -- a body that touches a workbook can leave a responding,
    #    windowless Excel behind with no reference outstanding -- and a record
    #    above 0 is the only leakage this gate can actually observe. Exit 7 on a
    #    non-zero count, and exit 5 when no teardown reports one, because the same
    #    fail-closed rule applies.
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

# One category per run, derived once. Discovery, the single-shard filter and the per-shard
# conjunction terms all read these variables: a literal at discovery beside a smoke run filter
# would make the coverage proof report every non-smoke test as missing -- a false failure
# wearing the proof's own verdict.
$categoryFilter = if ($Smoke) { 'Category=OfficeSmoke' } else { 'Category=OfficeIntegration' }
$categoryLabel  = if ($Smoke) { 'OfficeSmoke' } else { 'OfficeIntegration' }
# The smoke subset's bound is the plan's 420 s; an explicit -DeadlineSeconds still wins, because
# the diagnostic runs need to widen or shorten the bound deliberately.
$effectiveDeadlineSeconds = if ($Smoke -and -not $PSBoundParameters.ContainsKey('DeadlineSeconds')) { 420 } else { $DeadlineSeconds }

function Log { param($s) Write-Host $s; Add-Content -LiteralPath $report -Value $s }

# Per-run trx retention: dotnet test overwrites office.trx on every run, so a
# failing run's results were destroyed by the next invocation and failures
# could not be attributed after the fact. Archive a timestamped copy into the
# ignored evidence directory on every verdict path (timeout, fail, pass).
function Save-Trx {
    param([string]$TrxName = 'office.trx')

    $trxSource = Join-Path (Split-Path -Parent $PSScriptRoot) "tests/GanttCreator.Office.IntegrationTests/TestResults/$TrxName"
    if (-not (Test-Path -LiteralPath $trxSource)) {
        # Expected on a timeout: the killed run never flushed its results. Saying so is the
        # point - the alternative was archiving whatever an earlier run had left behind.
        Log "WARN: no trx was produced by this run ($trxSource); nothing archived."
        return
    }
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmssfff'
    $archivePath = Join-Path $evidence ("office-{0}-{1}.trx" -f $stamp, [System.IO.Path]::GetFileNameWithoutExtension($TrxName))
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
Log "Deadline: $effectiveDeadlineSeconds s"

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

Log "test $categoryLabel in $ShardCount shard(s) (per-shard external watchdog; blame collector omitted per L12)"

# Discover the gate's tests from the test host before running them, so the shard
# partition below covers every class the trait selects. A class added later is inside
# a shard by construction; nothing has to remember to add it to a table here.
$listOutput = & dotnet test $Solution -c $Configuration --no-build --no-restore --filter $categoryFilter --list-tests 2>&1
# Anchor the parse on an indented, full test name. A loose "contains the namespace" match
# also captured the `Test run for ... (.NETCoreApp,Version=v10.0)` banner, which became a
# fake "class", shifted every modulo assignment and produced a filter the runner rejected
# (`~NETCoreApp,Version=v10`), leaving one shard empty while the gate still passed.
#
# The optional `(arguments)` suffix is part of the name, not noise: a [Theory] is listed ONLY
# as one row per data set, so requiring the line to end after the method dropped every Theory
# from the discovered set. Measured 2026-11-25: the three
# `Initialise_preserves_a_cell_empty_sheet_with_live_non_cell_state(artefactName: ...)` rows
# were invisible, and the over-coverage check then reported that method as unlisted.
$listedTests = @($listOutput | ForEach-Object {
    $match = [regex]::Match(
        ([string]$_),
        '^\s+(GanttCreator\.Office\.IntegrationTests\.[A-Za-z0-9_.]+)(?:\(.*\))?\s*$')
    if ($match.Success) { $match.Groups[1].Value }
})
# Rows, not methods: a Theory contributes one entry per data set. The counts and verdicts use
# the distinct method names, so the report never claims more tests than the suite has.
$listedTestNames = @($listedTests | Sort-Object -Unique)
$gateClasses = @()
$effectiveShardCount = 1
if ($listedTests.Count -eq 0) {
    # Fail closed. The partition and the both-directions coverage proof both depend on this list,
    # so an empty list means the gate cannot prove which tests it ran. It used to degrade to one
    # unsharded pass with a warning, which reintroduced exactly the false green the coverage check
    # was built to stop (a run that reports PASS on an unproven suite). Only the Pester harness,
    # whose stub prints no test names, may opt out, and even then the loss is logged.
    if (-not $AllowUnprovenDiscovery) {
        Log "FAIL: the test host listed no OfficeIntegration tests, so the partition and the coverage proof cannot run. An unproven run is not a pass. Evidence preserved under $evidence."
        exit 6
    }

    Log 'WARN: discovery was unavailable and -AllowUnprovenDiscovery was passed; running one unsharded pass with NO coverage proof.'
} else {
    $gateClasses = @($listedTests | ForEach-Object { ($_ -split '\.')[-2] } | Sort-Object -Unique)
    $effectiveShardCount = $ShardCount
    Log "gate discovered $($listedTestNames.Count) test(s) in $($gateClasses.Count) class(es)"
}

$watchdogIntervalSeconds = 5
$dotnetExe = 'dotnet'
# A floor of 1 s, not 60: the isolated harness drives this script with a deliberately short
# -DeadlineSeconds, and a fixed floor would make its timeout cases unreachable.
$perShardDeadlineSeconds = [math]::Max(1, [int]($effectiveDeadlineSeconds / $effectiveShardCount))

# Discovery itself invokes dotnet, so anything a stub (or a previous run) wrote to the two
# signal paths before it is not this run's measurement. Reset after discovery, not only
# before the build.
Remove-Item -LiteralPath $ownedPidsPath, $forcedKillsPath -Force -ErrorAction SilentlyContinue
$shardTrxNames = @()
$shardOwnedPidPaths = @()
$shardForcedKillPaths = @()
$shardExitCodes = @()
$timedOutShards = @()
# Sum of the test processes' own wall time, so the phase breakdown below can say how
# much of it the fixtures account for and how much is the runner's own overhead.
$shardElapsedMsTotal = [long]0

for ($shard = 0; $shard -lt $effectiveShardCount; $shard++)
{
    if ($gateClasses.Count -eq 0) {
        # Unsharded fallback: the legacy single pass, with the legacy paths and the legacy
        # pre-run trx removal (a stale trx must never be archived as this run's).
        $shardClasses = @()
        $filterText = $categoryFilter
        $trxName = 'office.trx'
        $legacyTrxPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'tests/GanttCreator.Office.IntegrationTests/TestResults/office.trx'
        if (Test-Path -LiteralPath $legacyTrxPath) { Remove-Item -LiteralPath $legacyTrxPath -Force -ErrorAction SilentlyContinue }
        $shardOwnedPidPath = $ownedPidsPath
        $shardForcedKillPath = $forcedKillsPath
    } else {
        # Round-robin over the sorted class list: deterministic, and it spreads the heavy
        # classes rather than letting alphabetical order pile them into one shard.
        $shardClasses = @($gateClasses | Where-Object { ($gateClasses.IndexOf($_)) % $effectiveShardCount -eq $shard })

        if ($effectiveShardCount -eq 1) {
            # One shard is the whole gate, so use the exact category filter. An OR list of
            # every class would also pull in tests that share a class with an
            # OfficeIntegration test but carry no category themselves (measured: 117
            # listed against 102 discovered), silently widening what the gate runs.
            $filterText = $categoryFilter
        } else {
            # Every term carries its own category predicate, so the expression never needs
            # grouping: VSTest's filter grammar has none (`Category=OfficeIntegration&(A|B|C)`
            # executed 82 of 106 while every shard exited 0, 2026-11-25), but `&` binds
            # tighter than `|`, so `A&B1|A&B2` reads as `(A&B1)|(A&B2)`.
            #
            # The bare OR list this replaces matched by substring and ignored the category.
            # Measured in a rehearsal split the same day: 143 tests ran against the gate's
            # 103, because it also dragged in the 36 non-live bodies and the excluded
            # `Category=TransferSoak` body. Both directions of that mistake are now verdicts
            # in the coverage check below.
            $filterText = ($shardClasses |
                ForEach-Object { "$categoryFilter&FullyQualifiedName~$_" }) -join '|'
        }

        # Remove this shard's trx from a previous run BEFORE starting, so a take-over
        # archive can never present an old run's results as this run's.
        $trxName = "office-shard-$shard.trx"
        $trxPath = Join-Path (Split-Path -Parent $PSScriptRoot) "tests/GanttCreator.Office.IntegrationTests/TestResults/$trxName"
        if (Test-Path -LiteralPath $trxPath) { Remove-Item -LiteralPath $trxPath -Force -ErrorAction SilentlyContinue }

        $shardOwnedPidPath = Join-Path $evidence "owned-office-pids-shard-$shard.json"
        $shardForcedKillPath = Join-Path $evidence "office-forced-kills-shard-$shard.log"
        Remove-Item -LiteralPath $shardOwnedPidPath, $shardForcedKillPath -Force -ErrorAction SilentlyContinue
        $shardTrxNames += $trxName
        $shardOwnedPidPaths += $shardOwnedPidPath
        $shardForcedKillPaths += $shardForcedKillPath
    }

    # Each shard writes its own manifest and kill log; the merge after the loop points
    # the existing verdict block at one file per signal.
    $env:GANTTCREATOR_OWNED_PIDS_PATH = $shardOwnedPidPath
    $env:GANTTCREATOR_FORCED_KILLS_PATH = $shardForcedKillPath

    Log "shard ${shard}: $($shardClasses.Count) class(es), deadline $perShardDeadlineSeconds s"
    $testArgs = @(
        'test', $Solution, '-c', $Configuration, '--no-build', '--no-restore',
        '--filter', $filterText,
        '--logger', "trx;LogFileName=$trxName"
    )
    $testProc = Start-Process -FilePath $dotnetExe -ArgumentList $testArgs -NoNewWindow -PassThru
    $watchdog = [System.Diagnostics.Stopwatch]::StartNew()
    while (-not $testProc.HasExited -and $watchdog.Elapsed.TotalSeconds -lt $perShardDeadlineSeconds)
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
        $fixtureManifest = Read-OwnedPidManifest $shardOwnedPidPath

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
        Save-Trx -TrxName $trxName
        Log "TIMEOUT: shard ${shard} exceeded its $perShardDeadlineSeconds s deadline and was stopped; the other shards' results are kept. Evidence preserved under $evidence."
        $timedOutShards += $shard
        continue
    }

    Save-Trx -TrxName $trxName
    # The duration is logged because shard count is a performance decision: a one-shard run and a
    # two-shard run are only comparable if each shard's own wall time is in the report.
    Log ("shard {0}: exit {1} after {2:N1} s" -f $shard, $testProc.ExitCode, $watchdog.Elapsed.TotalSeconds)
    $shardElapsedMsTotal += [long]$watchdog.Elapsed.TotalMilliseconds
    $shardExitCodes += [int]$testProc.ExitCode
}

# Coverage proof, in BOTH directions. A malformed filter can silently run fewer tests while
# every shard still exits 0 (measured 2026-11-25: 82 of 106 ran and the gate passed), and a
# substring OR list can run MORE than discovery listed (same day: 143 against 103, including
# the excluded TransferSoak body). Either way the gate would be judging a different suite than
# the one it reports, so both are fails. Names are compared by method, so a [Theory]'s cases
# do not read as missing.
if ($gateClasses.Count -gt 0) {
    $executedTests = @()
    foreach ($shardTrxName in $shardTrxNames) {
        $shardTrxPath = Join-Path (Split-Path -Parent $PSScriptRoot) "tests/GanttCreator.Office.IntegrationTests/TestResults/$shardTrxName"
        if (-not (Test-Path -LiteralPath $shardTrxPath)) { continue }
        try {
            [xml]$shardTrx = Get-Content -LiteralPath $shardTrxPath
            $executedTests += @($shardTrx.TestRun.Results.UnitTestResult |
                ForEach-Object { (([string]$_.testName) -split '\(')[0] })
        } catch {
            Log "WARN: could not read $shardTrxName for the coverage check: $($_.Exception.Message)"
        }
    }
    # Both sides are reduced to the same canonical name (everything before the argument
    # list) BEFORE comparing. The two sources spell a [Theory] differently: --list-tests
    # lists each data row as `Method(artefactName: "Comment")`, while the trx carries the
    # same display name for the row and the bare method name for an ordinary test. Comparing
    # a stripped name against a raw list reported `Initialise_preserves_a_cell_empty_sheet_
    # with_live_non_cell_state` as "ran but was not listed" on 2026-11-25, which is exactly
    # the kind of false verdict this check exists to avoid.
    $executedNames = @($executedTests |
        Where-Object { -not [string]::IsNullOrEmpty($_) } |
        ForEach-Object { ($_ -split '\(')[0] } |
        Sort-Object -Unique)
    $listedNames = @($listedTests |
        ForEach-Object { ($_ -split '\(')[0] } |
        Sort-Object -Unique)

    $missingTests = @($listedNames | Where-Object { $executedNames -notcontains $_ })
    if ($missingTests.Count -gt 0) {
        Log "FAIL: $($missingTests.Count) discovered test(s) ran in no shard, so the partition is incomplete. First: $($missingTests[0]). Evidence preserved under $evidence."
        exit 6
    }
    $unexpectedTests = @($executedNames | Where-Object { $listedNames -notcontains $_ })
    if ($unexpectedTests.Count -gt 0) {
        Log "FAIL: $($unexpectedTests.Count) test(s) ran that discovery did not list, so the filter is wider than the gate accounts for. First: $($unexpectedTests[0]). Evidence preserved under $evidence."
        exit 6
    }
    Log "coverage: all $($listedTestNames.Count) discovered test(s) ran in a shard, and nothing else ran."
}

# Merge the per-shard signals into the two paths the verdict block already reads, so
# the shard split changes where the evidence is written, not how it is judged. The
# merged owned-PID manifest is also what the NEXT run's preflight sweeps. The unsharded
# fallback already wrote those paths, so there is nothing to merge.
if ($shardForcedKillPaths.Count -gt 0) {
    if (Test-Path -LiteralPath $forcedKillsPath) { Remove-Item -LiteralPath $forcedKillsPath -Force -ErrorAction SilentlyContinue }
    foreach ($shardLogPath in $shardForcedKillPaths)
    {
        if (Test-Path -LiteralPath $shardLogPath) { Get-Content -LiteralPath $shardLogPath | Add-Content -LiteralPath $forcedKillsPath }
    }
    if (Test-Path -LiteralPath $ownedPidsPath) { Remove-Item -LiteralPath $ownedPidsPath -Force -ErrorAction SilentlyContinue }
    foreach ($shardManifestPath in $shardOwnedPidPaths)
    {
        if (Test-Path -LiteralPath $shardManifestPath) { Get-Content -LiteralPath $shardManifestPath | Add-Content -LiteralPath $ownedPidsPath }
    }
}
$env:GANTTCREATOR_OWNED_PIDS_PATH = $ownedPidsPath
$env:GANTTCREATOR_FORCED_KILLS_PATH = $forcedKillsPath

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
$rcwMeasured = 0
$rcwLeaks = @()
$hostSumMs = 0
$hostMaxMs = 0
# Phase sums, so the report says which part of the wall time the fixtures own. Setup is
# a subset of Body, so the body term subtracts it; the graceful poll is a subset of
# teardown and is reported beside it rather than added.
$setupSumMs = [long]0
$bodyOnlySumMs = [long]0
$teardownSumMs = [long]0
$gracefulSumMs = [long]0
$phaseMeasured = 0
try {
    if (Test-Path -LiteralPath $forcedKillsPath) {
        foreach ($line in @(Get-Content -LiteralPath $forcedKillsPath -ErrorAction SilentlyContinue)) {
            $trimmed = $line.Trim()
            if ($trimmed.Length -eq 0) { continue }
            $count = 0
            $test = $null
            $quit = $null
            # -1 means the record carried no reference count (a legacy record, or a
            # teardown that never reached the release). It is unmeasured, never zero.
            $rcw = -1
            $hostMs = [long]0
            $bodyMs = [long]0
            $setupMs = -1
            $teardownMs = [long]0
            $gracefulMs = [long]0
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
                    if ($rec.PSObject.Properties['Rcw']) { $rcw = [int]$rec.Rcw }
                    if ($rec.PSObject.Properties['HostMs']) { $hostMs = [long]$rec.HostMs }
                    if ($rec.PSObject.Properties['BodyMs']) { $bodyMs = [long]$rec.BodyMs }
                    if ($rec.PSObject.Properties['SetupMs']) { $setupMs = [long]$rec.SetupMs }
                    if ($rec.PSObject.Properties['TeardownMs']) { $teardownMs = [long]$rec.TeardownMs }
                    if ($rec.PSObject.Properties['GracefulExitMs']) { $gracefulMs = [long]$rec.GracefulExitMs }
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
            if ($rcw -ge 0) { $rcwMeasured++ }
            if ($rcw -gt 0) { $rcwLeaks += [pscustomobject]@{ Test = $test; Rcw = $rcw } }
            $hostSumMs += $hostMs
            if ($hostMs -gt $hostMaxMs) { $hostMaxMs = $hostMs }
            # Setup 0 is a real measurement (a fixture that never initialised), so the
            # discriminator is the field's presence, not its value: -1 means a record
            # from before this split.
            if ($setupMs -ge 0) {
                $phaseMeasured++
                $setupSumMs += $setupMs
                $bodyOnlySumMs += [math]::Max(0, $bodyMs - $setupMs)
                $teardownSumMs += $teardownMs
                $gracefulSumMs += $gracefulMs
            }
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
# The killed fraction is computed as a whole percent with no culture formatting,
# so the report text is identical on every host locale.
$killPercent = if ($teardowns -gt 0) { [int][math]::Round(100.0 * $forcedKills / $teardowns) } else { 0 }
Log "Forced-exit signal: $forcedKills forced exit(s) across $teardowns recorded teardown(s) ($killPercent% of teardowns killed; per-teardown ceiling $MaxForcedExitsPerTeardown). This is Excel's residency, not a leaked proxy: the calibration run of 2026-11-25 released the application proxy to 0 references in all 68 teardowns and still killed 46 of them. The total is REPORTED, not gated: it tracks how many Office test bodies touch a workbook."
Log "Reference-count signal: $rcwMeasured of $teardowns teardown(s) reported an application reference count, and $($rcwLeaks.Count) left one behind. Zero is the leak gate: the fixture must be the last managed holder when it tears down."
# Reported, never gated: the per-body host cost, so the next question about a per-class host has a
# number in the run log instead of an estimate carried over from an earlier experiment.
$hostMeanMs = [int][math]::Round($hostSumMs / [math]::Max(1, $teardowns))
Log "Host-launch signal: $hostSumMs ms across $teardowns teardown(s), max $hostMaxMs ms, mean $hostMeanMs ms per body."

# Phase signal: where the wall time went. The fixture covers construction-to-ready (setup),
# the body, and teardown; what is left of the test processes' own time is the runner --
# host start-up, discovery, reporting and per-test turnaround. Reported, never gated: the
# numbers move with the suite, but the RATIO is what makes a regression localisable without
# a probe run, and a residual that grows faster than the fixtures is a runner problem, not a
# product one. -1 setup means a record from before this split, so the count that follows says
# how many records the sums cover.
$fixtureSumMs = $setupSumMs + $bodyOnlySumMs + $teardownSumMs
$residualMs = [math]::Max(0, $shardElapsedMsTotal - $fixtureSumMs)
Log "Phase signal: setup $setupSumMs ms, body $bodyOnlySumMs ms, teardown $teardownSumMs ms (of which graceful poll $gracefulSumMs ms) across $phaseMeasured of $teardowns record(s), so fixtures own $fixtureSumMs ms."
Log "Phase signal: the test processes ran $shardElapsedMsTotal ms, leaving $residualMs ms outside every fixture - host start-up, discovery, reporting and per-test turnaround."

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

if ($timedOutShards.Count -gt 0) {
    Log "FAIL: shard(s) $($timedOutShards -join ', ') exceeded their deadline; the shards that finished are archived. Evidence preserved under $evidence."
    exit 124
}

$failedShards = @($shardExitCodes | Where-Object { $_ -ne 0 })
if ($failedShards.Count -gt 0) {
    $firstFailedExitCode = [int]$failedShards[0]
    Log "FAIL: $($failedShards.Count) of $($shardExitCodes.Count) shard(s) exited non-zero (first $firstFailedExitCode). Evidence preserved under $evidence."
    exit $firstFailedExitCode
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

# The same fail-closed rule for the leak signal. A build whose fixture stopped
# writing the field would otherwise read as "no leaked references", which is the
# defect exit 5 exists to prevent.
if ($rcwMeasured -eq 0) {
    Log "FAIL: no teardown record carried an application reference count, so the leak signal was NOT measured. An unmeasured signal is not a zero signal. Evidence preserved under $evidence."
    exit 5
}

# The leak gate. A non-zero count means the fixture was not the last managed holder
# of the Application proxy when it tore down -- a real defect in the body or in the
# teardown order, unlike the forced exit above, which the host controls.
if ($rcwLeaks.Count -gt 0) {
    foreach ($entry in $rcwLeaks) {
        Log ("  teardown {0} left {1} application reference(s)" -f $entry.Test, $entry.Rcw)
    }
    Log "FAIL: $($rcwLeaks.Count) teardown(s) left an application reference behind after releasing the proxy. Evidence preserved under $evidence."
    exit 7
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
