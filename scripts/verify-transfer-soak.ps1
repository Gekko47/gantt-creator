#requires -Version 7
<#
.SYNOPSIS
    Transfer-soak verification (R7.7). Run on the reference machine with a real
    Microsoft 365 x64 install and no user-owned Excel or PowerPoint session.

.DESCRIPTION
    Runs the 25-transfer TransferSoak category through the same fixtures, owned-PID
    manifest and external watchdog discipline as verify-office.ps1, but with its own
    measured gate:

      1. records the Office version, channel, bitness, locale and display scale
      2. arms BOTH harness env vars: the owned-PID manifest and the forced-kill log
      3. sweeps the previous run's manifest before the build (a stray holds the XLL)
      4. captures any existing EXCEL.EXE / POWERPNT.EXE processes (user-owned, never killed)
      5. builds Release with warnings as errors
      6. runs the tests with -Filter Category=TransferSoak under an external deadline
      7. on timeout, kills the owned dotnet-test tree and sweeps the manifest-recorded
         Office processes (PID plus start time; parentage only as a fallback)
      8. on failure, retains logs and a trx archive under the ignored evidence folder
      9. fails closed (exit 5) when the forced-kill log recorded no teardown, or when the
         owned-PID manifest was not written — an unmeasured handoff is not a zero signal
     10. fails (exit 3) when a stray harness-owned Office process survived the run
     11. fails (exit 4) when one teardown forced more than one exit

    D4: the routine verify-office.ps1 gate selects Category=OfficeIntegration and keeps
    its 780 s deadline; this soak is tagged TransferSoak ONLY and runs here, never through
    verify.ps1 or the every-commit non-office run (both negated filters exclude it).
#>

[CmdletBinding()]
param(
    [string]$Solution = 'GanttCreator.slnx',
    [string]$Configuration = 'Release',
    # MEASURED, never guessed. The 25-transfer body measured 1 m 56 s (116 s) on the
    # reference host on 2026-11-22: 15 attach iterations at a 2290 ms median and 10
    # fresh-create iterations at ~3.5 s, plus Excel/PowerPoint cold start. 360 s allows
    # a 3x slowdown before the gate fires on its own clock and still bounds a real hang.
    # Raise only with a fresh measurement, never to make a red run green.
    [int]$DeadlineSeconds = 360,
    # One fixture owns one Excel process, so one teardown contributes at most one forced
    # exit. The soak has a single OfficeFixture teardown; PowerPoint kills are the sweep's
    # territory, not the histogram's.
    [int]$MaxForcedExitsPerTeardown = 1
)

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $PSCommandPath
$artifacts  = Join-Path $scriptRoot '_artifacts'
$evidence   = Join-Path $artifacts 'office-evidence'
if (-not (Test-Path -LiteralPath $artifacts)) { New-Item -ItemType Directory -Path $artifacts | Out-Null }
if (-not (Test-Path -LiteralPath $evidence))  { New-Item -ItemType Directory -Path $evidence  | Out-Null }
$report = Join-Path $artifacts 'verify-transfer-soak.txt'
"" | Set-Content -LiteralPath $report

function Log { param($s) Write-Host $s; Add-Content -LiteralPath $report -Value $s }

function Save-Trx {
    $trxSource = Join-Path (Split-Path -Parent $PSScriptRoot) 'tests/GanttCreator.Office.IntegrationTests/TestResults/soak.trx'
    if (-not (Test-Path -LiteralPath $trxSource)) {
        # Expected on a timeout: the killed run never flushed its results.
        Log "WARN: no trx was produced by this run ($trxSource); nothing archived."
        return
    }
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmssfff'
    $archivePath = Join-Path $evidence "soak-$stamp.trx"
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

# Reads the fixture's owned-PID manifest into a hashtable of ProcessId -> StartTimeUtcTicks.
# Each record is one complete JSON object on its own line, so every line is parsed
# independently (parsing the whole file as one document fails on concatenated objects).
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
            Log "WARN: could not parse owned-PID manifest line: $($_.Exception.Message)"
        }
    }

    return $result
}

# A manifest PID is owned only when its recorded start time still matches the live one.
# A PID alone is not identity: Windows recycles process IDs between runs.
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

        # Primary ownership test: the fixture's own recorded PID and start time.
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

        # Fallback: Windows parentage, only when the manifest did not record the PID.
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

Log "verify-transfer-soak: started $(Get-Date -Format 'o')"
Log "Solution: $Solution"
Log "Configuration: $Configuration"
Log "Deadline: $DeadlineSeconds s (measured 2026-11-22: 116 s on the reference host)"

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

# Both harness handoffs are armed here. OfficeFixture records the EXCEL.EXE PID and
# PowerPointFixture records every POWERPNT.EXE PID it starts, each with a start time;
# the forced-kill log receives one record per OfficeFixture teardown.
$ownedPidsPath = Join-Path $evidence 'owned-office-pids.json'
$forcedKillsPath = Join-Path $evidence 'soak-forced-kills.log'

# Preflight sweep: a harness-owned Excel left by a previous run holds the packed XLL and
# breaks the build's ExcelDnaPack step before any post-test sweep could run.
$staleManifest = Read-OwnedPidManifest $ownedPidsPath
if ($staleManifest.Count -gt 0) {
    Log "preflight: sweeping $(@($staleManifest.Keys).Count) process(es) recorded by a previous run"
    $preflightKilled = Remove-HarnessOwnedOfficeProcesses -BeforeSnapshot @() -OwnedTreePids @() -FixtureManifest $staleManifest
    if ($preflightKilled -gt 0) {
        Log "preflight: cleared $preflightKilled harness-owned Office process(es) left by a previous run"
    }
}

$officeBeforeSnapshot = @()
try {
    $officeBeforeSnapshot = @(Get-Process -Name 'EXCEL','POWERPNT' -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) | Where-Object { $_ } | Sort-Object -Unique
} catch {
    $officeBeforeSnapshot = @()
}

if (Test-Path -LiteralPath $ownedPidsPath) { Remove-Item -LiteralPath $ownedPidsPath -Force -ErrorAction SilentlyContinue }
if (Test-Path -LiteralPath $forcedKillsPath) { Remove-Item -LiteralPath $forcedKillsPath -Force -ErrorAction SilentlyContinue }
$env:GANTTCREATOR_OWNED_PIDS_PATH = $ownedPidsPath
$env:GANTTCREATOR_FORCED_KILLS_PATH = $forcedKillsPath

Log 'build Release -warnaserror'
dotnet build $Solution -c $Configuration --no-restore -warnaserror
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Log 'test TransferSoak (external watchdog deadline; blame collector omitted per L12)'

# Remove any trx from a previous run BEFORE starting, so a take-over archive can never
# present an old run's results as this run's.
$staleTrx = Join-Path (Split-Path -Parent $PSScriptRoot) 'tests/GanttCreator.Office.IntegrationTests/TestResults/soak.trx'
if (Test-Path -LiteralPath $staleTrx) { Remove-Item -LiteralPath $staleTrx -Force -ErrorAction SilentlyContinue }

$testArgs = @(
    'test', $Solution, '-c', $Configuration, '--no-build', '--no-restore',
    '--filter', 'Category=TransferSoak',
    '--logger', 'trx;LogFileName=soak.trx'
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
    $ownedTreePids = @(Get-CimInstance -ClassName Win32_Process -ErrorAction SilentlyContinue |
        Where-Object { $_.ParentProcessId -eq $testProc.Id } | ForEach-Object { $_.ProcessId })
    $fixtureManifest = Read-OwnedPidManifest $ownedPidsPath

    $cleanupDeadlineSeconds = 30
    $cleanupWatchdog = [System.Diagnostics.Stopwatch]::StartNew()
    & taskkill /PID $testProc.Id /T /F 2>$null
    $taskkillExitCode = $LASTEXITCODE
    if ($taskkillExitCode -ne 0) {
        Log "WARN: taskkill failed to stop the dotnet-test process tree (exit $taskkillExitCode)"
    }

    while ($cleanupWatchdog.Elapsed.TotalSeconds -lt $cleanupDeadlineSeconds) {
        if (-not (Test-HarnessProcessTreeActive -RootProcessId $testProc.Id -KnownChildPids $ownedTreePids)) {
            break
        }
        Start-Sleep -Seconds 1
    }
    if (Test-HarnessProcessTreeActive -RootProcessId $testProc.Id -KnownChildPids $ownedTreePids) {
        Log "WARN: Cleanup failed: dotnet-test process tree for PID $($testProc.Id) is still active after ${cleanupDeadlineSeconds}s"
    }

    Remove-HarnessOwnedOfficeProcesses -BeforeSnapshot $officeBeforeSnapshot -OwnedTreePids $ownedTreePids -FixtureManifest $fixtureManifest | Out-Null
    Save-Trx
    Log "TIMEOUT: TransferSoak tests exceeded the $DeadlineSeconds s deadline and were stopped. Evidence preserved under $evidence."
    exit 124
}
Save-Trx

$finalManifest = Read-OwnedPidManifest $ownedPidsPath
$straysKilled = Remove-HarnessOwnedOfficeProcesses -BeforeSnapshot $officeBeforeSnapshot -OwnedTreePids @() -FixtureManifest $finalManifest

# The forced-exit signal. Every record is ONE OfficeFixture teardown; the soak has a
# single Excel fixture, so exactly one record is expected and none means the handoff
# broke. PowerPoint kills are the sweep's territory (PowerPointFixture writes no record).
$forcedKills = 0
$teardowns = 0
$killAttribution = [ordered]@{}
$cleanTeardowns = [System.Collections.Generic.List[string]]::new()
$quitOutcomes = [ordered]@{}
$overTeardownCeiling = @()
$rcwMeasured = 0
$rcwLeaks = @()
try {
    if (Test-Path -LiteralPath $forcedKillsPath) {
        foreach ($line in @(Get-Content -LiteralPath $forcedKillsPath -ErrorAction SilentlyContinue)) {
            $trimmed = $line.Trim()
            if ($trimmed.Length -eq 0) { continue }
            $count = 0
            $test = $null
            $quit = $null
            # -1 means the record carried no reference count; that is unmeasured, never zero.
            $rcw = -1
            if ($trimmed.StartsWith('{')) {
                try {
                    $rec = $trimmed | ConvertFrom-Json -ErrorAction Stop
                    $count = [int]$rec.Count
                    $test = [string]$rec.Test
                    if ($rec.PSObject.Properties['Quit']) { $quit = [string]$rec.Quit }
                    if ($rec.PSObject.Properties['Rcw']) { $rcw = [int]$rec.Rcw }
                } catch { continue }
            } elseif ([int]::TryParse($trimmed, [ref]$count)) {
                $test = '(unattributed legacy record)'
            } else {
                continue
            }
            $teardowns++
            $forcedKills += $count
            if ($rcw -ge 0) { $rcwMeasured++ }
            if ($rcw -gt 0) { $rcwLeaks += [pscustomobject]@{ Test = $test; Rcw = $rcw } }
            if (-not [string]::IsNullOrEmpty($quit)) {
                if (-not $quitOutcomes.Contains($quit)) { $quitOutcomes[$quit] = 0 }
                $quitOutcomes[$quit]++
            }
            if ($count -gt $MaxForcedExitsPerTeardown) {
                $overTeardownCeiling += [pscustomobject]@{ Test = $test; Count = $count }
            }
            if ($count -eq 0) {
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

Log "Owned-PID manifest record(s): $($finalManifest.Count) (Excel fixture plus every PowerPoint PID it started)"
Log "Forced-exit signal: $forcedKills forced exit(s) across $teardowns recorded teardown(s) (per-teardown ceiling $MaxForcedExitsPerTeardown, end state 0). This is Excel's residency, not a leaked proxy: the leak gate is the reference count below, and the orphan verdict is the end-state sweep."
Log "Reference-count signal: $rcwMeasured of $teardowns teardown(s) reported an application reference count, and $($rcwLeaks.Count) left one behind. Zero is the leak gate: the fixture must be the last managed holder when it tears down."
if ($killAttribution.Count -gt 0) {
    Log "Forced-exit signal by test body (descending):"
    foreach ($entry in ($killAttribution.GetEnumerator() | Sort-Object -Property Value -Descending)) {
        Log ("  {0,3}  {1}" -f $entry.Value, $entry.Key)
    }
}
if ($cleanTeardowns.Count -gt 0) {
    $cleanTeardowns.Sort([System.StringComparer]::Ordinal)
    Log "Forced-exit signal, clean teardowns:"
    foreach ($name in $cleanTeardowns) {
        Log ("  {0,3}  {1}" -f 0, $name)
    }
}
if ($quitOutcomes.Count -gt 0) {
    Log "Forced-exit signal by Quit() outcome:"
    foreach ($entry in ($quitOutcomes.GetEnumerator() | Sort-Object -Property Key)) {
        Log ("  {0,3}  {1}" -f $entry.Value, $entry.Key)
    }
}

if ($testProc.ExitCode -ne 0) {
    Log "FAIL: TransferSoak tests exited $($testProc.ExitCode). Evidence preserved under $evidence."
    exit $testProc.ExitCode
}

if ($straysKilled -gt 0) {
    Log "FAIL: $straysKilled harness-owned Office process(es) survived the test run and had to be killed. Evidence preserved under $evidence."
    exit 3
}

# Fail closed. An absent log is not a zero signal: OfficeFixture writes one record per
# teardown, so a green soak always produces exactly one.
if ($teardowns -eq 0) {
    Log "FAIL: no teardown record was found in $forcedKillsPath, so the forced-exit signal was NOT measured. An unmeasured signal is not a zero signal. Evidence preserved under $evidence."
    exit 5
}

# The same fail-closed rule for the leak signal, which the soak is best placed to catch: it is the
# one gate that performs 25 transfers with repeated staging and deletion, so a proxy left behind by
# a transfer path would show here first.
if ($rcwMeasured -eq 0) {
    Log "FAIL: no teardown record carried an application reference count, so the leak signal was NOT measured. An unmeasured signal is not a zero signal. Evidence preserved under $evidence."
    exit 5
}

# The leak gate itself. A non-zero count means the fixture was not the last managed holder.
if ($rcwLeaks.Count -gt 0) {
    foreach ($entry in $rcwLeaks) {
        Log ("  teardown {0} left {1} application reference(s)" -f $entry.Test, $entry.Rcw)
    }
    Log "FAIL: $($rcwLeaks.Count) teardown(s) left an application reference behind after releasing the proxy. Evidence preserved under $evidence."
    exit 7
}

# Fail closed on the second handoff too: a green soak starts PowerPoint hosts, so an empty
# manifest means GANTTCREATOR_OWNED_PIDS_PATH or the fixture write broke.
if ($finalManifest.Count -eq 0) {
    Log "FAIL: the owned-PID manifest is empty after a green soak, so the PID-ownership handoff was NOT measured. Evidence preserved under $evidence."
    exit 5
}

if ($overTeardownCeiling.Count -gt 0) {
    foreach ($entry in $overTeardownCeiling) {
        Log ("  teardown {0} forced {1} exit(s)" -f $entry.Test, $entry.Count)
    }
    Log "FAIL: $($overTeardownCeiling.Count) teardown(s) forced more than $MaxForcedExitsPerTeardown exit(s). One fixture owns one Excel process, so a teardown above the ceiling owned more than one process or escalated twice. Evidence preserved under $evidence."
    exit 4
}

Log "verify-transfer-soak: PASS. Report: $report"
exit 0
