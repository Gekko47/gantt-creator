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
#>

[CmdletBinding()]
param(
    [string]$Solution = 'GanttCreator.slnx',
    [string]$Configuration = 'Release',
    [int]$DeadlineSeconds = 600
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
        Log "WARN: trx not found at $trxSource; no per-run archive."
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
                    & taskkill /PID $processId /T /F 2>$null
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
                    & taskkill /PID $processId /T /F 2>$null
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

# Report the COM-proxy leak signal. A non-zero total means a test body left a
# proxy alive and the owned Excel only exited because the fixture killed it.
$forcedKills = 0
try {
    if (Test-Path -LiteralPath $forcedKillsPath) {
        foreach ($line in @(Get-Content -LiteralPath $forcedKillsPath -ErrorAction SilentlyContinue)) {
            $parsed = 0
            if ([int]::TryParse($line.Trim(), [ref]$parsed)) { $forcedKills += $parsed }
        }
    }
} catch {
    Log "WARN: could not read the forced-kill log: $($_.Exception.Message)"
}
Log "COM proxy leak signal: $forcedKills forced kill(s) across the run (target 0; each one is a test body that left a COM proxy alive)"

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

Log "verify-office: PASS. Report: $report"
exit 0
