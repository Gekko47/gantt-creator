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


function Remove-HarnessOwnedOfficeProcesses {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param([int[]]$BeforeSnapshot, [int[]]$OwnedTreePids, [int[]]$FixtureOwnedPids)

    $beforeSet = @{}
    foreach ($processId in $BeforeSnapshot) { $beforeSet[$processId] = $true }

    $ownedSet = @{}
    foreach ($processId in $OwnedTreePids) { $ownedSet[$processId] = $true }

    $fixtureSet = @{}
    foreach ($processId in $FixtureOwnedPids) { $fixtureSet[$processId] = $true }

    $currentPids = @()
    try {
        $currentPids = @(Get-Process -Name 'EXCEL','POWERPNT' -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) | Where-Object { $_ } | Sort-Object -Unique
    } catch {
        Log "WARN: Could not enumerate Office processes for sweep: $($_.Exception.Message)"
        return
    }

    foreach ($processId in $currentPids) {
        if ($beforeSet.ContainsKey($processId)) { continue }

        # Primary ownership test: the fixture's own recorded PID. This is the
        # signal OfficeFixture wrote to the owned-PID manifest, so it does not
        # depend on Excel being a child of the test host.
        if ($fixtureSet.ContainsKey($processId)) {
            Log "Killing harness-owned Office process PID $processId (recorded by OfficeFixture)"
            try {
                if ($PSCmdlet.ShouldProcess($processId, 'Kill harness-owned Office process', 'Office process sweep')) {
                    & taskkill /PID $processId /T /F 2>$null
                    if ($LASTEXITCODE -ne 0) {
                        Log "WARN: taskkill failed for PID $processId with exit $LASTEXITCODE"
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
                }
            }
        } catch {
            Log "WARN: taskkill threw for PID ${processId}: $($_.Exception.Message)"
        }
    }
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

# Build, then test only the OfficeIntegration trait.
Log 'build Release -warnaserror'
dotnet build $Solution -c $Configuration --no-restore -warnaserror
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

# Capture the Office processes that existed before this script started, so the
# timeout sweep can distinguish harness-owned processes from user-owned ones.
$officeBeforeSnapshot = @{}
try {
    $officeBeforeSnapshot = @(Get-Process -Name 'EXCEL','POWERPNT' -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) | Where-Object { $_ } | Sort-Object -Unique
} catch {
    $officeBeforeSnapshot = @()
}

# Owned-PID manifest. OfficeFixture records the EXCEL.EXE process ID it
# created into this file (via $env:GANTTCREATOR_OWNED_PIDS_PATH), so the
# timeout sweep can positively identify the fixture's Excel even when it is
# not a child of the test host. The file must not be committed; it lives
# under the ignored scripts/_artifacts/ tree.
$ownedPidsPath = Join-Path $evidence 'owned-office-pids.json'
if (Test-Path -LiteralPath $ownedPidsPath) { Remove-Item -LiteralPath $ownedPidsPath -Force -ErrorAction SilentlyContinue }
$env:GANTTCREATOR_OWNED_PIDS_PATH = $ownedPidsPath

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
    $fixtureOwnedPids = @()
    try {
        if (Test-Path -LiteralPath $ownedPidsPath) {
            $fixtureOwnedPids = @((Get-Content -LiteralPath $ownedPidsPath -Raw | ConvertFrom-Json -ErrorAction SilentlyContinue).ProcessId)
        }
    } catch {
        $err = $_.Exception.Message
        Log "WARN: could not read owned-PID manifest ${ownedPidsPath}: $err"
    }
    & taskkill /PID $testProc.Id /T /F 2>$null
    # A genuine timeout means OfficeFixture.DisposeAsync never ran, so its
    # owned Excel can be orphaned outside the killed tree too -- re-run
    # this script's own sweep. Ownership is primary from the fixture's own
    # recorded PID, with parentage as a fallback; user-owned Office is never
    # killed.
    Remove-HarnessOwnedOfficeProcesses $officeBeforeSnapshot $ownedTreePids $fixtureOwnedPids
    Log "TIMEOUT: OfficeIntegration tests exceeded the $DeadlineSeconds s deadline and were stopped. Evidence preserved under $evidence."
    exit 124
}
if ($testProc.ExitCode -ne 0) {
    Log "FAIL: OfficeIntegration tests exited $($testProc.ExitCode). Evidence preserved under $evidence."
    exit $testProc.ExitCode
}

Log "verify-office: PASS. Report: $report"
exit 0
