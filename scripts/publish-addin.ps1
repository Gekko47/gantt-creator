#requires -Version 7
<#
.SYNOPSIS
    Publishes the AddIn and produces the packed XLL (W8 local/CI parity entry point).

.DESCRIPTION
    Single source of truth for the AddIn publish command shared by
    verify-quick.ps1 step 10, verify.ps1 step 10, and the GitHub CI
    'Publish AddIn (packed XLL artifact)' step. Inline `dotnet publish`
    in ci.yml is forbidden by scripts/ci-parity.Tests.ps1 (W8); this
    entry point is the replacement, so local and CI cannot drift.

    Before publishing, this clears Office processes that a previous
    Office-integration run left holding the packed XLL. ExcelDnaPack fails
    first with "Existing output .xll file ... could not be deleted.
    (Perhaps loaded in Excel?)", which fails the whole gate on a
    precondition the gate itself created. The sweep is keyed on the PID
    *and* the start time recorded in the previous run's owned-PID manifest,
    so a recycled PID or a user-owned Excel is never touched; anything not
    positively identified is skipped.

    Exit 0 on clean; the dotnet exit code on failure.
#>

[CmdletBinding()]
param(
    [string]$Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

function Read-OwnedPidManifest {
    <#
    .SYNOPSIS
        Reads the fixture's owned-PID manifest into ProcessId -> StartTimeUtcTicks.
    .DESCRIPTION
        Each record is one complete JSON object on its own line, so every line is
        parsed independently. Parsing the whole file as one JSON document fails on
        concatenated objects, which silently emptied this signal.
    #>
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
        }
    }

    return $result
}

# Returns whether the manifest record still identifies the live process. A PID
# alone is not identity: Windows recycles process IDs, so a stale manifest entry
# can name an unrelated process. A record with no start time fails safe and is
# NOT owned.
function Test-ManifestProcessIdentity {
    param(
        [Parameter(Mandatory)][hashtable]$Manifest,
        [Parameter(Mandatory)][int]$ProcessId
    )

    if (-not $Manifest.ContainsKey($ProcessId)) { return $false }

    $recordedTicks = $Manifest[$ProcessId]
    if ($recordedTicks -le 0) { return $false }

    try {
        $live = Get-Process -Id $ProcessId -ErrorAction Stop
        $liveTicks = $live.StartTime.ToUniversalTime().Ticks
    } catch {
        return $false
    }

    return ($liveTicks -eq $recordedTicks)
}

$ownedPidsPath = Join-Path $PSScriptRoot '_artifacts\office-evidence\owned-office-pids.json'
$staleManifest = Read-OwnedPidManifest $ownedPidsPath
if ($staleManifest.Count -gt 0) {
    $cleared = 0
    foreach ($processId in @($staleManifest.Keys)) {
        if (-not (Test-ManifestProcessIdentity -Manifest $staleManifest -ProcessId $processId)) { continue }

        $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
        if ($null -eq $process) { continue }
        if ($process.ProcessName -notin @('EXCEL', 'POWERPNT')) { continue }

        & taskkill /PID $processId /T /F 2>$null | Out-Null
        if ($LASTEXITCODE -eq 0) { $cleared++ }
    }

    if ($cleared -gt 0) {
        Write-Host "publish-addin: cleared $cleared harness-owned Office process(es) left by a previous run" -ForegroundColor Yellow
    }
}

dotnet publish src/GanttCreator.AddIn/GanttCreator.AddIn.csproj -c $Configuration --no-build
exit $LASTEXITCODE
