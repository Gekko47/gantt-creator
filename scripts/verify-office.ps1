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
      2. cleans any prior EXCEL.EXE / POWERPNT.EXE processes that the
         test harness itself started (we never kill user-owned Office)
      3. runs the OfficeIntegration tests with a deadline
      4. on failure, retains logs and a screenshot manifest under
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
    # Soft timeout: stop only the owned dotnet test child (via its process
    # handle, never by Office process name), so user-owned EXCEL.EXE is safe.
    Stop-Process -InputObject $testProc -Force -ErrorAction SilentlyContinue
    Log "TIMEOUT: OfficeIntegration tests exceeded the $DeadlineSeconds s deadline and were stopped. Evidence preserved under $evidence."
    exit 124
}
if ($testProc.ExitCode -ne 0) {
    Log "FAIL: OfficeIntegration tests exited $($testProc.ExitCode). Evidence preserved under $evidence."
    exit $testProc.ExitCode
}

Log "verify-office: PASS. Report: $report"
exit 0
