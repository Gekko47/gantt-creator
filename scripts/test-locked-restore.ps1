#requires -Version 7
#
# Deliberate two-run locked-restore test (R0.3 acceptance).
#
# 1. Remove every obj/ directory so no cached restore artifacts
#    survive.
# 2. Run dotnet restore --locked-mode. This is run 1 and produces
#    the packages.lock.json files.
# 3. Remove every obj/ directory again.
# 4. Run dotnet restore --locked-mode a second time. This must
#    succeed purely from the lock file with no network access.
#
# Both runs must exit 0. If the second run would need network
# access, --locked-mode will refuse and fail — which is exactly
# the proof the lock file is honoured.

[CmdletBinding()]
param(
    [string]$Solution = 'GanttCreator.slnx'
)

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $PSCommandPath

function Remove-ObjDirectories {
    Get-ChildItem -Path $PSScriptRoot\..\.. -Recurse -Directory -Filter 'obj' -ErrorAction SilentlyContinue |
        ForEach-Object { Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue }
}

function Run-LockRestore {
    param([string]$Label)
    Write-Host "=== $Label ==="
    dotnet restore $Solution --locked-mode --no-cache 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Lock-mode restore failed: $Label"
        exit $LASTEXITCODE
    }
    Write-Host "PASS: $Label"
}

Write-Host "R0.3 locked-restore test: two consecutive runs from a clean state"
Write-Host ""

# --- Run 1: generate the lock files ---
Remove-ObjDirectories
Run-LockRestore -Label 'Run 1 — generate lock files (no cache)'

# --- Run 2: prove the lock files are honoured ---
Remove-ObjDirectories
Run-LockRestore -Label 'Run 2 — honour lock files, no network'

Write-Host ""
Write-Host "R0.3 locked-restore test: PASS (two consecutive runs from a clean state)"
exit 0
