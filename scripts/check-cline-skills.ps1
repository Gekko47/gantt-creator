#requires -Version 7
<#
.SYNOPSIS
    Drift gate for the .cline/skills/ tree.

.DESCRIPTION
    Runs the sync script, then fails the gate if a fresh sync produced
    any uncommitted change under .clinerules/, .cline/skills/, or docs/.
    Wire into verify-quick.ps1 / verify.ps1 so that drift fails the
    build.

    Exit 0 on clean; exit 1 on drift.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

# 1. Re-run the sync to make sure the working tree matches the canonical
#    source. The sync is idempotent; if the tree is already in sync, it
#    produces no diff.
$scriptRoot = Split-Path -Parent $PSCommandPath
pwsh -NoProfile -File (Join-Path $scriptRoot 'sync-cline-skills.ps1')
if ($LASTEXITCODE -ne 0) {
    Write-Error 'sync-cline-skills.ps1 failed.'
    exit $LASTEXITCODE
}

# 2. Fail if any of the three views changed.
$drift = git diff --name-only -- .clinerules/ .cline/skills/ docs/
if ($drift) {
    Write-Host 'check-cline-skills: DRIFT DETECTED'
    $drift | ForEach-Object { Write-Host "  $_" }
    Write-Error ('Skill tree is out of date. Run scripts/sync-cline-skills.ps1 ' +
                 'and commit the result. See docs/clinerules/SYNC.md for the discipline.')
    exit 1
}

Write-Host 'check-cline-skills: in sync'
exit 0
