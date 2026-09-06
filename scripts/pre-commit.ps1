#requires -Version 7
<##
.SYNOPSIS
    Lightweight pre-commit gate for the fast, deterministic checks only.

.DESCRIPTION
    Installed as the git pre-commit hook by scripts/install-pre-commit.ps1.
    Runs the fast checks that catch drift and formatting regressions without
    paying the ~60s cost of a full verify-quick on every commit.

    The hook is intentionally NOT the full verify-quick.ps1. Full
    verification stays a developer-initiated step before commit/PR per
    docs/05-GIT-QUALITY.md. The hook catches:

      1. skill-tree drift (check-cline-skills.ps1: working tree must
         be clean under docs/ and .cline/skills/)
      2. SKILL.md canonical-phrase presence (check-skill-summary.ps1)
      3. STATUS.md accuracy (check-status.ps1)
      4. Markdown link sanity (check-md-links.ps1) -- staged docs only

    Exit 0 allows the commit; exit 1 blocks it. The commit is not
    aborted mid-flight; this script runs BEFORE the commit object is
    created, so a failure leaves the working tree intact.

    Install: scripts/install-pre-commit.ps1
    Remove:  git config --unset core.hooksPath
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $PSCommandPath

# Fast checks. Each is idempotent and cheap (<3s). Fail fast on the first
# issue so the developer sees one clear reason instead of a cascade.
$checks = @(
    @{ Name = 'skill tree in sync (Phase 2)';  Script = 'check-cline-skills.ps1' }
    @{ Name = 'skill summary phrases';         Script = 'check-skill-summary.ps1' }
    @{ Name = 'status accuracy';               Script = 'check-status.ps1' }
    @{ Name = 'markdown link sanity';          Script = 'check-md-links.ps1' }
)

foreach ($check in $checks) {
    Write-Host "[pre-commit] $($check.Name)..."
    & (Join-Path $scriptRoot $check.Script) 2>&1 | ForEach-Object { Write-Host "  $_" }
    if ($LASTEXITCODE -ne 0) {
        Write-Error "[pre-commit] $($check.Name) failed. Commit blocked. Run the full verification gate and fix before retrying."
        exit $LASTEXITCODE
    }
}

Write-Host 'pre-commit: quick gates PASS'
exit 0
