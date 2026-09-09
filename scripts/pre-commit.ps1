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
      2. STATUS.md accuracy (check-status.ps1)
      3. Markdown link sanity (check-md-links.ps1) -- staged docs only

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
$repoRoot   = Split-Path -Parent $scriptRoot

# The four checkers below read the working tree (docs/, .cline/skills/,
# AGENTS.md, .github/). Validating the working tree is only equivalent to
# validating the content that will be committed (the staged index) when the
# two are identical. If a developer has unstaged edits under any inspected
# root, the checkers could pass on a clean working tree while the staged
# content that actually gets committed is dirty — or vice versa. Rather than
# refactor every checker to read the index snapshot via `git show` (a broad
# change to four scripts with different file-reading shapes), the minimal
# correct alternative is to refuse to run while the working tree is dirty
# under the inspected roots: that guarantees working tree == index for the
# validated content, so no masking is possible.
$inspectedRoots = @('docs', '.cline/skills', 'AGENTS.md', '.github')
# Mirror check-cline-skills.ps1 Phase 1: only UNSTAGED changes and
# UNTRACKED files count as "dirty". Staged changes (working tree == index)
# are allowed so the gate is usable from the pre-commit hook, which runs
# after staging. `git status --porcelain` would also flag staged changes
# (`M `), which would make the hook block every commit touching an
# inspected root — unusable. So we use `git diff` for the working-tree
# half and `git status` only for the untracked (`??`) half.
$unstaged = git -C $repoRoot diff --name-only -- $inspectedRoots 2>$null
$untracked = git -C $repoRoot status --porcelain -- $inspectedRoots 2>$null | Where-Object { $_.StartsWith('??') }
if ($LASTEXITCODE -eq 0 -and ($unstaged -or $untracked)) {
    Write-Host '[pre-commit] Working tree has unstaged and/or untracked changes under the inspected roots:'
    if ($unstaged) { $unstaged | ForEach-Object { Write-Host "  $_" } }
    if ($untracked) { $untracked | ForEach-Object { Write-Host "  $_" } }
    Write-Host '[pre-commit] Commit blocked. Stage or stash these changes first so the gate validates exactly the content that will be committed.'
    exit 1
}

# Fast checks. Each is idempotent and cheap (<3s). Fail fast on the first
# issue so the developer sees one clear reason instead of a cascade.
$checks = @(
    @{ Name = 'skill tree in sync';  Script = 'check-cline-skills.ps1' }
    @{ Name = 'status accuracy';     Script = 'check-status.ps1' }
    @{ Name = 'markdown link sanity'; Script = 'check-md-links.ps1' }
)

foreach ($check in $checks) {
    Write-Host "[pre-commit] $($check.Name)..."
    # Run the checker inside try/catch so a terminating Write-Error from
    # the child (e.g. check-cline-skills.ps1 on a dirty working tree) does
    # not abort the parent before we can emit the blocking message. We
    # also need to capture $LASTEXITCODE BEFORE any other native command
    # runs: piping the call through ForEach-Object / Write-Host would
    # otherwise reset $LASTEXITCODE to 0 (the last native command in the
    # pipeline is the Write-Host, which always succeeds). The two bugs
    # combine to silently mask checkers that fail via a terminating
    # Write-Error, which is why the fix touches both lines.
    $checkerOutput = @()
    $checkerExit   = 0
    try {
        $checkerOutput = & (Join-Path $scriptRoot $check.Script) 2>&1
        $checkerExit   = $LASTEXITCODE
    } catch {
        $checkerOutput = @($_.Exception.Message)
        $checkerExit   = 1
    }
    foreach ($line in $checkerOutput) { Write-Host "  $line" }
    if ($checkerExit -ne 0) {
        Write-Host "[pre-commit] $($check.Name) failed. Commit blocked. Run the full verification gate and fix before retrying."
        exit $checkerExit
    }
}

Write-Host 'pre-commit: quick gates PASS'
exit 0
