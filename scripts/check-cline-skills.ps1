#requires -Version 7
<#
.SYNOPSIS
    Drift gate for the .cline/skills/ tree.

.DESCRIPTION
    Fails if either view (docs/ or .cline/skills/) has uncommitted
    changes, OR if the skill content is out of date relative to the
    canonical source.

    Always-on rules live in AGENTS.md at the repository root and are
    hand-maintained; they are not checked by this gate.

    The check is two-phase:
      1. If there are uncommitted changes anywhere in the two views,
         the developer must run scripts/sync-cline-skills.ps1 and commit
         the result. We do not auto-sync here because the sync would
         silently overwrite a hand-edited skill, which the discipline
         forbids.
      2. After the working tree is clean, we re-run the sync into a
         temporary directory and byte-compare against the committed
         versions. If the regeneration produces a diff the canonical
         source has changed without the skills being refreshed.

    Exit 0 on clean; exit 1 on drift.

    See docs/clinerules/SYNC.md for the full discipline.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

# Phase 1: working tree must be clean under the two views.
$dirty = git diff --name-only -- docs/ .cline/skills/
$untracked = git status --porcelain -- docs/ .cline/skills/ | Where-Object { $_.StartsWith('??') }
if ($dirty -or $untracked) {
    Write-Host 'check-cline-skills: WORKING TREE DIRTY'
    if ($dirty) {
        Write-Host '  modified:'
        $dirty | ForEach-Object { Write-Host "    $_" }
    }
    if ($untracked) {
        Write-Host '  untracked:'
        $untracked | ForEach-Object { Write-Host "    $_" }
    }
    Write-Error ('One or more of docs/ or .cline/skills/ has uncommitted ' +
                 'changes. Either commit the regenerated skill files, or revert them.')
    exit 1
}

# Phase 2: re-run the sync into a temp directory and byte-compare
# against the committed versions. The sync is idempotent; if the
# generated content differs from the committed content, the canonical
# source has changed without the skills being regenerated.
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('cline-skills-' + [System.Guid]::NewGuid().ToString('N'))
$tmpSkills = Join-Path $tmp '.cline/skills'
try {
    pwsh -NoProfile -File (Join-Path $PSScriptRoot 'sync-cline-skills.ps1') `
        -SkillsRoot $tmpSkills 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Error "sync-cline-skills.ps1 failed with exit $LASTEXITCODE"
        exit $LASTEXITCODE
    }

    $diffs = New-Object System.Collections.Generic.List[string]
    $committed = Join-Path (Split-Path -Parent $PSScriptRoot) '.cline/skills'
    # git diff --no-index: exit 0 = no diff, exit 1 = diff, exit 128 = error.
    # -c core.autocrlf=false keeps the comparison byte-exact and avoids
    # CRLF-normalisation noise in the violation report.
    $null = git -c core.autocrlf=false diff --no-index --quiet -- "$tmpSkills" "$committed" 2>&1
    $code = $LASTEXITCODE
    if ($code -eq 1) {
        $diffOut = git -c core.autocrlf=false diff --no-index -- "$tmpSkills" "$committed" 2>&1 | Out-String
        $diffs.Add("--- drift in $committed ---")
        $diffs.Add($diffOut)
    } elseif ($code -ne 0) {
        Write-Error "git diff failed with exit $code comparing $tmpSkills vs $committed"
        exit $code
    }

    if ($diffs.Count -gt 0) {
        Write-Host 'check-cline-skills: DRIFT DETECTED between canonical source and generated views'
        $diffs | ForEach-Object { Write-Host $_ }
        Write-Error ('The skill files are out of date relative to the canonical ' +
                     'source. Run scripts/sync-cline-skills.ps1 and commit the result.')
        exit 1
    }

    Write-Host 'check-cline-skills: in sync (working tree clean; generated views match committed views)'
    exit 0
} finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}