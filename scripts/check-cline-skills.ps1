#requires -Version 7
<#
.SYNOPSIS
    Drift gate for the .clinerules/ and .cline/skills/ trees.

.DESCRIPTION
    Fails if any of the three views (docs/, .clinerules/, .cline/skills/)
    has uncommitted changes, OR if the rule / skill content is out of
    date relative to the canonical source.

    The check is two-phase:
      1. If there are uncommitted changes anywhere in the three views,
         the developer must run scripts/sync-cline-skills.ps1 and commit
         the result. We do not auto-sync here because the sync would
         silently overwrite a hand-edited rule, which the discipline
         forbids.
      2. After the working tree is clean, we re-run the sync into a
         temporary directory and byte-compare against the committed
         versions. If the regeneration produces a diff the canonical
         source has changed without the rules/skills being refreshed.

    Exit 0 on clean; exit 1 on drift.

    See docs/clinerules/SYNC.md for the full discipline.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

# Phase 1: working tree must be clean under the three views.
$dirty = git diff --name-only -- docs/ .clinerules/ .cline/skills/
$untracked = git status --porcelain -- docs/ .clinerules/ .cline/skills/ | Where-Object { $_.StartsWith('??') }
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
    Write-Error ('One or more of docs/, .clinerules/, .cline/skills/ has uncommitted ' +
                 'changes. Either commit the regenerated rule/skill files, or revert them.')
    exit 1
}

# Phase 2: re-run the sync into a temp directory and byte-compare
# against the committed versions. The sync is idempotent; if the
# generated content differs from the committed content, the canonical
# source has changed without the rules/skills being regenerated.
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('cline-skills-' + [System.Guid]::NewGuid().ToString('N'))
$tmpRules = Join-Path $tmp '.clinerules'
$tmpSkills = Join-Path $tmp '.cline/skills'
try {
    pwsh -NoProfile -File (Join-Path $PSScriptRoot 'sync-cline-skills.ps1') `
        -RulesRoot $tmpRules -SkillsRoot $tmpSkills 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Error "sync-cline-skills.ps1 failed with exit $LASTEXITCODE"
        exit $LASTEXITCODE
    }

    $diffs = New-Object System.Collections.Generic.List[string]
    # Compare the generated rule/skill files against the committed ones.
    $comparePaths = @(
        @{ Generated = $tmpRules; Committed = (Join-Path (Split-Path -Parent $PSScriptRoot) '.clinerules') }
        @{ Generated = $tmpSkills; Committed = (Join-Path (Split-Path -Parent $PSScriptRoot) '.cline/skills') }
    )
    foreach ($pair in $comparePaths) {
        # git diff --no-index: exit 0 = no diff, exit 1 = diff, exit 128 = error
        $null = git diff --no-index --quiet -- "$($pair.Generated)" "$($pair.Committed)" 2>&1
        $code = $LASTEXITCODE
        if ($code -eq 0) {
            continue
        }
        if ($code -ne 1) {
            Write-Error "git diff failed with exit $code comparing $($pair.Generated) vs $($pair.Committed)"
            exit $code
        }
        # Capture the diff for the violation report
        $diffOut = git diff --no-index -- "$($pair.Generated)" "$($pair.Committed)" 2>&1 | Out-String
        $diffs.Add("--- drift in $($pair.Committed) ---")
        $diffs.Add($diffOut)
    }

    if ($diffs.Count -gt 0) {
        Write-Host 'check-cline-skills: DRIFT DETECTED between canonical source and generated views'
        $diffs | ForEach-Object { Write-Host $_ }
        Write-Error ('The rule and skill files are out of date relative to the canonical ' +
                     'source. Run scripts/sync-cline-skills.ps1 and commit the result.')
        exit 1
    }

    Write-Host 'check-cline-skills: in sync (working tree clean; generated views match committed views)'
    exit 0
} finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}