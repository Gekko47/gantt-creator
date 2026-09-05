#requires -Version 7
<#
.SYNOPSIS
    Gate on the regenerated skill summaries. Wired into verify-quick.ps1
    after sync-cline-skills.ps1 so the SKILL.md views are proven to
    contain the canonical phrases that the source documents depend on
    downstream (R0.8 note, artifact contract, culture-test policy).

.DESCRIPTION
    Each entry maps a relative path under .cline/skills/ to one or more
    phrases that must appear in the regenerated SKILL.md. Phrase presence
    is a coarse proxy for "the canonical source was not truncated by the
    summary budget" and for "the cross-reference contract is intact".

    A missing phrase is a defect: either the canonical source dropped the
    content, the sync budget cut it, or the SKILL.md file was hand-edited.
    Failures name the missing phrase and the file so the developer can
    decide which layer to fix.

    Exit 0 on clean; exit 1 with one message per violation.
#>

[CmdletBinding()]
param(
    [string]$SkillsRoot = '.cline/skills'
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$root = Join-Path $repoRoot $SkillsRoot

if (-not (Test-Path $root)) { Write-Error "Missing $root"; exit 1 }

# Maps a skill-dir/SKILL.md relative path to required canonical phrases.
# A phrase is matched as a literal substring (case-sensitive) so the
# canonical wording in the source document is what the assertion checks.
$assertions = @{
    '03-roadmap/SKILL.md'        = @('R0.8 note', 'do not skip', 'L6 and L7')
    '02-architecture/SKILL.md'   = @('additionally produces', 'coverage/')
    '04-test-strategy/SKILL.md'  = @('set `CurrentCulture`', 'ambient')
}

$violations = New-Object System.Collections.Generic.List[string]

foreach ($kv in $assertions.GetEnumerator()) {
    $path = Join-Path $root $kv.Key
    if (-not (Test-Path $path)) {
        $violations.Add("Skill summary file not found: $path")
        continue
    }
    $text = Get-Content -LiteralPath $path -Raw
    foreach ($phrase in $kv.Value) {
        if ($text -notmatch [regex]::Escape($phrase)) {
            $violations.Add("$($kv.Key) missing canonical phrase: '$phrase'")
        }
    }
}

if ($violations.Count -gt 0) {
    Write-Host "check-skill-summary: $($violations.Count) violation(s):"
    $violations | ForEach-Object { Write-Host "  - $_" }
    exit 1
}

$total = ($assertions.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
Write-Host "check-skill-summary: OK ($total phrases across $($assertions.Count) files verified)"
exit 0
