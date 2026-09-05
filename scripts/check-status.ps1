#requires -Version 7
<#
.SYNOPSIS
    Accuracy gate for docs/STATUS.md. Wired into verify-quick.ps1 so the
    status file cannot drift from the repository it describes.

.DESCRIPTION
    Three checks, all derived from the text of docs/STATUS.md itself:

      1. Commit hashes  — every backticked token of 7-40 lowercase hex
         characters must resolve to a commit (git rev-parse --verify).
      2. Repo paths     — every backticked token that looks like a
         repo-relative file path (contains a separator, ends in an
         extension-like suffix, no globs or URLs) must exist on disk.
      3. Roadmap IDs    — every backticked R<major>.<minor> token must
         appear in docs/03-ROADMAP.md, so the status cannot reference a
         work item the roadmap does not define.

    Globs (tokens containing *) are skipped: the R0.6 entry legitimately
    references a file that does not exist.

    Exit 0 on clean; exit 1 with one message per violation.
#>

[CmdletBinding()]
param(
    [string]$StatusPath = 'docs/STATUS.md',
    [string]$RoadmapPath = 'docs/03-ROADMAP.md'
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$statusFile = Join-Path $repoRoot $StatusPath
$roadmapFile = Join-Path $repoRoot $RoadmapPath

if (-not (Test-Path $statusFile)) { Write-Error "Missing $statusFile"; exit 1 }
if (-not (Test-Path $roadmapFile)) { Write-Error "Missing $roadmapFile"; exit 1 }

$status  = Get-Content -LiteralPath $statusFile -Raw
$roadmap = Get-Content -LiteralPath $roadmapFile -Raw

# Backticked inline tokens. The status file only references hashes, paths,
# and IDs inside backticks, so this bounds the scan to deliberate claims.
$tokens = [regex]::Matches($status, '`([^`\r\n]+)`') | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique

$violations = New-Object System.Collections.Generic.List[string]

# --- 1. Commit hashes ---
$hashPattern = '^[0-9a-f]{7,40}$'
$hashTokens = $tokens | Where-Object { $_ -match $hashPattern }
foreach ($h in $hashTokens)
{
    $null = git -C $repoRoot rev-parse --verify ("$h^{commit}") 2>$null
    if ($LASTEXITCODE -ne 0)
    {
        $violations.Add("STATUS references commit '$h' which does not resolve in git.")
    }
}

# --- 2. Repo paths ---
foreach ($t in $tokens)
{
    $isPath = $t -match '[/\\]' `
        -and $t -notmatch '[\s*(){}]' `
        -and $t -notmatch '^-' `
        -and $t -notmatch '://' `
        -and $t -match '\.[A-Za-z0-9]+$'
    if (-not $isPath) { continue }

    $candidate = Join-Path $repoRoot $t
    # Reject path traversal: the resolved path must stay inside $repoRoot.
    # GetRelativePath gives the canonical lexical difference; ".." or a
    # path that starts with "..\" means the candidate escapes the root
    # (e.g. C:\repos\gantt-creator-evil would pass a StartsWith check
    # against C:\repos\gantt-creator on Ordinal comparison).
    $resolved = [System.IO.Path]::GetFullPath($candidate)
    $repoRootFull = [System.IO.Path]::GetFullPath($repoRoot)
    $rel = [System.IO.Path]::GetRelativePath($repoRootFull, $resolved)
    if ($rel -eq '..' -or $rel.StartsWith('..\')) {
        $violations.Add("STATUS references path '$t' which resolves outside the repository.")
        continue
    }

    if (-not (Test-Path $candidate))
    {
        $violations.Add("STATUS references path '$t' which does not exist.")
    }
}

# --- 3. Roadmap IDs ---
# Roadmap IDs appear in bold and plain prose as well as backticks, so scan
# the raw status text rather than the backticked token list.
$idTokens = [regex]::Matches($status, '\bR\d+\.\d+\b') |
    ForEach-Object { $_.Value } | Select-Object -Unique
foreach ($id in $idTokens)
{
    if ($roadmap -notmatch [regex]::Escape("| $id |"))
    {
        $violations.Add("STATUS references roadmap item '$id' which is absent from $RoadmapPath.")
    }
}

if ($violations.Count -gt 0)
{
    Write-Host "check-status: $($violations.Count) violation(s):"
    $violations | ForEach-Object { Write-Host "  - $_" }
    exit 1
}

Write-Host ("check-status: OK ({0} hashes, {1} paths, {2} roadmap IDs verified)" -f $hashTokens.Count, ($tokens | Where-Object { $_ -match '[/\\]' -and $_ -notmatch '[\s*(){}]' -and $_ -notmatch '^-' -and $_ -notmatch '://' -and $_ -match '\.[A-Za-z0-9]+$' }).Count, $idTokens.Count)
exit 0
