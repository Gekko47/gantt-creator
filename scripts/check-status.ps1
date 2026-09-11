#requires -Version 7
<#
.SYNOPSIS
    Accuracy gate for docs/STATUS.md. Wired into verify-quick.ps1 so the
    status file cannot drift from the repository it describes.

.DESCRIPTION
    Three checks, all derived from the text of docs/STATUS.md itself:

      1. Commit hashes  -- every backticked token of 7-40 lowercase hex
         characters must resolve to a commit (git rev-parse --verify).
      2. Repo paths     -- every backticked token that looks like a
         repo-relative file path (contains a separator, ends in an
         extension-like suffix, no globs or URLs) must exist on disk.
      3. Roadmap IDs    -- every backticked R<major>.<minor> token must
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

if (-not (Test-Path -LiteralPath $statusFile)) { Write-Error "Missing $statusFile"; exit 1 }
if (-not (Test-Path -LiteralPath $roadmapFile)) { Write-Error "Missing $roadmapFile"; exit 1 }

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
function Get-RepoPathTokenStatus {
    <#
    .SYNOPSIS
        Classifies a backticked STATUS token for the repo-path check.

    .DESCRIPTION
        Returns a status object describing whether the token is a
        repo-relative path candidate and, when it is, whether it carries
        wildcard metacharacters that must be rejected before any filesystem
        test is attempted.

          IsPath      $true  the token matches the repo-path shape: it has a
                                separator, ends in an extension-like suffix, and
                                carries no whitespace, glob, URL, or leading-dash
                                noise.
          HasWildcard $true  when IsPath is $true but the token contains ?[]
                                (a glob-traversal risk that must be rejected).
          Both $false        the token is not a path at all (a commit hash, a
                                URL, an attribute annotation like [Fact], ...).

        The base-shape predicate and the wildcard rejection live here together
        so the validation loop and the verified-path summary count share one
        definition; a token rejected for wildcard metacharacters is excluded
        from the check-status: OK total.
    #>
    param([Parameter(Mandatory)][string]$Token)

    $isPath = $Token -match '[/\\]' `
        -and $Token -notmatch '[\s*(){}]' `
        -and $Token -notmatch '^-' `
        -and $Token -notmatch '://' `
        -and $Token -match '\.[A-Za-z0-9]+$'

    if (-not $isPath) {
        return [pscustomobject]@{ Token = $Token; IsPath = $false; HasWildcard = $false }
    }

    # Reject wildcard metacharacters that could be used for glob traversal.
    # Only real paths are checked here; attribute annotations like [Fact] or
    # [Trait(...)] never reach this branch because they fail the base shape
    # (no separator, and they contain spaces or parentheses).
    $hasWildcard = $Token -match '[?\[\]]'

    return [pscustomobject]@{ Token = $Token; IsPath = $isPath; HasWildcard = $hasWildcard }
}

foreach ($t in $tokens)
{
    $pathStatus = Get-RepoPathTokenStatus -Token $t
    if (-not $pathStatus.IsPath) { continue }

    if ($pathStatus.HasWildcard) {
        $violations.Add("STATUS references path '$t' contains wildcard metacharacters (?[\]); rejected.")
        continue
    }

    $candidate = Join-Path $repoRoot $t
    # Reject path traversal: the resolved path must stay inside $repoRoot.
    # GetRelativePath gives the canonical lexical difference; ".." or a
    # path that starts with "..\" means the candidate escapes the root
    # (e.g. C:\repos\gantt-creator-evil would pass a StartsWith check
    # against C:\repos\gantt-creator on Ordinal comparison).
    $resolved = [System.IO.Path]::GetFullPath($candidate)
    $repoRootFull = [System.IO.Path]::GetFullPath($repoRoot)
    # Normalise the relative path to forward slashes so parent traversal is
    # recognised with both Windows ('..\') and POSIX ('../') separators,
    # regardless of the host that runs the gate.
    $rel = [System.IO.Path]::GetRelativePath($repoRootFull, $resolved).Replace('\', '/')
    if ($rel -eq '..' -or $rel.StartsWith('../')) {
        $violations.Add("STATUS references path '$t' which resolves outside the repository.")
        continue
    }

    # Use -LiteralPath so the candidate is treated as a literal path,
    # not a glob pattern (wildcards already rejected above, but -LiteralPath
    # is the correct API for exact filesystem checks).
    if (-not (Test-Path -LiteralPath $candidate))
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

# Verified-path count: every token classified as a repo path by the shared
# predicate, minus those rejected for wildcard metacharacters, so the OK
# total never reports a rejected token as verified.
$verifiedPathCount = ($tokens | ForEach-Object { Get-RepoPathTokenStatus -Token $_ } |
    Where-Object { $_.IsPath -and -not $_.HasWildcard }).Count
Write-Host ("check-status: OK ({0} hashes, {1} paths, {2} roadmap IDs verified)" -f $hashTokens.Count, $verifiedPathCount, $idTokens.Count)
exit 0
