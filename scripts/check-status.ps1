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
         extension-like suffix, no globs or URLs) must exist on disk
         AND be tracked by git. The tracking requirement matters because
         this gate also runs in CI, where only tracked files exist: a
         git-ignored path referenced here passes locally and fails the
         checkout, which no local filesystem test can reveal.
      3. Roadmap IDs    -- every R<major>.<minor>[<suffix>] token must
         appear in docs/03-ROADMAP.md, so the status cannot reference a
         work item the roadmap does not define. The letter suffix is part
         of the ID (R4.7A, R4.8A). A lowercase suffix (R2.7a) is a
         sub-item of a base row the roadmap already defines, so it is
         normalised to that base (R2.7) before the lookup and never
         demands a guide of its own. An uppercase-suffixed ID must
         additionally have a work-item guide in docs/work-items/ that
         appears in the git INDEX -- present on disk alone is not enough,
         because a clean CI checkout has only tracked files, while a
         staged new guide stays valid so a work item may be cited in the
         same commit that introduces it.
      4. Work-item evidence commands -- inside fenced code blocks in
         docs/work-items/*.md, every path a *command* line names must exist
         on disk. Both a separated path (`src/Foo/Bar.cs`) and a bare
         filename (`GanttCreator.slnx`) are checked, each resolved against the
         repository root. Roadmap/work-item IDs, quoted search patterns, and
         `$var.Member` accesses are excluded. This exists because 73 work
         items carried an evidence command naming src/GanttCreator.slnx,
         which does not exist: the command could not run at all, so the
         acceptance evidence it claimed was never produced. Comment lines
         (leading '#') are skipped, which is what lets a work item record
         a path it is deliberately correcting. Git tracking is NOT
         required here, unlike check 2: a work item may legitimately cite
         a file a later item will create, and refusing that would make the
         gate cry wolf. Measured before this check was added: zero
         violations across every work item, so the rule is quiet today.

    Globs (tokens containing *) are skipped: the R0.6 entry legitimately
    references a file that does not exist.

    Exit 0 on clean; exit 1 with one message per violation.
#>

[CmdletBinding()]
param(
    [string]$StatusPath = 'docs/STATUS.md',
    [string]$RoadmapPath = 'docs/03-ROADMAP.md',
    [string]$WorkItemsPath = 'docs/work-items'
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$statusFile = Join-Path $repoRoot $StatusPath
$roadmapFile = Join-Path $repoRoot $RoadmapPath

# The work-item guide directory, resolved once against the repository root and
# shared by the two checks that read it. Resolving it here rather than in the
# guide lookup alone means the roadmap-ID check and the evidence-command check
# can never disagree about where the guides live.
$workItemsDir = Join-Path $repoRoot $WorkItemsPath

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
    if ($t -match '(?i)^scripts[\\/]_artifacts(?:[\\/]|$)') {
        $violations.Add("STATUS references generated/ignored artifact path '$t'; document the command or work item instead.")
        continue
    }

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
        continue
    }

    # Existence on disk is not enough. This gate also runs in CI, which checks
    # out only tracked files, so a path that exists locally but is untracked
    # (or git-ignored) passes here and fails there -- the exact divergence the
    # local pre-commit run cannot see. Require git to know the path, which is
    # the condition a fresh checkout reproduces.
    #
    # The *index*, not HEAD. This gate is a pre-commit hook, so what it must
    # validate is the tree of the commit about to be made, and the index is
    # exactly that tree. `git ls-tree HEAD` excluded staged content, so a new
    # file that STATUS names in the same commit that introduces it was reported
    # as untracked and the commit was blocked -- a false rejection of a correct
    # change. `git ls-files` reads the index, so a staged path is accepted while
    # a path that exists only in the working tree (never added) is still rejected,
    # which is the distinction that actually matters. The earlier reasoning that
    # a staged-only path is a false pass for CI did not hold: once committed, the
    # file is in the commit, so a CI checkout of that commit does have it.
    #
    # CI behaviour is unchanged: in a clean checkout the index equals its HEAD,
    # so the two forms agree there. The token is matched in the canonical
    # root-relative form -- the same $rel the traversal check above already
    # computed -- rather than the raw backticked text, because that text may use a
    # backslash, carry a redundant './' or '../' segment, or end in a separator,
    # none of which is the spelling git records or prints. Deriving it from the raw
    # token made a genuinely tracked path fail the `-contains` test below and be
    # reported as untracked. A trailing separator is trimmed because the
    # tracked-directory test appends its own.
    $gitToken = $rel.TrimEnd('/')
    $tracked = @(git -C $repoRoot ls-files -- $gitToken 2>$null)

    # Captured immediately after the call, before any other command runs. Reading
    # $LASTEXITCODE further down would be reading whatever ran last, not git's
    # status, and the commands between (Join-Path, Where-Object, .StartsWith) are
    # exactly the kind that can move it.
    $lsFilesExitCode = $LASTEXITCODE

    # A token is known to git when it is exactly a tracked file, or when it is a
    # directory that contains tracked files. The directory case is required
    # because git tracks files, not directories: `ls-tree -r` reports the entries
    # beneath a directory token rather than the token itself, and STATUS cites
    # project directories such as `src/GanttCreator.Office` as readily as files.
    $isTrackedFile = @($tracked).Count -gt 0 -and $tracked -contains $gitToken
    $isTrackedDirectory = @(
        $tracked | Where-Object { $_.StartsWith($gitToken + '/', [StringComparison]::Ordinal) }
    ).Count -gt 0

    if ($lsFilesExitCode -ne 0 -or -not ($isTrackedFile -or $isTrackedDirectory))
    {
        $violations.Add("STATUS references path '$t' which is not tracked by git; it exists on disk but is untracked or ignored, so a clean checkout (CI) will not have it.")
    }
}

# --- 3. Roadmap IDs ---
# Returns whether a candidate work-item guide is known to git, reading the INDEX
# rather than HEAD.
#
# The index is what this gate must validate: it runs as a pre-commit hook, so the
# tree of the commit about to be made IS the index, and a newly authored guide
# that has been staged is legitimate even though HEAD does not have it yet.
# Using HEAD would reject exactly the correct workflow -- add a work item and cite
# it in STATUS in the same commit -- and block the change for being new.
#
# The path handed to git is the canonical root-relative form with forward
# slashes, because that is the spelling git records and prints on every host.
function Test-GuideIsTracked {
    <#
    .SYNOPSIS
        Whether a work-item guide file appears in the git index.

    .DESCRIPTION
        Takes a FileInfo for a guide inside the work-items directory and asks git
        whether the index knows that path. Reads the index (the default for
        `git ls-files`) rather than HEAD, so a staged new guide counts as tracked
        and an untracked working-tree file does not.

        An absent repository, or a git invocation that fails, is treated as
        "not tracked" rather than as a pass: a guide the gate cannot confirm is
        one the gate must report.
    #>
    param([Parameter(Mandatory)][System.IO.FileInfo]$Guide)

    $relative = [System.IO.Path]::GetRelativePath(
        [System.IO.Path]::GetFullPath($repoRoot),
        [System.IO.Path]::GetFullPath($Guide.FullName)).Replace('\', '/')

    $tracked = @(git -C $repoRoot ls-files --error-unmatch -- $relative 2>$null)
    return $LASTEXITCODE -eq 0 -and $tracked.Count -gt 0
}

# Roadmap IDs appear in bold and plain prose as well as backticks, so scan
# the raw status text rather than the backticked token list.
#
# The suffix is part of the ID. Phase 4 inserts R4.7A..R4.7H and R4.8A, and
# the previous `\bR\d+\.\d+\b` could not match them: there is no word boundary
# between "7" and "A", because both are word characters. That regex would have
# silently skipped every letter-suffixed row, so STATUS could cite a work item
# the roadmap never defined and the gate would still pass -- the exact failure
# this check exists to prevent, on exactly the rows it was about to be needed
# for.
#
# The suffix class is `[A-Za-z]` and the case is decided afterwards, because the
# two spellings mean different things. An UPPERCASE suffix (R4.7A..R4.7H,
# R4.8A) is part of the ID: the roadmap defines that exact row, and a suffixed ID
# must additionally have a work-item guide. A LOWERCASE suffix (R2.7a..R2.7d,
# R5.6a, R8.2a) is a sub-item of a base row the roadmap already defines, so it is
# normalised to that base ID (R2.7a -> R2.7) before the roadmap lookup and is
# never treated as guide-requiring.
#
# The capture therefore has to accept BOTH cases and strip the lowercase one
# explicitly. Matching only `[A-Z]` — as this check did — meant a lowercase
# reference such as `R2.7a` was captured as `R2.7` only by accident of the
# optional group's greediness, while a lowercase ID whose base row is genuinely
# absent from the roadmap (`R2.9z`) either matched its base and passed, or failed
# for the wrong reason. Normalising here makes the lowercase case deliberate:
# the base row is what must exist.
$idTokens = [regex]::Matches($status, '\bR\d+\.\d+[A-Za-z]*\b') |
    ForEach-Object {
        # Uppercase suffix kept whole; lowercase suffix removed so the base row
        # is what the roadmap is asked about.
        if ($_.Value -cmatch '^[Rr]\d+\.\d+([a-z]+)$') { $_.Value.Substring(0, $_.Value.Length - $Matches[1].Length) }
        else { $_.Value }
    } | Select-Object -Unique
foreach ($id in $idTokens)
{
    if ($roadmap -notmatch [regex]::Escape("| $id |"))
    {
        $violations.Add("STATUS references roadmap item '$id' which is absent from $RoadmapPath.")
    }

    # A roadmap row with a work-item guide is the norm; a letter-suffixed row
    # that cites no guide is how R4.7A-R4.7H would ship as nine rows nobody can
    # implement from. Only enforced for suffixed IDs so the pre-existing rows,
    # several of which are landed or deferred by name, are not retro-required.
    if ($id -match '\d+[A-Z]$')
    {
        # Resolved against $repoRoot, not the current directory. The guide
        # directory is a repository path like every other one this gate checks,
        # and `-LiteralPath 'docs/work-items'` silently resolved against the
        # caller's working directory instead -- so a run started from anywhere
        # but the repository root found no guide and reported a violation for
        # every suffixed ID. $workItemsDir is the same resolved path check 4
        # already uses, so the two cannot drift.
        #
        # A file merely PRESENT on disk is not enough, for the same reason check
        # 2 requires git tracking: a guide that exists only in the working tree
        # is absent from a clean CI checkout, so the row would be unimplementable
        # there. The *index* rather than HEAD, for the reason check 2 records --
        # this gate is a pre-commit hook, so the index is the tree of the commit
        # about to be made and a staged-but-uncommitted new guide is legitimate.
        $guide = Get-ChildItem -LiteralPath $workItemsDir -Filter "$id-*.md" -ErrorAction SilentlyContinue |
            Where-Object { Test-GuideIsTracked $_ }
        if (-not $guide)
        {
            $violations.Add("STATUS references suffixed roadmap item '$id', which has no tracked work-item guide in $WorkItemsPath.")
        }
    }
}

# --- 4. Work-item evidence commands ---
# A work item's evidence block is a promise: "run these and the gates are
# green". A command naming a path that does not exist cannot run, so the
# evidence was never produced and the promise was never kept. 73 items
# carried `src/GanttCreator.slnx` for exactly this reason.
#
# Scope is deliberately narrow, because a doc scanner that flags legitimate
# text gets disabled and then catches nothing:
#   * only inside fenced code blocks -- prose may name anything;
#   * only on non-comment lines -- a '#' line is a note, not a command, and
#     is what lets an item record a path it is deliberately correcting;
#   * every filename-shaped token is resolved against the repository root, so
#     both `src/Foo/Bar.cs` and a bare `GanttCreator.slnx` are checked --
#     skipping bare names is how the most common evidence-command shape went
#     unverified. Roadmap/work-item IDs, quoted search patterns, and
#     `$var.Member` accesses are the measured exclusions;
#   * existence on disk only, NOT git tracking. Check 2 requires tracking
#     because STATUS claims describe the repository as it stands; a work
#     item may legitimately cite a file a later item will create, and
#     refusing that would be a false positive.
# $workItemsDir is resolved once at the top of this script and shared with the
# roadmap-ID guide lookup above.
$workItemCommandLineCount = 0
if (Test-Path -LiteralPath $workItemsDir)
{
    foreach ($item in Get-ChildItem -LiteralPath $workItemsDir -Filter '*.md' -File)
    {
        $inFence = $false
        $lineNumber = 0
        foreach ($line in [System.IO.File]::ReadAllLines($item.FullName))
        {
            $lineNumber++
            if ($line -match '^\s*```')
            {
                $inFence = -not $inFence
                continue
            }
            if (-not $inFence) { continue }

            $command = $line.Trim()
            if ($command.StartsWith('#')) { continue }
            $workItemCommandLineCount++

            foreach ($match in [regex]::Matches($command, '[A-Za-z0-9_./\\-]+\.[A-Za-z0-9]+'))
            {
                $token = $match.Value

                # A URL is not a repository path. `https://example.com/build.json`
                # tokenises to `//example.com/build.json`, which carries a separator
                # and would otherwise be joined onto the repo root and tested as a
                # local path that can never exist. The scheme and the `//` authority
                # marker are both excluded, so a genuine relative path is unaffected.
                if ($token -match '://') { continue }
                if ($token.StartsWith('//')) { continue }

                if ($token -match '[?*\[\]]') { continue }

                # A bare filename is a legitimate file argument: `dotnet build
                # GanttCreator.slnx` names a real repository-root file, and the whole
                # reason this check exists is a command naming a path that cannot
                # resolve. So a token with no separator is tested against the repo
                # root instead of skipped -- with two exclusions, both measured
                # against the current corpus rather than guessed:
                #   * a roadmap/work-item ID (R3.12, R2.7d) is an identifier, not a file;
                #   * a PowerShell member access on a *variable* (`$r.FailedCount`,
                #     `_.FullName`) is a property on an object, not a filename. The
                #     `$`/`_` prefix is required: without it this rule swallowed every
                #     `Word.Word` token, including the real filenames it exists to
                #     check (`GanttCreator.slnx` matched it too).
                if ($token -notmatch '[/\\]')
                {
                    if ($token -match '^R\d+(\.\d+[a-z]?)?$') { continue }

                    # A quoted token is a search pattern, not a path the command opens:
                    # `Get-ChildItem -Filter 'UnitTest1.cs'` is a command whose expected
                    # result is that the file is *absent*, and R0.5 records exactly that.
                    # Requiring the file to exist would invert its meaning.
                    if ($command -match ('[\''"]' + [regex]::Escape($token) + '[\''"]')) { continue }

                    # A PowerShell variable member access: the token's first segment
                    # must actually be written with a `$` on the command line, so a
                    # bare `Word.Word` filename is still checked. The `$`/`_` prefix
                    # alone was not enough -- it matched every `Word.Word` token,
                    # including GanttCreator.slnx, making the check vacuous.
                    $isVariableMember = $false
                    if ($token -match '^([A-Za-z_][A-Za-z0-9_]*)\.([A-Za-z][A-Za-z0-9]*)$')
                    {
                        $isVariableMember = $command -match ('[\$]' + [regex]::Escape($Matches[1]) + '\.')
                    }

                    if ($isVariableMember) { continue }
                    if ($token -notmatch '^[\w.-]+\.[A-Za-z][A-Za-z0-9]*$') { continue }
                }

                $candidate = Join-Path $repoRoot $token
                if (-not (Test-Path -LiteralPath $candidate))
                {
                    $violations.Add("$WorkItemsPath/$($item.Name) line $lineNumber names '$token' in an evidence command, but that path does not exist; the command cannot run.")
                }
            }
        }
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
Write-Host ("check-status: OK ({0} hashes, {1} paths, {2} roadmap IDs, {3} work-item evidence lines verified)" -f $hashTokens.Count, $verifiedPathCount, $idTokens.Count, $workItemCommandLineCount)
exit 0
