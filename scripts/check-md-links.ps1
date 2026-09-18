#requires -Version 7
# Validates that all relative links in the kit markdown files resolve.
# Roots are resolved against the repository root (the parent of this script's
# directory), so the check behaves identically no matter the caller's CWD.
# Walks docs/, .github/, and AGENTS.md by default.
# Exits 0 if all links resolve; exits 1 with a list of broken links otherwise.
[CmdletBinding()]
param(
    [string[]]$Roots = @('docs', '.github'),
    [string]$Entry  = 'AGENTS.md'
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

$broken = New-Object System.Collections.Generic.List[string]
$scanned = 0
$roots = $Roots + $Entry

# Removes fenced code blocks (``` or ~~~, CommonMark rules) so link
# examples inside documentation samples are never validated as real
# links. Only the fence structure is parsed, not Markdown itself: an
# opening fence is up to 3 spaces of indent, a run of 3+ backticks
# (with an info string that contains no backtick) or tildes, and a
# closing fence is the same character, at least as long, with nothing
# but whitespace after it. Shorter or different-character runs and
# unclosed fences stay code through end of file.
function Get-MarkdownTextWithoutFencedCodeBlock {
    [CmdletBinding()]
    param([string]$Text)
    $lines = $Text -split "`r?`n"
    $kept = New-Object System.Collections.Generic.List[string]
    $fenceChar = $null
    $fenceLen = 0
    $fenceIsQuoted = $false
    $fenceBuffer = New-Object System.Collections.Generic.List[string]
    foreach ($line in $lines) {
        if ($null -eq $fenceChar) {
            if ($line -match '^(?<quote>(?: {0,3}> {0,3})?)(?: {0,3})(?<fence>`{3,}|~{3,})(?<info>.*)$') {
                $fence = $Matches['fence']
                $info = $Matches['info']
                $quotePrefix = $Matches['quote']
                # A backtick info string must not contain a backtick
                # (CommonMark); a tilde info string has no such rule.
                if ($fence[0] -eq '`' -and $info -match '`') { $kept.Add($line); continue }
                $fenceChar = $fence[0]
                $fenceLen = $fence.Length
                $fenceIsQuoted = ($quotePrefix -ne '')
                $fenceBuffer.Clear()
                continue
            }
            $kept.Add($line)
        } elseif ($line -match '^(?<quote>(?: {0,3}> {0,3})?)(?: {0,3})(?<fence>`{3,}|~{3,})\s*$') {
            $fence = $Matches['fence']
            $quotePrefix = $Matches['quote']
            # Require the same block-quote prefix style as the opening fence.
            if ($fence[0] -eq $fenceChar -and $fence.Length -ge $fenceLen -and ($quotePrefix -ne '') -eq $fenceIsQuoted) {
                $fenceChar = $null
                $fenceLen = 0
                $fenceIsQuoted = $false
                $fenceBuffer.Clear()
                continue
            }
            $fenceBuffer.Add($line)
        } else {
            $fenceBuffer.Add($line)
        }
    }
    # Append any remaining buffered lines when the last fence was unclosed
    if ($null -ne $fenceChar) {
        foreach ($bufferedLine in $fenceBuffer) {
            $kept.Add($bufferedLine)
        }
    }
    return ($kept -join "`n")
}
foreach ($root in $roots)
{
    $rootPath = Join-Path $repoRoot $root
    if (-not (Test-Path -LiteralPath $rootPath)) {
        Write-Error "check-md-links: configured scan root '$root' does not exist at $rootPath. A missing configured root fails the gate instead of being silently skipped."
        exit 1
    }
    Get-ChildItem -LiteralPath $rootPath -Recurse -File -Filter '*.md' | ForEach-Object {
        $script:scanned++
        $file = $_.FullName
        $content = Get-Content -LiteralPath $file -Raw
        $content = Get-MarkdownTextWithoutFencedCodeBlock -Text $content
        # Relative and root-relative targets only; absolute URLs, anchors,
        # and mailto links are skipped. './relative.md' links ARE validated:
        # a leading dot no longer exempts a link from the check.
        $rx = [regex]'\]\((?!https?://|#|mailto:)([^)]+)\)'
        foreach ($m in $rx.Matches($content)) {
            $rel = $m.Groups[1].Value.Trim()
            # Strip anchors
            $pathPart = ($rel -split '#')[0]
            if ([string]::IsNullOrWhiteSpace($pathPart)) { continue }
            $target = if ([System.IO.Path]::IsPathRooted($pathPart)) { Join-Path $repoRoot $pathPart.TrimStart('/', '\') } else { Join-Path $_.DirectoryName $pathPart }
            $resolved = (Resolve-Path -LiteralPath $target -ErrorAction SilentlyContinue)
            if (-not $resolved) { $broken.Add("$file -> $rel") }
        }
    }
}

if ($broken.Count -gt 0) {
    Write-Host "BROKEN LINKS:"
    $broken | ForEach-Object { Write-Host "  $_" }
    exit 1
}
if ($scanned -eq 0) {
    # No-silent-pass: a zero-document scan proves nothing about the
    # link budget, so an empty scan is a failure with an instruction,
    # never a PASS. See W9.
    Write-Error "check-md-links: scanned zero markdown files; verify the roots exist and contain .md files."
    exit 1
}
Write-Host "OK: all relative markdown links resolve ($scanned file(s) scanned)"
exit 0
