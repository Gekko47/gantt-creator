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
$roots = $Roots + $Entry
foreach ($root in $roots)
{
    $rootPath = Join-Path $repoRoot $root
    if (-not (Test-Path -LiteralPath $rootPath)) { continue }
    Get-ChildItem -LiteralPath $rootPath -Recurse -File -Filter '*.md' | ForEach-Object {
        $file = $_.FullName
        $content = Get-Content -LiteralPath $file -Raw
        # Relative and root-relative targets only; absolute URLs, anchors,
        # and mailto links are skipped. './relative.md' links ARE validated:
        # a leading dot no longer exempts a link from the check.
        $rx = [regex]'\]\((?!https?://|#|mailto:)([^)]+)\)'
        foreach ($m in $rx.Matches($content)) {
            $rel = $m.Groups[1].Value.Trim()
            # Strip anchors
            $pathPart = ($rel -split '#')[0]
            if ([string]::IsNullOrWhiteSpace($pathPart)) { continue }
            $target = if ([System.IO.Path]::IsPathRooted($pathPart)) { $pathPart } else { Join-Path $_.DirectoryName $pathPart }
            $resolved = (Resolve-Path -LiteralPath $target -ErrorAction SilentlyContinue)
            if (-not $resolved) { $broken.Add("$file -> $rel") }
        }
    }
}

if ($broken.Count -eq 0) {
    Write-Host "OK: all relative markdown links resolve"
    exit 0
}
Write-Host "BROKEN LINKS:"
$broken | ForEach-Object { Write-Host "  $_" }
exit 1
