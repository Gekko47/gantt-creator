#requires -Version 7
# Validates that all relative links in the kit markdown files resolve.
# Walks docs/, .clinerules/, AGENTS.md, and .github/ recursively.
# Exits 0 if all links resolve; exits 1 with a list of broken links otherwise.
[CmdletBinding()]
param(
    [string[]]$Roots = @('docs', '.clinerules', '.github'),
    [string]$Entry  = 'AGENTS.md'
)

$broken = New-Object System.Collections.Generic.List[string]
$roots = $Roots + $Entry
foreach ($root in $roots) {
    if (-not (Test-Path $root)) { continue }
    Get-ChildItem -Path $root -Recurse -File -Filter '*.md' | ForEach-Object {
        $file = $_.FullName
        $content = Get-Content -Raw $file
        $rx = [regex]'\]\((?!https?://|#|mailto:|\.)([^)]+)\)'
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
