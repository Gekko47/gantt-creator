#requires -Version 7
<#
.SYNOPSIS
    Installs the lightweight pre-commit gate as the repository git hook.

.DESCRIPTION
    Configures core.hooksPath to .githooks/ (repository root) and installs the
    pre-commit wrapper that delegates to scripts/pre-commit.ps1. The
    hooksPath indirection keeps hooks versioned in-repo (they cannot
    drift from the codebase) and works on Windows where a bare
    .git/hooks/pre-commit file is not an executable bit issue.

    The gate runs only the fast, deterministic checks (a few seconds):
    skill-tree drift, status accuracy, and markdown links. It
    deliberately does NOT run the full verify-quick.ps1 (~2-5 min;
    runtime depends on the machine) so a commit is not slowed down by
    the Release build and test suite.

    Policy per AGENTS.md / docs/05-GIT-QUALITY.md: the developer still
    runs pwsh ./scripts/verify-quick.ps1 during editing and
    pwsh ./scripts/verify.ps1 before a PR. The hook is a safety net
    that guarantees the commit itself cannot introduce drift, not a
    replacement for the gates.

    Uninstall: git config --unset core.hooksPath; Remove-Item .githooks -Recurse (optional).
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $PSCommandPath
$repoRoot   = Split-Path -Parent $scriptRoot
$hooksDir   = Join-Path $repoRoot '.githooks'

New-Item -ItemType Directory -Path $hooksDir -Force | Out-Null

# The on-disk hook is a tiny POSIX sh shim that delegates to the versioned
# scripts/pre-commit.ps1. Git for Windows executes hooks with the bundled
# /bin/sh (msys), so the hook file must be a sh script, not a Windows
# batch file. A shim is required because the hook file cannot itself be a
# PowerShell script (git does not know how to exec a .ps1).
    $shim = @'
#!/bin/sh
# Delegates to the versioned pre-commit gate. Installed by
# scripts/install-pre-commit.ps1; do not edit by hand.
# Resolve scripts/pre-commit.ps1 relative to the hooks directory so the
# shim is portable across checkout paths and core.hooksPath configurations.
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
exec pwsh -NoProfile -ExecutionPolicy Bypass -File "$SCRIPT_DIR/../scripts/pre-commit.ps1"
'@
$hookFile = Join-Path $hooksDir 'pre-commit'
# Write as UTF-8 without BOM, LF line endings (sh requires LF).
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$contentLf = $shim -replace "`r`n", "`n"
[System.IO.File]::WriteAllText($hookFile, $contentLf, $utf8NoBom)

# POSIX hosts execute the hook file directly, so the shim must carry the
# executable bit; git for Windows runs hooks through its bundled sh and
# does not need one.
if (-not $IsWindows) {
    & chmod +x $hookFile
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Failed to set executable permission on $hookFile."
        exit 1
    }
}

git -C $repoRoot config core.hooksPath .githooks
if ($LASTEXITCODE -ne 0) { Write-Error 'Failed to set core.hooksPath'; exit 1 }

Write-Host 'pre-commit hook installed.'
Write-Host '  hooksPath.cfg : .githooks'
Write-Host '  delegate      : scripts/pre-commit.ps1'
Write-Host '  gates         : skill-tree drift, status accuracy, markdown links'
Write-Host '  full verify   : still run pwsh ./scripts/verify-quick.ps1 during editing and pwsh ./scripts/verify.ps1 before a PR'
exit 0