#requires -Version 7
<#
.SYNOPSIS
    Verifies an actionlint release archive against its pinned SHA-256.

.DESCRIPTION
    Single source of the archive-integrity check used by scripts/lint-ci.ps1
    (W11). Computes the SHA-256 of the archive at <ZipPath> and compares it
    to <ExpectedHash> (the value single-sourced from
    scripts/tool-versions.psd1#actionlint.Sha256). Exits 0 on a match, exits
    1 (with a Write-Error) on a mismatch. Kept standalone so the check can be
    exercised in an isolated child process with a fixture archive — no network
    download required.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ZipPath,
    [Parameter(Mandatory)][string]$ExpectedHash
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $ZipPath)) {
    Write-Error "actionlint archive not found: $ZipPath"
    exit 1
}

$actualHash = (Get-FileHash -Path $ZipPath -Algorithm SHA256).Hash.ToLower()
if ($actualHash -ne $ExpectedHash) {
    Write-Error "actionlint archive SHA-256 mismatch: expected $ExpectedHash, got $actualHash"
    exit 1
}

exit 0