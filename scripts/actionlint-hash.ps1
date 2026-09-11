#requires -Version 7
<#
.SYNOPSIS
    Verifies an actionlint payload (release archive or executable) against
    its pinned SHA-256.

.DESCRIPTION
    Single source of the integrity check used by scripts/lint-ci.ps1 (W11).
    Computes the SHA-256 of the file at <ZipPath> (release archive) or
    <FilePath> (extracted executable) and compares it to <ExpectedHash>
    (a value single-sourced from scripts/tool-versions.psd1#actionlint).
    Exits 0 on a match, exits 1 (with a Write-Error) on a mismatch. Kept
    standalone so the check can be exercised in an isolated child process
    with a fixture file — no network download required.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory, ParameterSetName = 'Zip', Position = 0)][string]$ZipPath,
    [Parameter(Mandatory, ParameterSetName = 'File', Position = 0)][string]$FilePath,
    [Parameter(Mandatory)][string]$ExpectedHash
)

$ErrorActionPreference = 'Stop'

$path = if ($PSCmdlet.ParameterSetName -eq 'Zip') { $ZipPath } else { $FilePath }

if (-not (Test-Path -LiteralPath $path)) {
    Write-Error "actionlint payload not found: $path"
    exit 1
}

$actualHash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLower()
if ($actualHash -ne $ExpectedHash) {
    Write-Error "actionlint SHA-256 mismatch: expected $ExpectedHash, got $actualHash"
    exit 1
}

exit 0