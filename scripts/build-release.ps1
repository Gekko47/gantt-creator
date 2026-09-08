#requires -Version 7
<#
.SYNOPSIS
    Release build with warnings-as-errors (W8 local/CI parity entry point).

.DESCRIPTION
    Single source of truth for the Release build command shared by
    verify-quick.ps1 step 10, verify.ps1 step 10, and the GitHub CI
    'Build Release (warnings as errors)' step. Inline `dotnet build` in
    ci.yml is forbidden by scripts/ci-parity.Tests.ps1 (W8); this entry
    point is the replacement, so local and CI cannot drift.

    Exit 0 on clean; the dotnet exit code on failure.
#>

[CmdletBinding()]
param(
    [string]$Solution      = 'GanttCreator.slnx',
    [string]$Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

dotnet build $Solution -c $Configuration --no-restore -warnaserror
exit $LASTEXITCODE