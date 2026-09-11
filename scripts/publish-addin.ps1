#requires -Version 7
<#
.SYNOPSIS
    Publishes the AddIn and produces the packed XLL (W8 local/CI parity entry point).

.DESCRIPTION
    Single source of truth for the AddIn publish command shared by
    verify-quick.ps1 step 10, verify.ps1 step 10, and the GitHub CI
    'Publish AddIn (packed XLL artifact)' step. Inline `dotnet publish`
    in ci.yml is forbidden by scripts/ci-parity.Tests.ps1 (W8); this
    entry point is the replacement, so local and CI cannot drift.

    Exit 0 on clean; the dotnet exit code on failure.
#>

[CmdletBinding()]
param(
    [string]$Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

dotnet publish src/GanttCreator.AddIn/GanttCreator.AddIn.csproj -c $Configuration --no-build
exit $LASTEXITCODE