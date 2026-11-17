#requires -Version 7
<#
.SYNOPSIS
    NuGet package vulnerability scan. CI entry point (W15).

.DESCRIPTION
    W15 moved this scan out of verify.ps1. `dotnet list package --vulnerable`
    reads advisory data from the configured NuGet sources, so it needs network
    access and can block for a long time on a machine that has none. The local
    branch-final gate is required to work offline; CI has network, so CI runs the
    scan and the local gate stays offline-capable.

    Informational by contract: `dotnet list` exits 0 whether or not it finds a
    vulnerability, so this script cannot fail the build on a finding. The team
    has not approved a vulnerable-package gate; when it does, the explicit gate
    belongs here and the caller already fails on a non-zero exit.

    Exit 0 when the scan completed (findings or not); the dotnet exit code when
    the command itself failed.
#>

[CmdletBinding()]
param(
    [string]$Solution = 'GanttCreator.slnx'
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

dotnet list $Solution package --vulnerable --include-transitive
exit $LASTEXITCODE
