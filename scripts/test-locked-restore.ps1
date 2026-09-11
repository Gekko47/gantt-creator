#requires -Version 7
#
# Deliberate locked-restore test (R0.3 acceptance).
#
# Four-phase proof that the committed lock files are honoured without
# network access:
#
#   Run 1  Clean restore (obj/ removed). Regenerates the restore graph
#          and warms the global packages folder.
#   Run 2  Clean restore again (obj/ removed). Proves the restore is
#          deterministic from the lock files across a clean state.
#   Run 3  NEGATIVE CONTROL. Restore with a NuGet.config whose only
#          source is https://127.0.0.1:1 (a closed loopback port: every
#          connection is refused immediately, deterministically, with no
#          firewall or container support) and NUGET_PACKAGES pointed at
#          an empty directory. This MUST fail: with the package cache
#          empty and no reachable source, NuGet cannot acquire anything.
#          If it unexpectedly succeeds, the blocked source is not
#          actually denying package acquisition and run 4 would be a
#          false PASS. Expect roughly a minute: NuGet retries refused
#          connections before giving up.
#   Run 4  OFFLINE PROOF. Clean restore (obj/ removed) with the same
#          blocked-source NuGet.config but the normal global packages
#          folder. --locked-mode must succeed purely from the lock
#          files plus the local package cache: no reachable source
#          exists, so any network dependency would fail the run.
#
# Both blocked runs pass -p:NuGetAudit=false. NuGetAudit fetches
# vulnerability data from the configured source; that is a network
# dependency orthogonal to lock-file discipline (this script's subject)
# and it fails the blocked runs with NU1900 "Warning As Error" under the
# repo's warnings-as-errors policy even when the lock file alone is
# sufficient (observed empirically on .NET SDK 10.0.400). The
# vulnerability check itself is verify.ps1's
# 'dotnet list package --vulnerable' step.
#
# Exit 0 only if runs 1, 2, and 4 pass AND run 3 fails as required.

[CmdletBinding()]
param(
    [string]$Solution = 'GanttCreator.slnx'
)

$ErrorActionPreference = 'Stop'

function Remove-ObjDirectory {
    [CmdletBinding(SupportsShouldProcess)]
    param()

    Get-ChildItem -Path $PSScriptRoot\.. -Recurse -Directory -Filter 'obj' -ErrorAction SilentlyContinue |
        ForEach-Object {
            if ($PSCmdlet.ShouldProcess($_.FullName, 'Delete obj directory')) {
                Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
}

function Write-BlockedNuGetConfig {
    <#
        The only package source is a closed loopback port. NuGet refuses
        plain-HTTP sources outright (NU1302, at config validation, before
        any cache lookup) so the endpoint must be HTTPS; connecting to a
        closed port then fails with NU1301 (connection refused), exactly
        like an unreachable remote source.
    #>
    [CmdletBinding()]
    param([string]$Path)

    $config = @'
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <packageSources>
    <!-- Clear every inherited source (nuget.org etc.): nothing reachable remains. -->
    <clear />
    <!-- Closed loopback port: every connection refused, offline-deterministic. -->
    <add key="offline-blocked" value="https://127.0.0.1:1/v3/index.json" />
  </packageSources>
</configuration>
'@
    Set-Content -LiteralPath $Path -Value $config -Encoding utf8NoBOM
}

function Invoke-LockRestore {
    param(
        [string]$Label,
        [string]$ConfigFile,
        [switch]$ExpectFailure
    )
    Write-Host "=== $Label ==="
    $restoreArgs = @('restore', $Solution, '--locked-mode', '--no-cache')
    if ($ConfigFile) {
        $restoreArgs += @('--configfile', $ConfigFile, '-p:NuGetAudit=false')
    }
    dotnet @restoreArgs 2>&1
    if ($ExpectFailure) {
        # The negative control MUST fail. A success here means the blocked
        # source is not denying package acquisition (e.g. the config was
        # ignored), so run 4's offline proof would be a false PASS.
        if ($LASTEXITCODE -eq 0) {
            Write-Error "restore unexpectedly SUCCEEDED: $Label -- the blocked-source NuGet.config does not deny package acquisition, so the offline run would be a false PASS."
            exit 1
        }
        Write-Host "PASS (failed as required): $Label"
    }
    else {
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Lock-mode restore failed: $Label"
            exit $LASTEXITCODE
        }
        Write-Host "PASS: $Label"
    }
}

Write-Host "R0.3 locked-restore test: two clean runs, a denied-network control, and an offline run"
Write-Host ""

$blockedConfig = Join-Path ([System.IO.Path]::GetTempPath()) ("nuget-blocked-" + [guid]::NewGuid() + ".config")
Write-BlockedNuGetConfig -Path $blockedConfig

try {
    # --- Run 1: generate the lock files ---
    Remove-ObjDirectory
    Invoke-LockRestore -Label 'Run 1 -- generate lock files (no cache)'

    # --- Run 2: prove the lock files are honoured from a clean state ---
    Remove-ObjDirectory
    Invoke-LockRestore -Label 'Run 2 -- honour lock files from clean state'

    # --- Run 3: negative control. Empty cache + no reachable source:
    #     restore MUST fail, proving the blocked config genuinely denies
    #     package acquisition. (Expect ~1 min: NuGet retries refused
    #     connections.)
    $previousPackages = $env:NUGET_PACKAGES
    $emptyPackages = Join-Path ([System.IO.Path]::GetTempPath()) ("empty-nuget-packages-" + [guid]::NewGuid())
    New-Item -ItemType Directory -Path $emptyPackages -Force | Out-Null
    $env:NUGET_PACKAGES = $emptyPackages
    try {
        Invoke-LockRestore -Label 'Run 3 control -- empty cache with blocked sources must fail' `
            -ConfigFile $blockedConfig -ExpectFailure
    }
    finally {
        if ($null -ne $previousPackages) { $env:NUGET_PACKAGES = $previousPackages }
        else { Remove-Item Env:\NUGET_PACKAGES -ErrorAction SilentlyContinue }
        Remove-Item -LiteralPath $emptyPackages -Recurse -Force -ErrorAction SilentlyContinue
    }

    # --- Run 4: offline proof. Clean state + no reachable source, but the
    #     normal global packages folder: --locked-mode must succeed purely
    #     from the lock files and the local cache.
    Remove-ObjDirectory
    Invoke-LockRestore -Label 'Run 4 -- locked restore with no reachable source (offline proof)' `
        -ConfigFile $blockedConfig
}
finally {
    Remove-Item -LiteralPath $blockedConfig -Force -ErrorAction SilentlyContinue
}

Write-Host ""
Write-Host "R0.3 locked-restore test: PASS (runs 1, 2, and 4 green; run 3 failed as required)"
exit 0
