#requires -Version 7
<#
.SYNOPSIS
    Pester tests for verify-transfer-soak.ps1 (R7.7 D4 gate split, watchdog, fail-closed handoffs).
#>

Describe 'verify-transfer-soak.ps1' {
    BeforeAll {
        $script:scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'scripts\verify-transfer-soak.ps1'
        # Mirrors the gate's own prerequisite (verify-transfer-soak.ps1 exits 2 without this
        # key). Hosted CI runners have no Microsoft 365, so the executable positive below
        # cannot reach its fail-closed exit 5 there and is skipped rather than mis-reported.
        $script:officePrerequisitePresent = $null -ne (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue)
    }

    It 'exists and is readable' {
        (Test-Path -LiteralPath $script:scriptPath) | Should -BeTrue
    }

    It 'selects Category=TransferSoak and never Category=OfficeIntegration' {
        # D4: the soak is its own category. A body that also carried OfficeIntegration
        # would enter the routine 780 s gate instead of this one. The assertion targets the
        # filter argument itself, so the block comment's legitimate mention of the routine
        # gate cannot make a negative assertion vacuous or spuriously red.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match "--filter', 'Category=TransferSoak'"
        $codeOnly | Should -Not -Match "--filter', 'Category=OfficeIntegration'"
    }

    It 'uses its own measured deadline, not the routine 780 s' {
        # The routine OfficeIntegration gate keeps 780 s; the soak's bound is measured
        # (116 s on 2026-11-22) and must stay separate and documented.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match '\[int\]\$DeadlineSeconds = 360'
        $codeOnly | Should -Match 'MEASURED'
        $codeOnly | Should -Not -Match '\[int\]\$DeadlineSeconds = 780'
        $dict = Get-Content -LiteralPath $script:scriptPath -Raw
        $dict | Should -Match 'measured 2026-11-22: 116 s'
    }

    It 'gates the application reference count, the same leak signal as the routine gate' {
        # The soak is the one gate that performs 25 transfers with repeated staging and deletion, so
        # a leaked reference would surface here first. Before this fact existed the soak read the
        # same teardown record as the routine gate but could not fail on the count (2026-11-25).
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match "Properties\['Rcw'\]"
        $codeOnly | Should -Match '\$rcwMeasured -eq 0'
        $codeOnly | Should -Match 'leak signal was NOT measured'
        $codeOnly | Should -Match '\$rcwLeaks\.Count -gt 0'
        $codeOnly | Should -Match 'left an application reference behind'
        $codeOnly | Should -Match 'exit 7'
    }

    It 'reports the reference count beside the forced-exit signal, and names the kill as residency' {
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match 'Reference-count signal:'
        $codeOnly | Should -Match "Excel's residency, not a leaked proxy"
        $codeOnly | Should -Not -Match 'This measures the Excel shutdown race, not a proxy chain'
    }

    It 'positive control: the leak verdict is detected when it is absent' {
        # Proves the assertions above can fail: a gate with only the forced-exit signal must pass
        # none of the reference-count checks.
        $flagged = "if (`$teardowns -eq 0) { exit 5 }`n"
        $codeOnly = $flagged -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Not -Match "Properties\['Rcw'\]"
        $codeOnly | Should -Not -Match '\$rcwLeaks\.Count -gt 0'
        $codeOnly | Should -Not -Match 'Reference-count signal:'
    }

    It 'arms both harness handoff environment variables' {
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match 'GANTTCREATOR_OWNED_PIDS_PATH'
        $codeOnly | Should -Match 'GANTTCREATOR_FORCED_KILLS_PATH'
            $codeOnly | Should -Match 'soak-forced-kills.log'
    }

    It 'derives the owned-PID manifest from the evidence directory, never a committed path' {
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match "Join-Path \`$evidence 'owned-office-pids.json'"
        $codeOnly | Should -Match 'Remove-Item -LiteralPath \$ownedPidsPath'
    }

    It 'does not use --blame-hang-timeout (L12)' {
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Not -Match 'blame-hang-timeout'
    }

    It 'kills the dotnet-test process tree with taskkill /T on timeout' {
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match 'taskkill /PID \$testProc\.Id /T /F'
        $codeOnly | Should -Not -Match 'Stop-Process -InputObject \$testProc'
    }

    It 'sweeps harness-owned Office processes with PID plus start-time identity' {
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match 'Test-ManifestProcessIdentity'
        $codeOnly | Should -Match 'StartTimeUtcTicks'
        $codeOnly | Should -Match 'PID was recycled'
        $codeOnly | Should -Match 'taskkill /PID \$processId /T /F'
        $codeOnly | Should -Not -Match 'Stop-Process -InputObject \$processId'
    }

    It 'sweeps before the build and again after the run' {
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $preflightIndex = $codeOnly.IndexOf('preflight: sweeping')
        $buildIndex = $codeOnly.IndexOf('dotnet build $Solution')
        $preflightIndex | Should -BeGreaterThan 0
        $buildIndex | Should -BeGreaterThan 0
        $preflightIndex | Should -BeLessThan $buildIndex
        $saveTrx = $codeOnly.LastIndexOf('Save-Trx')
        $postRunSweep = $codeOnly.IndexOf('$finalManifest = Read-OwnedPidManifest')
        $postRunSweep | Should -BeGreaterThan $saveTrx
    }

    It 'fails closed (exit 5) when the forced-kill log records no teardown' {
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match '\$teardowns -eq 0'
        $codeOnly | Should -Match 'unmeasured signal is not a zero signal'
        $codeOnly | Should -Match 'exit 5'
    }

    It 'fails closed (exit 5) when the owned-PID manifest was not written' {
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match '\$finalManifest\.Count -eq 0'
        $codeOnly | Should -Match 'PID-ownership handoff was NOT measured'
    }

    It 'fails (exit 3) on a stray and (exit 4) above the per-teardown ceiling' {
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match '\$straysKilled -gt 0'
        $codeOnly | Should -Match 'survived the test run and had to be killed'
        $codeOnly | Should -Match 'exit 3'
        $codeOnly | Should -Match '\[int\]\$MaxForcedExitsPerTeardown = 1'
        $codeOnly | Should -Match '\$count -gt \$MaxForcedExitsPerTeardown'
        $codeOnly | Should -Match 'exit 4'
        $codeOnly | Should -Not -Match '\$forcedKills -gt'
    }

    It 'positive control: the no-blame assertion is not vacuous' {
        $flagged = "dotnet test --blame-hang-timeout 600`n"
        $codeOnly = $flagged -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match 'blame-hang-timeout' -Because 'the negative assertion must be able to detect the flag'
    }

    It 'fails closed (exit 5) on a stub run that writes no teardown record' {
        # The executable positive for the fail-closed validator: a stub dotnet passes the
        # build and the test but writes neither a forced-kill record nor a manifest, so the
        # gate must exit 5 rather than report a clean soak it never measured.
        if (-not $script:officePrerequisitePresent) {
            Set-ItResult -Skipped -Because 'verify-transfer-soak.ps1 exits 2 without the Office ClickToRun registry key, so the fail-closed exit 5 path is unreachable on this host'
            return
        }

        $stubDir = Join-Path ([System.IO.Path]::GetTempPath()) ('soak-stub-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $stubDir -Force | Out-Null
        try {
            @'
@echo off
exit /b 0
'@ | Set-Content -LiteralPath (Join-Path $stubDir 'dotnet.cmd') -Encoding utf8NoBOM
            $previousPath = $env:PATH
            $env:PATH = $stubDir + ';' + $previousPath
            try {
                & pwsh -NoProfile -File $script:scriptPath -DeadlineSeconds 30 *> $null
                $LASTEXITCODE | Should -Be 5
            } finally {
                $env:PATH = $previousPath
            }
        } finally {
            Remove-Item -LiteralPath $stubDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}
