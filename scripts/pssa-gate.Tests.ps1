#requires -Version 7
<#
.SYNOPSIS
    Pester tests for the PSScriptAnalyzer gate used by verify-quick.ps1
    and verify.ps1 (scripts/verify-helpers.ps1 -> Invoke-PssaGate).

.DESCRIPTION
    The original step inlined `Invoke-ScriptAnalyzer ... -EnableExit`.
    `-EnableExit` calls `exit` from inside PowerShell's function scope,
    which is not propagated to the caller when the cmdlet runs inside a
    Tee/ForEach pipeline (the Invoke-Step shape), so real failures were
    reported as PASS. The fix extracted the capture-then-judge gate into
    Invoke-PssaGate (scripts/verify-helpers.ps1) and both verify scripts
    now call that one function, so a behavioural test of the function is
    a behavioural test of both verify flows.

    The positive test runs the actual gate against a synthetic bad.ps1
    fixture through a child pwsh process and asserts the process exits
    non-zero; the positive control runs it against a clean fixture and
    asserts exit 0. No analyzer result objects or private implementation
    details are inspected.
#>

Describe 'PSScriptAnalyzer gate (verify scripts)' {
    BeforeAll {
        $script:repoRoot = Split-Path -Parent $PSScriptRoot
        $script:helpersPath  = Join-Path $script:repoRoot 'scripts\verify-helpers.ps1'
        $script:settingsPath = Join-Path $script:repoRoot 'scripts\PSScriptAnalyzerSettings.psd1'
        # Runs the real gate (Invoke-PssaGate from scripts/verify-helpers.ps1)
        # against a fixture directory in a child pwsh and returns the process.
        # Only the child's observable exit code is asserted by the caller.
        # Defined inside BeforeAll (Pester 6 scope): top-level function
        # declarations are not visible to It blocks.
        function Invoke-GateOnFixture {
            param([string]$FixtureDir)

            $outFile = Join-Path $FixtureDir 'gate-out.txt'
            $errFile = Join-Path $FixtureDir 'gate-err.txt'
            $runner  = Join-Path $FixtureDir 'run-gate.ps1'
            $runnerLines = @(
                ". '$($script:helpersPath)'",
                "Invoke-PssaGate -Path '$FixtureDir' -Settings '$($script:settingsPath)'"
            )
            Set-Content -LiteralPath $runner -Value ($runnerLines -join "`n") -Encoding utf8

            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile', '-File', $runner
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile
            return $proc
        }
    }

    It 'verify-quick.ps1 and verify.ps1 call the shared gate, never -EnableExit' {
        $helpers = Get-Content -LiteralPath $script:helpersPath -Raw
        $helpers | Should -Match 'function Invoke-PssaGate'
        # "-EnableExit" legitimately appears in explanatory comments; only a
        # real Invoke-ScriptAnalyzer call may never use it, and neither verify
        # script may call the analyzer inline at all (full delegation).
        $helpers | Should -Not -Match 'Invoke-ScriptAnalyzer[^\r\n]*-EnableExit'
        foreach ($name in @('verify-quick.ps1', 'verify.ps1')) {
            $text = Get-Content -LiteralPath (Join-Path $script:repoRoot "scripts\$name") -Raw
            $text | Should -Match 'Invoke-PssaGate'
            $text | Should -Not -Match 'Invoke-ScriptAnalyzer'
        }
    }

    It 'the real gate fails a script with a PSSA finding (positive test)' {
        if (-not (Get-Module -ListAvailable PSScriptAnalyzer)) {
            Set-ItResult -Skipped -Because 'PSScriptAnalyzer not installed on this host'
            return
        }
        $td = Join-Path ([System.IO.Path]::GetTempPath()) ('pssa-gate-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $td -Force | Out-Null
        try {
            @'
function foo { $x = 'declared but never read' }
foo
'@ | Set-Content -LiteralPath (Join-Path $td 'bad.ps1') -Encoding utf8

            $proc = Invoke-GateOnFixture -FixtureDir $td
            $proc.ExitCode | Should -Not -Be 0
        }
        finally {
            Remove-Item -LiteralPath $td -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'the real gate passes a clean script (positive control)' {
        if (-not (Get-Module -ListAvailable PSScriptAnalyzer)) {
            Set-ItResult -Skipped -Because 'PSScriptAnalyzer not installed on this host'
            return
        }
        $td = Join-Path ([System.IO.Path]::GetTempPath()) ('pssa-clean-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $td -Force | Out-Null
        try {
            Set-Content -LiteralPath (Join-Path $td 'ok.ps1') -Value "# fixture without analyzable code" -Encoding utf8
            $proc = Invoke-GateOnFixture -FixtureDir $td
            $proc.ExitCode | Should -Be 0
        }
        finally {
            Remove-Item -LiteralPath $td -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}
