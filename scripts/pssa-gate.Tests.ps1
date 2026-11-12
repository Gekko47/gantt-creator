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

            # The runner script lives in a SIBLING temp directory, not inside
            # $FixtureDir: Invoke-PssaGate -Path $FixtureDir must scan only the
            # fixture contents, and the runner (which sources verify-helpers.ps1)
            # would otherwise be PSSA-scanned alongside bad.ps1.
            $runnerDir = Join-Path ([System.IO.Path]::GetTempPath()) ('pssa-runner-' + [guid]::NewGuid().ToString('N'))
            New-Item -ItemType Directory -Path $runnerDir -Force | Out-Null
            $outFile = Join-Path $runnerDir 'gate-out.txt'
            $errFile = Join-Path $runnerDir 'gate-err.txt'
            $runner  = Join-Path $runnerDir 'run-gate.ps1'
            $runnerLines = @(
                ". '$($script:helpersPath)'",
                "Invoke-PssaGate -Path '$FixtureDir' -Settings '$($script:settingsPath)'"
            )
            Set-Content -LiteralPath $runner -Value ($runnerLines -join "`n") -Encoding utf8

            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile', '-File', $runner
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile
            return [pscustomobject]@{ Proc = $proc; RunnerDir = $runnerDir }
        }

        # Reproduces the Invoke-Step / Tee-Object / ForEach-Object pipeline
        # shape from scripts/verify-quick.ps1 (W13): the original -EnableExit
        # defect only surfaced as a false PASS inside this pipeline, so the
        # positive test must run the gate through the same shape, not only
        # through a bare child process.
        function Invoke-GateThroughPipeline {
            param([string]$FixtureDir)

            $runnerDir = Join-Path ([System.IO.Path]::GetTempPath()) ('pssa-pipe-' + [guid]::NewGuid().ToString('N'))
            New-Item -ItemType Directory -Path $runnerDir -Force | Out-Null
            $outFile = Join-Path $runnerDir 'pipe-out.txt'
            $errFile = Join-Path $runnerDir 'pipe-err.txt'
            $runner  = Join-Path $runnerDir 'run-gate-pipe.ps1'
            # Build the child script with string concatenation so the
            # pipeline's literal `$_` and `$null` are not interpolated away.
            $runnerContent = ". '$($script:helpersPath)'" + [Environment]::NewLine +
                ('& { Invoke-PssaGate -Path ' + "'$FixtureDir'" + ' -Settings ' +
                "'$($script:settingsPath)'" +
                ' } 2>&1 | Tee-Object -Variable stepOut | ForEach-Object { $null = $_ }')
            Set-Content -LiteralPath $runner -Value $runnerContent -Encoding utf8

            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile', '-File', $runner
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile
            return [pscustomobject]@{ Proc = $proc; RunnerDir = $runnerDir }
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

            $result = Invoke-GateOnFixture -FixtureDir $td
            $result.Proc.ExitCode | Should -Not -Be 0
        }
        finally {
            Remove-Item -LiteralPath $td -Recurse -Force -ErrorAction SilentlyContinue
            if ($result.RunnerDir -and (Test-Path -LiteralPath $result.RunnerDir)) {
                Remove-Item -LiteralPath $result.RunnerDir -Recurse -Force -ErrorAction SilentlyContinue
            }
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
            $result = Invoke-GateOnFixture -FixtureDir $td
            $result.Proc.ExitCode | Should -Be 0
        }
        finally {
            Remove-Item -LiteralPath $td -Recurse -Force -ErrorAction SilentlyContinue
            if ($result.RunnerDir -and (Test-Path -LiteralPath $result.RunnerDir)) {
                Remove-Item -LiteralPath $result.RunnerDir -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    It 'the real gate fails when the root holds nothing analyzable (positive test)' {
        # W15: the gate now resolves its own file list, so "no targets" became a
        # reachable state. A gate that analysed nothing cannot certify scripts/,
        # so it must fail with a reason rather than report PASS.
        if (-not (Get-Module -ListAvailable PSScriptAnalyzer)) {
            Set-ItResult -Skipped -Because 'PSScriptAnalyzer not installed on this host'
            return
        }
        $td = Join-Path ([System.IO.Path]::GetTempPath()) ('pssa-empty-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $td -Force | Out-Null
        try {
            Set-Content -LiteralPath (Join-Path $td 'notes.txt') -Value 'not a script' -Encoding utf8
            $result = Invoke-GateOnFixture -FixtureDir $td
            $result.Proc.ExitCode | Should -Not -Be 0
            $err = Get-Content -LiteralPath (Join-Path $result.RunnerDir 'gate-err.txt') -Raw
            $err | Should -Match 'Analysing nothing is not a pass'
        }
        finally {
            Remove-Item -LiteralPath $td -Recurse -Force -ErrorAction SilentlyContinue
            if ($result.RunnerDir -and (Test-Path -LiteralPath $result.RunnerDir)) {
                Remove-Item -LiteralPath $result.RunnerDir -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    It 'the real gate reports FAIL through the Invoke-Step/Tee pipeline (W13 shape)' {
        # The original -EnableExit defect only surfaced as a false PASS inside
        # the Tee/ForEach pipeline; a bare child-process exit code is not
        # enough to prove the fix. The gate must report FAIL through the same
        # shape verify-quick.ps1 uses.
        if (-not (Get-Module -ListAvailable PSScriptAnalyzer)) {
            Set-ItResult -Skipped -Because 'PSScriptAnalyzer not installed on this host'
            return
        }
        $td = Join-Path ([System.IO.Path]::GetTempPath()) ('pssa-pipe-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $td -Force | Out-Null
        try {
            @'
function foo { $x = 'declared but never read' }
foo
'@ | Set-Content -LiteralPath (Join-Path $td 'bad.ps1') -Encoding utf8

            $result = Invoke-GateThroughPipeline -FixtureDir $td
            $result.Proc.ExitCode | Should -Not -Be 0
        }
        finally {
            Remove-Item -LiteralPath $td -Recurse -Force -ErrorAction SilentlyContinue
            if ($result.RunnerDir -and (Test-Path -LiteralPath $result.RunnerDir)) {
                Remove-Item -LiteralPath $result.RunnerDir -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    Describe 'Get-PssaTargetFile (W15 explicit file list)' {
        BeforeAll {
            . (Join-Path (Split-Path -Parent $PSScriptRoot) 'scripts\verify-helpers.ps1')
        }

        It 'includes nested scripts and .psd1 files, and excludes the ignored _artifacts tree' {
            # The explicit list replaced `-Recurse -Exclude '_artifacts'`, so the
            # two properties that made the old call correct must be pinned:
            # recursion still reaches nested scripts, and the ignored artifact
            # tree is still skipped (it holds scratch scripts that are not gate
            # targets).
            $td = Join-Path ([System.IO.Path]::GetTempPath()) ('pssa-target-' + [guid]::NewGuid().ToString('N'))
            $deep = Join-Path $td 'sub\deeper'
            $ignored = Join-Path $td '_artifacts'
            New-Item -ItemType Directory -Path $deep -Force | Out-Null
            New-Item -ItemType Directory -Path $ignored -Force | Out-Null
            try {
                Set-Content -LiteralPath (Join-Path $td 'top.ps1') -Value '# top' -Encoding utf8
                Set-Content -LiteralPath (Join-Path $deep 'nested.psm1') -Value '# nested' -Encoding utf8
                Set-Content -LiteralPath (Join-Path $td 'settings.psd1') -Value '@{ A = 1 }' -Encoding utf8
                Set-Content -LiteralPath (Join-Path $ignored 'scratch.ps1') -Value '# scratch' -Encoding utf8
                Set-Content -LiteralPath (Join-Path $td 'notes.txt') -Value 'not a script' -Encoding utf8

                $names = @(Get-PssaTargetFile -Path @($td) | ForEach-Object { $_.Name })

                # Positive controls: recursion and the analyzer's own extension set.
                $names | Should -Contain 'top.ps1'
                $names | Should -Contain 'nested.psm1'
                $names | Should -Contain 'settings.psd1'
                # Exclusions: the ignored tree and non-script files.
                $names | Should -Not -Contain 'scratch.ps1'
                $names | Should -Not -Contain 'notes.txt'
            }
            finally {
                Remove-Item -LiteralPath $td -Recurse -Force -ErrorAction SilentlyContinue
            }
        }

        It 'returns the file itself when the root is a file path' {
            # Invoke-PssaGate accepts a file or a directory; a file must not be
            # silently dropped by the enumeration.
            $td = Join-Path ([System.IO.Path]::GetTempPath()) ('pssa-file-' + [guid]::NewGuid().ToString('N'))
            New-Item -ItemType Directory -Path $td -Force | Out-Null
            try {
                $file = Join-Path $td 'only.ps1'
                Set-Content -LiteralPath $file -Value '# only' -Encoding utf8
                $targets = @(Get-PssaTargetFile -Path @($file))
                $targets.Count | Should -Be 1
                $targets[0].FullName | Should -Be $file
            }
            finally {
                Remove-Item -LiteralPath $td -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
}
