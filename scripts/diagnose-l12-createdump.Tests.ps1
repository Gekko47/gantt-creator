#requires -Version 7
<#
.SYNOPSIS
    Pester tests for diagnose-l12-createdump.ps1.
#>

Describe 'diagnose-l12-createdump.ps1' {
    BeforeAll {
        $script:scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'scripts\diagnose-l12-createdump.ps1'
    }

    It 'exists and is readable' {
        (Test-Path -LiteralPath $script:scriptPath) | Should -BeTrue
    }

    It 'does not pass --blame-hang-timeout without --blame-hang-dump-type none' {
        # The whole point of this diagnostic is that it re-enables the blame
        # hang collector under a constrained flag. The script must pair
        # --blame-hang + --blame-hang-timeout + --blame-hang-dump-type none,
        # and must not ship a bare --blame-hang-timeout the way the old
        # verify-office.ps1 did before L12.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''

        # It must contain the paired flags. Step 1 passes them as a
        # comma-separated PowerShell argument array, so the dump-type option
        # and its 'none' value are separate quoted arguments, and the hang
        # collector must be a standalone argument (a longer related flag such
        # as --blame-hang-timeout must not satisfy this check).
        $codeOnly | Should -Match '--blame-hang-dump-type''?,\s*''none' -Because 'Step 1 must constrain the dump type'
        $codeOnly | Should -Match '--blame-hang''?,' -Because 'Step 1 must enable hang collection'
        $codeOnly | Should -Match '--blame-hang-timeout' -Because 'Step 1 must set a hang timeout'

        # The bare-without-dump-type pattern that caused L12 must not appear.
        # The numeric timeout may carry a duration suffix (the script itself
        # passes "${Step1TimeoutSeconds}s", e.g. '60s') or be followed
        # immediately by whitespace or end-of-input
        # (e.g. '... --blame-hang-timeout 600\n').
        $barePattern = '--blame-hang-timeout\s+''?\d+(ms|s|m|h)?e?''?(\s|$|,)'
        $codeOnly | Should -Not -Match $barePattern -Because 'a bare blame-hang-timeout without dump-type none is the L12 regression'
    }

    It 'Step 2 is a parent-process dump smoke test, not a debug-privilege probe' {
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''

        $codeOnly | Should -Match 'createdump' -Because 'Step 2 must invoke createdump directly'
        $codeOnly | Should -Match 'Start-Sleep\s+-Seconds\s+600' -Because 'Step 2 needs a predictable long-lived helper process'
        $codeOnly | Should -Match 'parent-process dump smoke test' -Because 'Step 2 must be classified only as a dump smoke test, not a debug-privilege probe'
        $codeOnly | Should -Match '\$helperProc' -Because 'the helper-process guidance must name the runtime-triggered dump target'
        $codeOnly | Should -Not -Match 'debug-privilege probe' -Because 'Step 2 must not claim to probe remote-process debug privilege'
    }

    It 'does not kill user-owned Office processes' {
        # The diagnostic runs the real OfficeIntegration suite in Step 1, and
        # the same ownership discipline as verify-office.ps1 must apply: never
        # kill by Office process name.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''

        $codeOnly | Should -Not -Match 'taskkill\s.*[\/]F.*EXCEL' -Because 'the diagnostic must not kill user-owned Excel even when it runs the real suite'
    }

    It 'writes a report under the ignored _artifacts tree' {
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''

        $codeOnly | Should -Match '(?s)_artifacts.*?l12-diagnostic' -Because 'evidence must land under the ignored artifacts tree, never committed'
        $codeOnly | Should -Match 'l12-diagnostic-report\.txt' -Because 'the report file name is stable for retrieval'
    }

    It 'does not declare the root cause solved' {
        # The script must present findings for a human to interpret, not claim
        # success based on a non-throwing run. This mirrors the L12 discipline
        # that a non-throwing COM / Office call does not prove the gate passed.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''

        $codeOnly | Should -Match 'Human interpretation' -Because 'the script must hand off interpretation to a human'
        $codeOnly | Should -Not -Match 'exit 0.*root cause' -Because 'a zero exit must not be conflated with a solved root cause'
    }

    It 'positive control: a flagged stub that re-enables bare --blame-hang-timeout fails the paired-flags assertion' {
        $flagged = @'
dotnet test --blame-hang-timeout 600
'@
        $codeOnly = $flagged -replace '(?m)^\s*#.*$', ''
        $barePattern = '--blame-hang-timeout\s+''?\d+(ms|s|m|h)?e?''?(\s|$|,)'
        $codeOnly | Should -Match $barePattern -Because 'the negative assertion must be able to detect a bare blame-hang-timeout'
        $codeOnly | Should -Not -Match '--blame-hang-dump-type\s+none' -Because 'this flagged stub deliberately omits the dump-type constraint'
    }

    It 'positive control: a suffixed bare --blame-hang-timeout 60s also fails the paired-flags assertion' {
        # The script passes "${Step1TimeoutSeconds}s" so a bare 60s timeout
        # must also trip the detector; otherwise the negative assertion has
        # a suffix blind spot.
        $flagged = @'
dotnet test --blame-hang-timeout 60s
'@
        $codeOnly = $flagged -replace '(?m)^\s*#.*$', ''
        $barePattern = '--blame-hang-timeout\s+''?\d+(ms|s|m|h)?e?''?(\s|$|,)'
        $codeOnly | Should -Match $barePattern -Because 'the negative assertion must detect a bare suffixed blame-hang-timeout'
        $codeOnly | Should -Not -Match '--blame-hang-dump-type\s+none' -Because 'this flagged stub deliberately omits the dump-type constraint'
    }

    It 'Step 1 validates inputs and build artifacts before launching' {
        # Step 1 runs with --no-build --no-restore, so missing inputs must
        # fail distinctly instead of being misclassified as a testhost abort.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''

        $codeOnly | Should -Match 'solution not found' -Because 'a missing solution must fail distinctly before launch'
        $codeOnly | Should -Match 'built .* test artifacts' -Because 'absent no-build artifacts must fail distinctly before launch'
        $codeOnly | Should -Match 'invalid-args' -Because 'invalid arguments must keep distinct handling from a testhost abort'
        $codeOnly | Should -Match 'RedirectStandardOutput' -Because 'classification must capture command output, not just timing'
        $codeOnly | Should -Match '\$step1Crashed' -Because 'the crash flag must remain the abort signal'
    }

    It 'Step 2 never reuses a stale dump from a previous run' {
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''

        $codeOnly | Should -Match 'step2-parent\.dmp' -Because 'Step 2 writes its dump to a stable evidence path'
        $codeOnly | Should -Match 'Remove-Item\s+-LiteralPath\s+\$dumpPath' -Because 'a stale dump must be removed before createdump runs so failure cannot read as dump-written'
    }

    It 'positive control: a flagged stub that lacks createdump lacks the Step 2 probes' {
        $flagged = @'
dotnet test --blame-hang --blame-hang-timeout 60s --blame-hang-dump-type none
'@
        $codeOnly = $flagged -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match '--blame-hang-dump-type\s+none' -Because 'this stub passes Step 1''s constraints'
        $codeOnly | Should -Not -Match 'createdump' -Because 'this stub deliberately omits Step 2'
    }
}
