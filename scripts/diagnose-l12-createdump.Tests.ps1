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

        $codeOnly | Should -Match 'Start-Process\s+-FilePath\s+\$createdump\.FullName' -Because 'Step 2 must invoke createdump directly via Start-Process -FilePath $createdump.FullName, not merely mention it in comments or log messages'
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

    It 'Step 1 launch, poll, and classification are wrapped in try/finally that kills the owned process tree' {
        # A terminating error or interruption after launch must not leak the
        # owned testhost process tree; the explicit deadline branch alone is
        # not enough.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''

        $codeOnly | Should -Match '\$step1Proc\s*=\s*\$null\s*\r?\ntry\s*\{' -Because 'the process variable must be nulled before the guarded launch'
        $codeOnly | Should -Match '(?s)finally\s*\{.{0,500}?taskkill\s+/PID\s+\$step1Proc\.Id\s+/T\s+/F' -Because 'the finally block must terminate the owned testhost process tree'
    }

    It 'positive control: a stub without the Step 1 try/finally teardown fails the owned-process assertion' {
        $flagged = @'
$step1Proc = Start-Process -FilePath 'dotnet' -ArgumentList $step1Args -NoNewWindow -PassThru
while (-not $step1Proc.HasExited) { Start-Sleep -Seconds 1 }
'@
        $codeOnly = $flagged -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Not -Match '\$step1Proc\s*=\s*\$null\s*\r?\ntry\s*\{'
        $codeOnly | Should -Not -Match '(?s)finally\s*\{.{0,500}?taskkill\s+/PID\s+\$step1Proc\.Id\s+/T\s+/F'
    }

    It 'Step 0 resolves the runtime from runtimeOptions.framework/frameworks and never falls back to the first recursive createdump' {
        # The old 'Microsoft.NETCore.App.RuntimeVersion' property read never
        # matched the standard runtimeconfig schema, so Step 0 silently
        # degraded to an arbitrary recursive createdump selection.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''

        $codeOnly | Should -Match '\$rc\.runtimeOptions\.framework' -Because 'the requested version comes from the standard single-framework schema'
        $codeOnly | Should -Match '\$rc\.runtimeOptions\.frameworks' -Because 'the multi-framework schema must be handled too'
        $codeOnly | Should -Not -Match "'Microsoft\.NETCore\.App\.RuntimeVersion'" -Because 'that property is not part of the standard runtimeconfig schema'
        $codeOnly | Should -Match 'roll-forward-matches' -Because 'framework parsing or version resolution failure must skip Step 2, not pick an arbitrary binary'
        $bareRecursive = 'Get-ChildItem\s+-Path\s+\$dotnetRoot\s+-Recurse\s+-Filter\s+''createdump\.exe''\s+-ErrorAction\s+SilentlyContinue\s*\|\s*Select-Object\s+-First\s+1'
        $codeOnly | Should -Not -Match $bareRecursive -Because 'the first recursive binary must never be selected when runtime resolution fails'
    }

    It 'positive control: a stub falling back to the first recursive createdump fails the runtime-selection assertion' {
        $flagged = @'
$createdump = Get-ChildItem -Path $dotnetRoot -Recurse -Filter 'createdump.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
'@
        $codeOnly = $flagged -replace '(?m)^\s*#.*$', ''
        $bareRecursive = 'Get-ChildItem\s+-Path\s+\$dotnetRoot\s+-Recurse\s+-Filter\s+''createdump\.exe''\s+-ErrorAction\s+SilentlyContinue\s*\|\s*Select-Object\s+-First\s+1'
        $codeOnly | Should -Match $bareRecursive -Because 'this stub deliberately re-introduces the arbitrary recursive pick'
    }

    It 'Step 2 reports dump-written only on a zero exit code with a non-empty dump file' {
        # A failed createdump can leave a zero-byte or partial artifact on
        # disk; a file merely existing must not read as dump-written.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''

        $codeOnly | Should -Match '\$step2Result\.ExitCode\s+-eq\s+0\s+-and\s+\(Test-Path\s+-LiteralPath\s+\$dumpPath\)\s+-and\s+\(Get-Item\s+-LiteralPath\s+\$dumpPath\)\.Length\s+-gt\s+0' -Because 'dump-written requires a zero exit code and a non-empty dump'
        $codeOnly | Should -Not -Match '\}\s*elseif\s*\(Test-Path\s+-LiteralPath\s+\$dumpPath\)\s*\{' -Because 'the old classification reported any on-disk artifact as dump-written'
    }

    It 'positive control: a stub classifying any on-disk dump artifact as dump-written fails the exit/size assertion' {
        $flagged = @'
} elseif (Test-Path -LiteralPath $dumpPath) {
    $dumpBytes = (Get-Item -LiteralPath $dumpPath).Length
    $step2Result.Outcome = 'dump-written'
'@
        $codeOnly = $flagged -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match '\}\s*elseif\s*\(Test-Path\s+-LiteralPath\s+\$dumpPath\)\s*\{' -Because 'this stub deliberately re-introduces the artifact-only classification'
        $codeOnly | Should -Not -Match '\$step2Result\.ExitCode\s+-eq\s+0\s+-and\s+\(Test-Path\s+-LiteralPath\s+\$dumpPath\)'
    }
}
