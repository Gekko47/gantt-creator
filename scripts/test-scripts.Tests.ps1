#requires -Version 7
<#
.SYNOPSIS
    Pester tests for scripts/test-scripts.ps1, the script-lint gate.

.DESCRIPTION
    W15 changed the runner from one file at a time to parallel runspaces. The
    verdict rules were extracted into Get-ScriptTestGateVerdict
    (scripts/verify-helpers.ps1) so they can be tested in-process without
    spawning Pester: the tested judge is the judge the runner executes, and the
    suite stays fast enough to be worth running on every commit.

    Two end-to-end tests then run the real gate in a child process against
    fixture directories, proving the runner is wired to that judge and exits
    with its verdict. Everything else is an in-process verdict case.
#>

Describe 'Get-ScriptTestGateVerdict' {
    BeforeAll {
        . (Join-Path (Split-Path -Parent $PSScriptRoot) 'scripts\verify-helpers.ps1')

        function Get-FileResultStub {
            param(
                [string]$Name = 'a.Tests.ps1',
                [int]$Total = 1,
                [int]$Passed = 1,
                [int]$Failed = 0,
                [string[]]$Failures = @(),
                [string]$ErrorText = ''
            )
            [pscustomobject]@{
                Name     = $Name
                Total    = $Total
                Passed   = $Passed
                Failed   = $Failed
                Failures = $Failures
                Error    = $ErrorText
            }
        }
    }

    It 'passes when every file passes and every file returned a result (positive control)' {
        $verdict = Get-ScriptTestGateVerdict -DiscoveredCount 2 -Results @(
            (Get-FileResultStub -Name 'a.Tests.ps1' -Total 2 -Passed 2),
            (Get-FileResultStub -Name 'b.Tests.ps1' -Total 3 -Passed 3)
        )
        $verdict.ExitCode | Should -Be 0
        $verdict.Total    | Should -Be 5
        $verdict.Passed   | Should -Be 5
        $verdict.Errors.Count | Should -Be 0
    }

    It 'fails when one file fails even though another passes (positive test)' {
        $verdict = Get-ScriptTestGateVerdict -DiscoveredCount 2 -Results @(
            (Get-FileResultStub -Name 'a.Tests.ps1' -Total 1 -Passed 1),
            (Get-FileResultStub -Name 'b.Tests.ps1' -Total 1 -Passed 0 -Failed 1 -Failures @('b.Tests.ps1: expected 2 but got 1'))
        )
        $verdict.ExitCode | Should -Not -Be 0
        $verdict.Failed   | Should -Be 1
        ($verdict.Lines -join "`n") | Should -Match 'b\.Tests\.ps1'
        ($verdict.Lines -join "`n") | Should -Match 'expected 2 but got 1'
    }

    It 'fails when a discovered file contains zero tests (no-silent-pass policy)' {
        $verdict = Get-ScriptTestGateVerdict -DiscoveredCount 1 -Results @(
            (Get-FileResultStub -Name 'empty.Tests.ps1' -Total 0 -Passed 0)
        )
        $verdict.ExitCode | Should -Not -Be 0
        ($verdict.Errors -join "`n") | Should -Match 'empty\.Tests\.ps1 contains no Pester tests'
    }

    It 'fails when a file produced no result object (crashed worker)' {
        $verdict = Get-ScriptTestGateVerdict -DiscoveredCount 1 -Results @(
            (Get-FileResultStub -Name 'crash.Tests.ps1' -Total 0 -Passed 0 -ErrorText 'runspace terminated')
        )
        $verdict.ExitCode | Should -Not -Be 0
        ($verdict.Lines -join "`n") | Should -Match 'crash\.Tests\.ps1 did not produce a result'
        ($verdict.Lines -join "`n") | Should -Match 'runspace terminated'
    }

    It 'fails when fewer results came back than files were discovered (lost worker)' {
        # Judging only the results that arrived would silently shrink the suite.
        $verdict = Get-ScriptTestGateVerdict -DiscoveredCount 2 -Results @(
            (Get-FileResultStub -Name 'a.Tests.ps1' -Total 1 -Passed 1)
        )
        $verdict.ExitCode | Should -Not -Be 0
        ($verdict.Errors -join "`n") | Should -Match 'A worker was lost'
    }

    It 'reports a passing file as a positive line rather than staying silent' {
        $verdict = Get-ScriptTestGateVerdict -DiscoveredCount 1 -Results @(
            (Get-FileResultStub -Name 'a.Tests.ps1' -Total 4 -Passed 4)
        )
        ($verdict.Lines -join "`n") | Should -Match 'PASSED: 4 tests'
    }
}

Describe 'test-scripts.ps1 (end to end)' {
    BeforeAll {
        $script:gatePath = Join-Path (Split-Path -Parent $PSScriptRoot) 'scripts\test-scripts.ps1'

        # Writes the supplied test files into a fresh temp directory and runs the
        # real gate against it in a child process. -Serial keeps these two wiring
        # tests cheap; the parallel path shares the same judge (above).
        function Invoke-GateOnFixture {
            param(
                [Parameter(Mandatory)][hashtable]$Files,
                [switch]$Serial,
                # 0 leaves the gate's own default (4 workers). A explicit value
                # drives the parallel path, which is where the CI flake lived.
                [int]$ThrottleLimit = 0
            )

            $dir = Join-Path ([System.IO.Path]::GetTempPath()) ('test-scripts-' + [guid]::NewGuid().ToString('N'))
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            foreach ($name in $Files.Keys) {
                Set-Content -LiteralPath (Join-Path $dir $name) -Value $Files[$name] -Encoding utf8
            }

            $outFile = Join-Path $dir 'gate-out.txt'
            $errFile = Join-Path $dir 'gate-err.txt'
            $arguments = @('-NoProfile', '-File', $script:gatePath, '-ScriptsRoot', $dir)
            if ($Serial) { $arguments += '-Serial' }
            if ($ThrottleLimit -gt 0) { $arguments += @('-ThrottleLimit', $ThrottleLimit) }

            $proc = Start-Process -FilePath pwsh -ArgumentList $arguments -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile
            $combined = (Get-Content -LiteralPath $outFile -Raw) + (Get-Content -LiteralPath $errFile -Raw)

            Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
            return [pscustomobject]@{ Exit = $proc.ExitCode; Output = $combined }
        }
    }

    It 'exists and is readable' {
        (Test-Path -LiteralPath $script:gatePath) | Should -BeTrue
    }

    It 'exits non-zero and names the failing file when one file fails and another passes' {
        $r = Invoke-GateOnFixture -Serial -Files @{
            'a-pass.Tests.ps1' = "Describe 'pass' { It 'passes' { 1 | Should -Be 1 } }"
            'b-fail.Tests.ps1' = "Describe 'fail' { It 'fails' { 1 | Should -Be 2 } }"
        }
        $r.Exit   | Should -Not -Be 0
        $r.Output | Should -Match 'b-fail\.Tests\.ps1'
        $r.Output | Should -Match 'test-scripts: 1 test\(s\) failed'
    }

    It 'accepts an absolute -ScriptsRoot and runs every discovered file (positive control)' {
        # Join-Path concatenates a rooted second argument, so an absolute root
        # used to be reported as a missing directory.
        $r = Invoke-GateOnFixture -Serial -Files @{
            'dotted.Tests.ps1' = "Describe 'dotted' { It 'passes' { 1 | Should -Be 1 } }"
            'plainTests.ps1'   = "Describe 'plain' { It 'passes' { 1 | Should -Be 1 } }"
        }
        $r.Exit   | Should -Be 0
        $r.Output | Should -Match 'test-scripts: Total=2, Passed=2, Failed=0'
    }

    It 'preserves Describe/Context/It structure through the parallel ThrottleLimit path' {
        # Regression for CI 45193295738 / 45376223093: the parallel dispatch used
        # to hand FileInfo objects across the runspace boundary and then
        # intermittently died with "The provided Path argument was null or an
        # empty collection." before reporting a single file. Every test in a
        # nested tree must still be discovered, run, and counted.
        $r = Invoke-GateOnFixture -ThrottleLimit 2 -Files @{
            'a-nested.Tests.ps1' = @'
Describe 'outer' {
    Context 'inner' {
        It 'one' { 1 | Should -Be 1 }
        It 'two' { 2 | Should -Be 2 }
    }
}
'@
            'b-nested.Tests.ps1' = @'
Describe 'other' {
    Context 'deep' {
        It 'three' { 3 | Should -Be 3 }
        It 'four' { 4 | Should -Be 4 }
        It 'five' { 5 | Should -Be 5 }
    }
}
'@
        }
        $r.Output | Should -Match 'Mode: parallel'
        $r.Exit   | Should -Be 0
        $r.Output | Should -Match 'test-scripts: Total=5, Passed=5, Failed=0'
        $r.Output | Should -Not -Match 'did not produce a result'
        $r.Output | Should -Not -Match 'unhandled error'
    }

    It 'names the failing file through the parallel path instead of reporting an unhandled error' {
        $r = Invoke-GateOnFixture -ThrottleLimit 2 -Files @{
            'a-pass.Tests.ps1' = "Describe 'pass' { It 'passes' { 1 | Should -Be 1 } }"
            'b-fail.Tests.ps1' = "Describe 'fail' { It 'fails' { 1 | Should -Be 2 } }"
        }
        $r.Output | Should -Match 'Mode: parallel'
        $r.Exit   | Should -Not -Be 0
        $r.Output | Should -Match 'b-fail\.Tests\.ps1'
        $r.Output | Should -Match 'test-scripts: 1 test\(s\) failed'
        $r.Output | Should -Not -Match 'unhandled error'
    }
}

Describe 'Assert-UsablePath (a null or empty path is refused, never defaulted)' {
    BeforeAll {
        . (Join-Path (Split-Path -Parent $PSScriptRoot) 'scripts\verify-helpers.ps1')
    }

    It 'accepts a usable path (positive control)' {
        { Assert-UsablePath -Path (Join-Path $PSScriptRoot 'test-scripts.ps1') -Caller 'control' } |
            Should -Not -Throw
    }

    It 'refuses a null path, so Invoke-Pester can never fall back to the current directory' {
        { Assert-UsablePath -Path $null -Caller 'null case' } |
            Should -Throw '*null, empty, or whitespace-only path*'
    }

    It 'refuses an empty path collection' {
        { Assert-UsablePath -Path @() -Caller 'empty collection' } |
            Should -Throw '*refusing an empty path collection*'
    }

    It 'refuses an empty string' {
        { Assert-UsablePath -Path '' -Caller 'empty string' } |
            Should -Throw '*null, empty, or whitespace-only path*'
    }

    It 'refuses a whitespace-only path' {
        { Assert-UsablePath -Path '   ' -Caller 'whitespace' } |
            Should -Throw '*null, empty, or whitespace-only path*'
    }

    It 'refuses an array that contains a blank entry' {
        { Assert-UsablePath -Path @('C:\valid', '') -Caller 'mixed array' } |
            Should -Throw '*null, empty, or whitespace-only path*'
    }
}

Describe 'test-scripts.ps1 parallel path hardening (CI 45193295738 / 45376223093)' {
    BeforeAll {
        $script:gateText = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'test-scripts.ps1') -Raw
    }

    It 'discovers full path strings instead of passing FileInfo objects across -Parallel' {
        $script:gateText | Should -Match 'ForEach-Object \{ \$_\.FullName \}'
        $script:gateText | Should -Not -Match '\$file\.FullName'
    }

    It 'guards the worker path inside its try before Invoke-Pester runs' {
        $script:gateText | Should -Match 'IsNullOrWhiteSpace\(\$testPath\)'
        $script:gateText | Should -Match 'IsPathRooted\(\$testPath\)'
        $script:gateText | Should -Match 'Invoke-Pester -Path \$testPath'
    }

    It 'hands each worker the parent-resolved Pester manifest (fix B)' {
        # The parent resolves Pester once; a worker that imported by
        # -RequiredVersion repeated a PSModulePath lookup inside a fresh
        # runspace, which is where the hosted-runner failure surfaced.
        $script:gateText | Should -Match 'Import-Module -Name \$pesterManifest'
        $script:gateText | Should -Match 'toolVersions\.Pester\.Version'
    }

    It 'dispatches through one reusable runspace pool, not a per-file parallel pipeline (fix A)' {
        $script:gateText | Should -Match 'CreateRunspacePool'
        $script:gateText | Should -Not -Match 'ForEach-Object -Parallel'
        $script:gateText | Should -Match 'AddScript\(\$workerScript\)\.AddArgument\(\$testPath\)'
    }

    It 'prints the call site and exception identity when an unhandled error escapes' {
        # Three CI runs reported only the message, which is why the call site was
        # never visible.
        $script:gateText | Should -Match 'ScriptStackTrace'
        $script:gateText | Should -Match 'FullyQualifiedErrorId'
    }

    It 'pre-creates the Pester registry key before any worker starts (fresh-runner race)' {
        # Pester creates HKCU\Software\Pester on first use; concurrent workers
        # raced that create on a fresh runner and the losers failed their whole
        # container (CI 45411123131). A developer machine already has the key,
        # which is why this never reproduced locally.
        $script:gateText | Should -Match 'HKCU:\\Software\\Pester'
    }
}

Describe 'Format-CheckAnnotation' {
    BeforeAll {
        . (Join-Path (Split-Path -Parent $PSScriptRoot) 'scripts\verify-helpers.ps1')
        $script:gateText = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'test-scripts.ps1') -Raw
    }

    It 'formats a title-only annotation (positive control)' {
        Format-CheckAnnotation -Title 'test-scripts: a.Tests.ps1' -Message 'expected 2 but got 1' |
            Should -Be '::error title=test-scripts%3A a.Tests.ps1::expected 2 but got 1'
    }

    It 'points at the repository-relative test file when one is supplied' {
        Format-CheckAnnotation -Title 't' -Message 'm' -File 'scripts/a.Tests.ps1' |
            Should -Be '::error file=scripts/a.Tests.ps1,title=t::m'
    }

    It 'flattens a multi-line message and escapes the percent sign (command cannot be broken)' {
        $line = Format-CheckAnnotation -Title 't' -Message "line1`r`nline2 50%"
        $line | Should -Be '::error title=t::line1 line2 50%25'
        $line | Should -Not -Match "`n"
    }

    It 'escapes the property delimiters in a title' {
        Format-CheckAnnotation -Title 'a:b,c' -Message 'm' |
            Should -Be '::error title=a%3Ab%2Cc::m'
    }

    It 'is wired into the runner behind the GITHUB_ACTIONS gate' {
        $script:gateText | Should -Match 'Format-CheckAnnotation'
        $script:gateText | Should -Match ([regex]::Escape("if (`$env:GITHUB_ACTIONS -eq 'true')"))
    }
}
