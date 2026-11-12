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
                [switch]$Serial
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
}
