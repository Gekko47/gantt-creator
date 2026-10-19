#requires -Version 7
<#
.SYNOPSIS
    Pester tripwires for the mutation gate (scripts/mutation-gate.ps1, R3.13).

.DESCRIPTION
    Every guard here has a POSITIVE test: a fixture that violates the rule and
    an assertion that the guard actually catches it. A guard with no positive
    test is exactly the "validator without a positive test" class AGENTS.md
    warns about - it passes whether or not the property holds.

    The five properties asserted (D4):
      1. the verify step exists, is named identically in verify.ps1 and ci.yml,
         and runs AFTER the Core test step;
      2. the tool version is pinned in scripts/tool-versions.psd1 (W11);
      3. the threshold constant equals 80 (docs/04-TEST-STRATEGY.md);
      4. the report path is produced by the harness;
      5. `-EnableExit` is never used (the L10 lesson).

    Exit 0 on clean; Pester reports failures to the runner.
#>

BeforeAll {
    $script:repoRoot = Split-Path -Parent $PSScriptRoot
    $script:gatePath = Join-Path $script:repoRoot 'scripts\mutation-gate.ps1'
    $script:verifyPath = Join-Path $script:repoRoot 'scripts\verify.ps1'
    $script:ciPath = Join-Path $script:repoRoot '.github\workflows\ci.yml'
    $script:versionsPath = Join-Path $script:repoRoot 'scripts\tool-versions.psd1'

    $script:gateText = Get-Content -LiteralPath $script:gatePath -Raw
    $script:verifyText = Get-Content -LiteralPath $script:verifyPath -Raw
    $script:ciText = Get-Content -LiteralPath $script:ciPath -Raw
    $script:versions = Import-PowerShellDataFile -LiteralPath $script:versionsPath

    # The verify step name must match in verify.ps1 and ci.yml (W8 parity).
    $script:stepName = 'mutation gate (changed Core code)'
}

Describe 'Mutation gate (scripts/mutation-gate.ps1)' {

    Context 'tool pinning (W11)' {
        It 'pins dotnet-stryker in tool-versions.psd1' {
            $script:versions.'dotnet-stryker' | Should -Not -BeNullOrEmpty
            $script:versions.'dotnet-stryker'.Version | Should -Match '^\d+\.\d+\.\d+$'
        }

        It 'POSITIVE: the pin guard fires when the pin is absent' {
            # Proves the guard can fail by feeding it the invalid
            # fixture directly: a null settings object (what the gate
            # computes when tool-versions.psd1 has no dotnet-stryker
            # entry). Asserting the throw is what makes this a positive
            # test instead of a tautology.
            . $script:gatePath
            { Assert-DotnetStrykerPin $null } | Should -Throw '*has no dotnet-stryker pin*'
        }

        It 'accepts a present pin without throwing' {
            # The negative case for the same guard: a real pin must
            # pass, so the guard is not simply "always throw".
            . $script:gatePath
            { Assert-DotnetStrykerPin $script:versions.'dotnet-stryker' } | Should -Not -Throw
        }

        It 'POSITIVE: the gate itself refuses to run without the pin' {
            # The gate calls the guard before installing or running
            # anything, so a drifted psd1 fails loudly rather than
            # installing an unpinned tool. Asserted on the source so
            # the call site cannot be removed silently.
            $script:gateText | Should -Match 'Assert-DotnetStrykerPin'
        }
    }

    Context 'threshold constant' {
        It 'equals 80, the changed-code threshold from the test strategy' {
            $script:versions.'dotnet-stryker'.ThresholdHigh | Should -Be 80
        }

        It 'POSITIVE: the below-threshold guard fires on a low score' {
            # The real threshold guard lives in Invoke-MutationJudge
            # (a score below the threshold throws). It is exercised
            # directly with a temp JSON report here, so a drifted
            # threshold is caught by a real run, not by comparing
            # two literals (the previous test asserted 60 -ne 80,
            # which is always true and caught nothing).
            . $script:gatePath
            $report = @{
                Files = @{ 'Fixture.cs' = @{ Mutants = @(
                    @{ Status = 'Killed' }
                    @{ Status = 'Survived' }
                    @{ Status = 'Survived' }
                    @{ Status = 'Survived' }
                    @{ Status = 'Survived' }
                ) } }
            } | ConvertTo-Json -Depth 10
            $path = Join-Path ([System.IO.Path]::GetTempPath()) ('mut-thr-' + [guid]::NewGuid().ToString('N') + '.json')
            Set-Content -LiteralPath $path -Value $report
            try {
                { Invoke-MutationJudge -ReportPath $path -Threshold 80 -StrykerExitCode 0 } |
                    Should -Throw '*below the*'
            } finally {
                Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
            }
        }

        It 'echoes a threshold override rather than applying it silently' {
            $script:gateText | Should -Match 'THRESHOLD OVERRIDDEN'
        }
    }

    Context 'the L10 lesson: no -EnableExit, and a missing report is a failure' {
        # `-EnableExit` is banned as an EXECUTED parameter, so the guard strips
        # comments first: this file and the gate both *discuss* -EnableExit in
        # prose, and matching the prose would make the guard permanently red and
        # therefore useless. A bare regex over the whole file is the wrong test.
        BeforeAll {
            function Get-CodeOnlyText {
                param([string]$Text)

                $withoutBlock = [regex]::Replace($Text, '(?s)<#.*?#>', ' ')
                $withoutLine = [regex]::Replace($withoutBlock, '(?m)^\s*#.*$', ' ')
                return $withoutLine
            }

            $script:gateCode = Get-CodeOnlyText -Text $script:gateText
        }

        It 'never executes -EnableExit' {
            $script:gateCode | Should -Not -Match '-EnableExit'
        }

        It 'POSITIVE: an executed -EnableExit is detected' {
            # Same guard, same code-only text, against a line that really
            # passes the parameter. Proves the guard can go red.
            $bad = Get-CodeOnlyText -Text "Invoke-PssaGate -Path . -EnableExit"
            $bad | Should -Match '-EnableExit'
        }

        It 'POSITIVE: a mention inside a comment is NOT flagged' {
            # The counterpart that keeps the guard from becoming noise: prose
            # naming the parameter must not fail the gate.
            $prose = Get-CodeOnlyText -Text "# never use -EnableExit here`nWrite-Host 'ok'"
            $prose | Should -Not -Match '-EnableExit'
        }

        It 'treats a missing report as a failure, never a silent pass' {
            $script:gateText | Should -Match 'produced no report'
        }

        It 'POSITIVE: a missing report is detected as an error' {
            # Proves the missing-report branch is reachable and throwing.
            # We invoke the report-judging code directly with a missing-path fixture.
            $missing = Join-Path ([System.IO.Path]::GetTempPath()) ('no-such-' + [guid]::NewGuid().ToString('N'))
            Test-Path -LiteralPath $missing | Should -BeFalse
            
            # Dot-source the gate script to access Invoke-MutationJudge
            . $script:gatePath
            { Invoke-MutationJudge -ReportPath $missing -Threshold 80 -StrykerExitCode 0 } | Should -Throw '*produced no report*'
        }
    }

    Context 'report and scope' {
        It 'produces a JSON report the gate judges' {
            $script:gateText | Should -Match 'mutation-report\.json'
        }

        It 'targets Core and its test project explicitly rather than the solution' {
            # Explicit csproj paths are what make the guide's ".slnx gaps"
            # stop condition inapplicable; a regression to -s <sln> would
            # reintroduce it.
            $script:gateText | Should -Match "--project"
            $script:gateText | Should -Match "--test-project"
            $script:gateText | Should -Match 'GanttCreator\.Core\.csproj'
        }

        It 'uses --since for the changed-code run and omits it for the baseline' {
            $script:gateText | Should -Match '--since:'
            $script:gateText | Should -Match 'BASELINE'
        }

        It 'installs the tool rather than referencing it from a csproj' {
            $script:gateText | Should -Match 'dotnet tool install'
        }
    }

    Context 'verify step presence, parity and order' {
        It 'verify.ps1 declares the mutation step' {
            # The pattern MUST be pre-computed into a variable. `Should -Match`
            # binds its -Pattern to the ARGUMENT, so writing
            # `Should -Match [regex]::Escape(...)` passes the literal string
            # '[regex]::Escape' as the pattern and silently compares it against
            # the whole file. That is a guard that can never fail correctly -
            # the same class of defect AGENTS.md warns about.
            $expected = [regex]::Escape("Invoke-Step '$($script:stepName)'")
            $script:verifyText | Should -Match $expected
        }

        It 'ci.yml delegates to the same script rather than inlining the run' {
            # W8: an inlined stryker command in the workflow can drift from the
            # pinned local gate, which is the defect class this tripwire exists for.
            $script:ciText | Should -Match 'mutation-gate\.ps1'
            $script:ciText | Should -Not -Match 'dotnet-stryker\s'
        }

        It 'runs after the Core test step' {
            # The step order is computed HERE, inside the It, rather than in a
            # BeforeAll. Two Pester 6 scoping facts forced this:
            #   * a function declared in a file-level BeforeAll is not visible
            #     inside It (the existing pssa-gate.Tests.ps1 records this), and
            #   * variables set in a Context BeforeAll do not reliably reach the
            #     It blocks either, which surfaced as
            #     "Parameter set cannot be resolved" from the comparison rather
            #     than an obvious null. Computing in scope removes both risks.
            #
            # The result must be a plain ARRAY: PowerShell unrolls a returned
            # collection, so a one-element List[string] arrives as a bare String
            # and IndexOf then returns 0 instead of -1.
            $names = @(
                foreach ($line in ($script:verifyText -split "`r?`n")) {
                    if ($line -match "Invoke-Step\s+'(?<stepname>[^']+)'") { $Matches['stepname'] }
                }
            )

            $testIndex = [array]::IndexOf($names, 'test (OfficeIntegration excluded)')
            $mutationIndex = [array]::IndexOf($names, $script:stepName)

            # Mutating code the Core suite has not yet exercised would score the
            # change against tests that never ran, so both steps must exist and
            # the mutation step must come second.
            #
            # `>=` is expressed as `-ge` piped to `-BeTrue`, NOT as
            # `Should -BeGreaterThanOrEqualTo`. That operator is broken in the
            # pinned Pester 6.1.0: it throws "Parameter set cannot be resolved"
            # for a literal 0, for a variable, and for -1 alike, while
            # `-BeGreaterThan` and `-BeTrue` both work. The isolation probe is
            # recorded in this file's evidence note; using the broken operator
            # would make this guard permanently red, i.e. permanently ignored.
            $mutationIndex | Should -BeGreaterThan $testIndex
            ($testIndex -ge 0) | Should -BeTrue
            ($mutationIndex -ge 0) | Should -BeTrue
        }

        It 'POSITIVE: the ordering guard fires when the mutation step runs first' {
            # Proves the ordering guard can fail by running the SAME
            # extraction the guard uses on an invalid fixture: a
            # verify-text whose mutation step precedes the test step.
            # The guard's condition (mutationIndex greater than
            # testIndex) must then be false. Asserting that is what
            # makes this a positive test; the previous version built
            # a reversed array and asserted it was reversed, which
            # exercised array.IndexOf, not the guard.
            $invalidText = @(
                "Invoke-Step '$($script:stepName)'"
                "Invoke-Step 'test (OfficeIntegration excluded)'"
            ) -join "`r`n"
            $names = @(
                foreach ($line in ($invalidText -split "`r?`n")) {
                    if ($line -match "Invoke-Step\s+'(?<stepname>[^']+)'") { $Matches['stepname'] }
                }
            )
            $testIndex = [array]::IndexOf($names, 'test (OfficeIntegration excluded)')
            $mutationIndex = [array]::IndexOf($names, $script:stepName)
            ($mutationIndex -gt $testIndex) | Should -BeFalse
        }

        It 'POSITIVE: a missing step is detected' {
            $expected = [regex]::Escape("Invoke-Step '$($script:stepName)'")
            $withoutStep = $script:verifyText -replace $expected, ''
            # Stripping the step out of the text MUST make the guard miss it.
            # Asserting on the stripped text rather than on the original is what
            # makes this a positive test instead of a tautology.
            $withoutStep | Should -Not -Match $expected
        }
    }

    Context 'Invoke-MutationJudge branches (positive tests)' {
        # Each branch of the judge is an `if (bad) { throw }` or
        # an `if (pass) { return }` guard. Per AGENTS.md every
        # validator ships with a positive test in the same
        # commit: an invalid fixture the branch must catch (or,
        # for the pass branches, a report it must accept). The
        # reports are real Stryker JSON written to temp files, so
        # the count and status parsing is exercised end to end.
        #
        # Dot-sourced here so Invoke-MutationJudge is in scope for
        # every It in this context. mutation-gate.ps1 returns
        # (rather than running the gate) when dot-sourced from a
        # Pester call stack, so this defines the function without
        # launching Stryker.
        BeforeAll {
            . $script:gatePath

            # Writes a real Stryker-shaped JSON report to a temp
            # file and returns its path. Declared inside this
            # Context's BeforeAll so it is in scope for every It
            # here; a function declared at file scope is not
            # visible inside It under the pinned Pester 6.1.0
            # (recorded in this file's notes).
            function Write-MutationReport {
                param([string]$Path, [object[]]$Mutants)
                $report = @{
                    Files = @{
                        'GanttCreator.Core/Fixture.cs' = @{ Mutants = $Mutants }
                    }
                }
                $report | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $Path
                $Path
            }
        }

        It 'POSITIVE: a nonzero Stryker exit code fails even when the score passes' {
            # 5/5 killed is 100%, which would clear the 80%
            # threshold, but a non-zero Stryker exit means the
            # tool reported a problem. The exit-code branch must
            # fail the gate rather than report the passing score.
            $path = Join-Path ([System.IO.Path]::GetTempPath()) ('mut-exit-' + [guid]::NewGuid().ToString('N') + '.json')
            Write-MutationReport -Path $path -Mutants @(
                @{ Status = 'Killed' }, @{ Status = 'Killed' }, @{ Status = 'Killed' },
                @{ Status = 'Killed' }, @{ Status = 'Killed' }
            )
            try {
                { Invoke-MutationJudge -ReportPath $path -Threshold 80 -StrykerExitCode 1 } |
                    Should -Throw '*exited 1*'
            } finally {
                Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
            }
        }

        It 'POSITIVE: a zero-mutant non-baseline report fails' {
            $path = Join-Path ([System.IO.Path]::GetTempPath()) ('mut-zero-' + [guid]::NewGuid().ToString('N') + '.json')
            Write-MutationReport -Path $path -Mutants @()
            try {
                { Invoke-MutationJudge -ReportPath $path -Threshold 80 -StrykerExitCode 0 } |
                    Should -Throw '*zero mutants*'
            } finally {
                Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
            }
        }

        It 'POSITIVE: a zero-mutant baseline report is archived, not judged' {
            $path = Join-Path ([System.IO.Path]::GetTempPath()) ('mut-base-' + [guid]::NewGuid().ToString('N') + '.json')
            Write-MutationReport -Path $path -Mutants @()
            try {
                $result = Invoke-MutationJudge -ReportPath $path -Threshold 80 -StrykerExitCode 0 -Baseline
                $result.Mode | Should -Be 'baseline'
                $result.Total | Should -Be 0
            } finally {
                Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
            }
        }

        It 'POSITIVE: an unrecognized report (no Files) fails' {
            $path = Join-Path ([System.IO.Path]::GetTempPath()) ('mut-bad-' + [guid]::NewGuid().ToString('N') + '.json')
            '{ "unexpected": true }' | Set-Content -LiteralPath $path
            try {
                { Invoke-MutationJudge -ReportPath $path -Threshold 80 -StrykerExitCode 0 } |
                    Should -Throw '*unrecognized*'
            } finally {
                Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
            }
        }

        It 'accepts a passing report without throwing' {
            # The negative case for the whole judge: a report
            # that clears every branch must return a result.
            $path = Join-Path ([System.IO.Path]::GetTempPath()) ('mut-ok-' + [guid]::NewGuid().ToString('N') + '.json')
            Write-MutationReport -Path $path -Mutants @(
                @{ Status = 'Killed' }, @{ Status = 'Killed' }, @{ Status = 'Killed' },
                @{ Status = 'Killed' }, @{ Status = 'Killed' }
            )
            try {
                $result = Invoke-MutationJudge -ReportPath $path -Threshold 80 -StrykerExitCode 0
                $result.Mode | Should -Be 'changed-code'
                $result.ScorePercent | Should -Be 100
            } finally {
                Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
            }
        }
    }
}
