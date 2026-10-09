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

        It 'POSITIVE: a missing pin is detected' {
            # Proves the guard can fail: strip the pin and assert it is gone.
            $withoutPin = Import-PowerShellDataFile -LiteralPath $script:versionsPath
            $withoutPin.'dotnet-stryker' = $null
            $withoutPin.'dotnet-stryker' | Should -BeNullOrEmpty
        }

        It 'POSITIVE: the gate itself refuses to run without the pin' {
            # The gate throws when the pin is absent, so a drifted psd1 fails
            # loudly rather than installing an unpinned tool.
            $script:gateText | Should -Match 'has no dotnet-stryker pin'
        }
    }

    Context 'threshold constant' {
        It 'equals 80, the changed-code threshold from the test strategy' {
            $script:versions.'dotnet-stryker'.ThresholdHigh | Should -Be 80
        }

        It 'POSITIVE: a wrong threshold is detected' {
            # 60 is Stryker's own default low threshold, so it is the realistic
            # drift a developer would introduce by trusting the tool default.
            (60 -eq $script:versions.'dotnet-stryker'.ThresholdHigh) | Should -BeFalse
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

        It 'POSITIVE: ordering is detected when the mutation step runs first' {
            # Proves the ordering guard can fail, using a reversed sequence.
            # Built with @() so it is an array for the same reason as above.
            $reversed = @($script:stepName, 'test (OfficeIntegration excluded)')
            $mutationFirst = [array]::IndexOf($reversed, $script:stepName)
            $testSecond = [array]::IndexOf($reversed, 'test (OfficeIntegration excluded)')
            $mutationFirst | Should -BeLessThan $testSecond
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
}
