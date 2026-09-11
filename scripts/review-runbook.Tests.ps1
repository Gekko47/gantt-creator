#requires -Version 7
<#
.SYNOPSIS
    W-7 review runbook tripwires. The end-to-end review methodology
    (docs/06-LLM-PROTOCOL.md "End-to-end code review methodology") requires
    the following artefacts to exist and behave as documented; a future
    contributor who removes or regresses any of them fails this test,
    which forces a discussion rather than a silent regression.

    Exit 0 on clean; Pester reports failures to the runner.
#>

BeforeAll {
    # PSSA cannot flow-analyse a Pester BeforeAll across an It block, so
    # the repo root is held in $script: scope and referenced via the
    # same pattern as the other Pester test files (Phase A W14 lesson).
    $script:repoRoot = Split-Path -Parent $PSScriptRoot
}

Describe 'W-7 review methodology runbook' {
    It 'the canonical methodology section exists in docs/06-LLM-PROTOCOL.md' {
        $path = Join-Path $script:repoRoot 'docs\06-LLM-PROTOCOL.md'
        Test-Path -LiteralPath $path | Should -BeTrue
        $text = Get-Content -LiteralPath $path -Raw
        $text | Should -Match '## End-to-end code review methodology'
        # The six phases must be present; a partial migration is a
        # defect, not a stylistic choice.
        $text | Should -Match '### 1\. Read the source'
        $text | Should -Match '### 2\. Probe empirically'
        $text | Should -Match '### 3\. Identify defect classes'
        $text | Should -Match '### 4\. Plan'
        $text | Should -Match '### 5\. Implement, verify, commit'
        $text | Should -Match '### 6\. Self-application'
    }

    It 'the four defect classes from the W-13 run are named in the methodology' {
        # A runbook is useless if it does not list the defects it covers.
        # The names are stable across the W-13 / Phase A/B/C documentation;
        # if a future refactor renames them, the human reviewer will
        # see the test fail and decide whether the rename is real or
        # accidental.
        $path = Join-Path $script:repoRoot 'docs\06-LLM-PROTOCOL.md'
        $text = Get-Content -LiteralPath $path -Raw
        $text | Should -Match 'Gate integrity'
        $text | Should -Match 'Build pipeline drift'
        $text | Should -Match 'Docs/code drift'
        $text | Should -Match 'Local/CI view divergence'
    }

    It 'W-13 capture-then-judge pattern is documented in 04-test-strategy.md' {
        # The W-13 lesson lives in two places: the runbook (this test)
        # and the strategy skill. A future agent who removes the
        # strategy section without updating the runbook fails here.
        $path = Join-Path $script:repoRoot 'docs\04-TEST-STRATEGY.md'
        $text = Get-Content -LiteralPath $path -Raw
        $text | Should -Match '### Gate integrity \(W-13\)'
        $text | Should -Match 'capture, then judge'
        $text | Should -Match 'No-silent-pass on empty input'
    }

    It 'every gate integrity amendment has a working positive-control test' {
        # The four worked-example tests must exist as files; their absence
        # means the gate they protect is blind. The runbook pins their
        # presence so the next reviewer who needs the methodology can
        # read the actual working code, not just the docs.
        $missing = @()
        foreach ($rel in @(
            'scripts\pssa-gate.Tests.ps1',
            'scripts\step-parity.Tests.ps1',
            'tests\GanttCreator.Architecture.Tests\ArtifactSourceMarkerTests.cs',
            'scripts\ci-parity.Tests.ps1')) {
            $abs = Join-Path $script:repoRoot $rel
            if (-not (Test-Path -LiteralPath $abs)) { $missing += $rel }
        }
        $missing.Count | Should -Be 0 -Because "missing gate runbook tests: $($missing -join ', ')"
    }

    It 'tool-versions.psd1 is the single source of truth (W11) and three consumers exist' {
        # W-11: a tool pin declared in two places is the "local/CI view
        # divergence" defect class. The single source must exist, and
        # at least three consumers must read from it.
        $psd1 = Join-Path $script:repoRoot 'scripts\tool-versions.psd1'
        Test-Path -LiteralPath $psd1 | Should -BeTrue

        $ciPath    = Join-Path $script:repoRoot '.github\workflows\ci.yml'
        $lintPath  = Join-Path $script:repoRoot 'scripts\lint-ci.ps1'
        $testPath  = Join-Path $script:repoRoot 'scripts\test-scripts.ps1'
        foreach ($p in @($ciPath, $lintPath, $testPath)) {
            $text = Get-Content -LiteralPath $p -Raw
            $text | Should -Match 'tool-versions\.psd1' -Because "$p must consume scripts\tool-versions.psd1 (W-11)"
        }
    }

    It 'the W-12 branch policy is in docs/05-GIT-QUALITY.md' {
        # W-12: phases may not exit on local-only evidence. A future
        # contributor who edits the git-quality file to weaken or
        # remove this rule fails this test, which forces a conversation.
        $path = Join-Path $script:repoRoot 'docs\05-GIT-QUALITY.md'
        $text = Get-Content -LiteralPath $path -Raw
        $text | Should -Match 'phase may not exit on local-only evidence'
    }

    It 'no file claims verify-quick.ps1 or verify.ps1 completes in under ~2 minutes' {
        # Docs/code drift class: 2c1519e corrected the verify-quick
        # duration claim in docs/05-GIT-QUALITY.md, but stale copies
        # survived in .vscode/tasks.json, scripts/install-pre-commit.ps1,
        # and scripts/pre-commit.ps1. Every recorded verify-quick run in
        # docs/STATUS.md is in the 97-333 s range, so any sub-minute
        # duration token on a line that mentions one of the verify gates
        # is stale. (The hook's own honest "a few seconds" wording and
        # the "~N minutes" claims do not match this pattern.)
        $offenders = Get-ChildItem -Path $script:repoRoot -Recurse -Include '*.md', '*.ps1', '*.json' -File |
            Where-Object { $_.FullName -notmatch '\\(\.git|\.vs|bin|obj|node_modules|_artifacts)\\' } |
            Select-String -Pattern '~\s?(?:[1-9]|[1-5]\d)\s?s\b[\)\.;,]?' |
            Where-Object { $_.Line -match 'verify-quick|verify\.ps1|verify-full' }
        if ($offenders) {
            $offenders | ForEach-Object {
                Write-Host "  $($_.Path):$($_.LineNumber): $($_.Line.Trim())"
            }
        }
        $offenders | Should -BeNullOrEmpty -Because 'a stale sub-minute duration claim about a verify gate survived; this is the docs/code drift class that already shipped once (2c1519e fixed one copy, four more survived)'
    }
}