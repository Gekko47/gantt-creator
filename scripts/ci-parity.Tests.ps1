#requires -Version 7
<#
.SYNOPSIS
    Pester tripwires for CI parity between .github/workflows/ci.yml and
    the scripts/ gate entry points.

.DESCRIPTION
    The workflow must not inline logic that already exists in scripts/,
    and version-sensitive tool pins (Pester / PSScriptAnalyzer /
    actionlint SHA-256) must come from scripts/tool-versions.psd1 as a
    single source of truth. The defect class prevented here is "local
    says PASS, CI says FAIL because the two views diverged":
      - W8: ci.yml inlined `Invoke-Pester -Script` (Pester 4 param, removed
        in Pester 5); the first CI run would have failed even though the
        local gate was green. Now: no inline Pester, no inline dotnet
        test/build/publish (all delegate to a scripts/*.ps1 entry point).
      - W11: actionlint SHA-256 and PSScriptAnalyzer version were each
        declared in two places (ci.yml + a local script); one was patched
        without the other. Now both pull from tool-versions.psd1.

    Exit 0 on clean; Pester reports failures to the runner.
#>

BeforeAll {
    $repoRoot = Split-Path -Parent $PSScriptRoot
    $ciPath = Join-Path $repoRoot '.github\workflows\ci.yml'
    $script:ciText = Get-Content -LiteralPath $ciPath -Raw

    # Extracts the body of one workflow step (from its `- name:` line to the
    # next `- name:` line) so assertions are scoped to a single step instead
    # of matching anywhere in the file.
    function Get-CiStepBlock {
        param([string]$Text, [string]$StepName)
        $lines = $Text -split "`r?`n"
        $start = -1
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match '^\s*- name:\s*(.+?)\s*$') {
                if ($Matches[1] -eq $StepName) { $start = $i; break }
            }
        }
        if ($start -lt 0) { return $null }
        $body = New-Object System.Collections.Generic.List[string]
        for ($i = $start + 1; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match '^\s*- name:') { break }
            $null = $body.Add($lines[$i])
        }
        return ($body -join "`n")
    }
}

Describe 'ci.yml parity tripwires (W8)' {
    It 'has a Script lint (Pester) step' {
        $block = Get-CiStepBlock -Text $script:ciText -StepName 'Script lint (Pester)'
        $block | Should -Not -BeNullOrEmpty
    }

    It 'script lint step delegates to scripts/test-scripts.ps1 (verify-quick parity)' {
        $block = Get-CiStepBlock -Text $script:ciText -StepName 'Script lint (Pester)'
        $block | Should -Match 'test-scripts\.ps1'
    }

    It 'workflow never calls Invoke-Pester inline' {
        # Inline Pester invocation cannot stay in sync with the pinned
        # script gate; the parameter surface differs between Pester 5 and 6.
        $script:ciText | Should -Not -Match 'Invoke-Pester'
    }

    It 'format step delegates to dotnet format (allow-listed: native MSBuild command, not a script duplicate)' {
        $block = Get-CiStepBlock -Text $script:ciText -StepName 'Format check (production only; tests/ tolerated per tests/Directory.Build.props)'
        $block | Should -Not -BeNullOrEmpty
        $block | Should -Match 'dotnet format'
    }

    It 'workflow does not inline dotnet test/build/publish (must call scripts/*.ps1 or a vetted native command)' {
        # Verify scripts are the script-gate source of truth; the workflow
        # must call them, not re-implement the commands. The build, AddIn
        # publish, and non-Office test commands live in
        # scripts/build-release.ps1, scripts/publish-addin.ps1, and
        # scripts/test-non-office.ps1, shared with verify-quick.ps1 and
        # verify.ps1 (the W11/Phase C pairing completed here).
        # The tripwire is the full negative-match contract: the only
        # inline `run: dotnet <command>` lines allowed in ci.yml are the
        # vetted native MSBuild commands `dotnet format` and
        # `dotnet restore --locked-mode` (see docs/05-GIT-QUALITY.md
        # "CI parity (W8)"). Any other inline dotnet gate command -- a
        # `dotnet test` with or without trailing arguments, `dotnet
        # build`, `dotnet publish`, or a future gate command -- is a
        # violation, whether it sits on a single `run:` line, a YAML
        # sequence `- run: dotnet ...` line, or inside a multiline
        # `run: |` block. (?mi) makes ^ match each line start so the
        # tripwire sees the command at any indentation; `\s*`/`\s+`
        # consume extra whitespace, `-?\s*` accepts a YAML sequence
        # dash before the `run:` label, and command-line options are
        # irrelevant because the match stops at the command word.
        $script:ciText | Should -Not -Match '(?mi)^\s*-?\s*(?:run:\s*)?dotnet\s+(?!format\b|restore\b)[a-z][a-z0-9-]*\b'
    }

    It 'inline dotnet gate regex fires on deliberate violations and not on vetted commands (positive control)' {
        # Guards the negative tripwire above: if the regex cannot detect
        # a single-line `run: dotnet test ...` with trailing args, the
        # YAML sequence form `- run: dotnet test ...`, or an indented
        # `dotnet test` inside a multiline block, the anti-inline
        # assertion could never fire and an inline slip in ci.yml would
        # silently pass. The vetted commands must, conversely, not fire.
        $single = @'
steps:
  - name: Test
    run: dotnet test GanttCreator.slnx -c Release --filter 'Category!=OfficeIntegration'
'@
        $single | Should -Match '(?mi)^\s*-?\s*(?:run:\s*)?dotnet\s+(?!format\b|restore\b)[a-z][a-z0-9-]*\b'

        $dashed = @'
steps:
  - name: Test
    - run: dotnet test GanttCreator.slnx -c Release --no-build
'@
        $dashed | Should -Match '(?mi)^\s*-?\s*(?:run:\s*)?dotnet\s+(?!format\b|restore\b)[a-z][a-z0-9-]*\b'

        $block = @'
steps:
  - name: Test
    run: |
      Set-Location "$env:GITHUB_WORKSPACE"
      dotnet test GanttCreator.slnx --no-build
'@
        $block | Should -Match '(?mi)^\s*-?\s*(?:run:\s*)?dotnet\s+(?!format\b|restore\b)[a-z][a-z0-9-]*\b'

        'run: dotnet restore --locked-mode' | Should -Not -Match '(?mi)^\s*-?\s*(?:run:\s*)?dotnet\s+(?!format\b|restore\b)[a-z][a-z0-9-]*\b'
        'run: dotnet format GanttCreator.slnx --verify-no-changes --exclude tests' | Should -Not -Match '(?mi)^\s*-?\s*(?:run:\s*)?dotnet\s+(?!format\b|restore\b)[a-z][a-z0-9-]*\b'
        '- run: dotnet restore --locked-mode' | Should -Not -Match '(?mi)^\s*-?\s*(?:run:\s*)?dotnet\s+(?!format\b|restore\b)[a-z][a-z0-9-]*\b'
        '- run: dotnet format GanttCreator.slnx --verify-no-changes --exclude tests' | Should -Not -Match '(?mi)^\s*-?\s*(?:run:\s*)?dotnet\s+(?!format\b|restore\b)[a-z][a-z0-9-]*\b'
    }

    It 'test step delegates to scripts/test-non-office.ps1 (verify-quick parity)' {
        # The non-Office dotnet test command is version-sensitive and must
        # not be re-implemented inline in the workflow; the step points at
        # the same entry point verify-quick.ps1 step 12 runs.
        $block = Get-CiStepBlock -Text $script:ciText -StepName 'Test (OfficeIntegration excluded)'
        $block | Should -Not -BeNullOrEmpty
        $block | Should -Match 'test-non-office\.ps1'
    }

    It 'step extractor isolates exactly one step block (positive control)' {
        # Guards the helper itself: a block must stop at the next - name:
        # line, otherwise the delegation assertions could pass on a
        # neighbouring step's body.
        $fragment = @'
steps:
  - name: First
    run: echo one
  - name: Second
    run: echo two
'@
        $block = Get-CiStepBlock -Text $fragment -StepName 'First'
        $block | Should -Match 'echo one'
        $block | Should -Not -Match 'echo two'
    }
}

Describe 'PSScriptAnalyzer gate delegates to Invoke-PssaGate (W13)' {
    It 'Script analyzer step routes through Invoke-PssaGate, not -EnableExit' {
        # W13: -EnableExit's function-level `exit` is swallowed by the
        # Tee/ForEach pipeline in Invoke-Step. The verify scripts route
        # through Invoke-PssaGate (scripts/verify-helpers.ps1); CI must do the
        # same or it re-introduces the W13 defect via a divergent path.
        $block = Get-CiStepBlock -Text $script:ciText -StepName 'Script analyzer (PSScriptAnalyzer)'
        $block | Should -Not -BeNullOrEmpty
        $block | Should -Match 'Invoke-PssaGate'
        # Strip comment lines (full-line # comments) before asserting the
        # gate is not invoked via -EnableExit: the step comment legitimately
        # names the forbidden switch as the reason for the delegation.
        $codeLines = ($block -split "`n") | Where-Object { $_ -notmatch '^\s*#' }
        ($codeLines -join "`n") | Should -Not -Match '-EnableExit'
    }
}

Describe 'tool-versions.psd1 single source of truth (W11)' {
    BeforeAll {
        $versionsPath = Join-Path $repoRoot 'scripts\tool-versions.psd1'
        $script:versions = Import-PowerShellDataFile -LiteralPath $versionsPath
        $script:lintCiText = Get-Content -LiteralPath (Join-Path $repoRoot 'scripts\lint-ci.ps1') -Raw
    }

    It 'pins actionlint version + SHA-256' {
        $script:versions.actionlint.Version | Should -Be '1.7.7'
        $script:versions.actionlint.Sha256 | Should -Match '^[0-9a-f]{64}$'
    }

    It 'actionlint SHA-256 in ci.yml equals the value in tool-versions.psd1' {
        $localHash = $script:versions.actionlint.Sha256
        $ciHash = [regex]::Match($script:ciText, '[0-9a-f]{64}').Value
        $ciHash | Should -Be $localHash
    }

    It 'actionlint SHA-256 in scripts/lint-ci.ps1 is sourced from tool-versions.psd1' {
        # W11: lint-ci.ps1 reads the pin from the psd1 at runtime; the
        # script source no longer carries the literal. The integrity
        # assertion follows the same indirection: both consumers read the
        # same property of the same psd1. $script:lintCiText is already
        # initialized by this Describe block's BeforeAll.
        $script:lintCiText | Should -Match 'Import-PowerShellDataFile'
        $script:lintCiText | Should -Match 'actionlint\.Sha256'
    }

    It 'ci.yml reads PSScriptAnalyzer version from scripts/tool-versions.psd1' {
        # W11: the version is consumed from the single source. The literal
        # "-RequiredVersion 1.25.0" is no longer in the workflow; what must
        # be true is that the workflow imports the psd1 and uses the
        # version variable.
        $script:ciText | Should -Match 'Import-PowerShellDataFile.*tool-versions\.psd1'
        $script:ciText | Should -Match 'PSScriptAnalyzer -RequiredVersion \$pssaVersion'
    }
}