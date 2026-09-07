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

    It 'workflow does not inline dotnet test (must call scripts/verify-quick.ps1 or verify.ps1 instead)' {
        # Verify scripts are the script-gate source of truth; the workflow
        # must call them, not re-implement the test invocation. The two
        # test steps (Test/OfficeIntegration-excluded and Test/OfficeIntegration)
        # are paired with verify scripts in W11/Phase C; for now the
        # tripwire is the broader 'no bare dotnet test outside verify' rule.
        $script:ciText | Should -Not -Match '^\s*run:\s*dotnet test\s*$'
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

    It 'actionlint SHA-256 in scripts/lint-ci.ps1 equals the value in tool-versions.psd1' {
        $localHash = $script:versions.actionlint.Sha256
        $ciHash = [regex]::Match($script:lintCiText, '[0-9a-f]{64}').Value
        $ciHash | Should -Be $localHash
    }

    It 'ci.yml installs PSScriptAnalyzer at the pinned version' {
        $pinned = $script:versions.PSScriptAnalyzer.Version
        $script:ciText | Should -Match "PSScriptAnalyzer -RequiredVersion $pinned"
    }
}