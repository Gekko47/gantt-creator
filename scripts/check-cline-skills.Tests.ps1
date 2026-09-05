#requires -Version 7
<#
.SYNOPSIS
    Pester tests for check-cline-skills.ps1
#>

BeforeAll {
    $repoRoot = Split-Path -Parent $PSScriptRoot
    $scriptPath = Join-Path $repoRoot 'scripts\check-cline-skills.ps1'
}

Describe 'check-cline-skills.ps1' {
    It 'exists and is readable' {
        (Test-Path -LiteralPath $scriptPath) | Should -BeTrue
    }

    It 'implements Phase 2: re-runs sync into a temp dir and byte-compares' {
        # The Phase 2 implementation must (a) re-run sync-cline-skills.ps1
        # with a -SkillsRoot override, then (b) compare generated content
        # against the committed view using git diff --no-index --quiet
        # (exit 0 = same, exit 1 = different).
        $content = Get-Content -LiteralPath $scriptPath -Raw
        $content | Should -Match 'sync-cline-skills\.ps1'
        $content | Should -Match 'SkillsRoot'
        $content | Should -Not -Match 'RulesRoot'
        $content | Should -Match 'git.*diff.*--no-index --quiet'
    }

    It 'documents the temp-dir cleanup on failure' {
        # The finally block must remove the temp directory even on
        # failure; this is defence-in-depth against /tmp filling up
        # when the gate runs in CI thousands of times.
        (Get-Content -LiteralPath $scriptPath -Raw) | Should -Match 'finally'
        (Get-Content -LiteralPath $scriptPath -Raw) | Should -Match 'Remove-Item -LiteralPath \$tmp'
    }
}