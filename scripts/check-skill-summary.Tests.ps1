#requires -Version 7
<#
.SYNOPSIS
    Pester tests for check-skill-summary.ps1
#>

BeforeAll {
    $repoRoot = Split-Path -Parent $PSScriptRoot
    $scriptPath = Join-Path $repoRoot 'scripts\check-skill-summary.ps1'
}

Describe 'check-skill-summary.ps1' {
    It 'exists and is readable' {
        (Test-Path $scriptPath) | Should -BeTrue
    }

    It 'uses case-sensitive -cnotmatch for phrase validation' {
        (Get-Content -LiteralPath $scriptPath -Raw) | Should -Match '-cnotmatch'
    }

    It 'asserts canonical phrases from 03-roadmap' {
        (Get-Content -LiteralPath $scriptPath -Raw) | Should -Match 'R0\.8 note'
    }

    It 'asserts canonical phrases from 02-architecture' {
        (Get-Content -LiteralPath $scriptPath -Raw) | Should -Match 'additionally produces'
    }

    It 'asserts canonical phrases from 04-test-strategy' {
        (Get-Content -LiteralPath $scriptPath -Raw) | Should -Match 'set `CurrentCulture`'
    }
}