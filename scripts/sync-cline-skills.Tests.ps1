#requires -Version 7
<#
.SYNOPSIS
    Pester tests for sync-cline-skills.ps1
#>

BeforeAll {
    $repoRoot = Split-Path -Parent $PSScriptRoot
    $scriptPath = Join-Path $repoRoot 'scripts\sync-cline-skills.ps1'
}

Describe 'sync-cline-skills.ps1' {
    It 'exists and is readable' {
        (Test-Path $scriptPath) | Should -BeTrue
    }

    It 'defines all 7 canonical source files in $map' {
        (Get-Content -LiteralPath $scriptPath -Raw) | Should -Match '01-ENVIRONMENT\.md'
        (Get-Content -LiteralPath $scriptPath -Raw) | Should -Match '02-ARCHITECTURE\.md'
        (Get-Content -LiteralPath $scriptPath -Raw) | Should -Match '03-ROADMAP\.md'
        (Get-Content -LiteralPath $scriptPath -Raw) | Should -Match '04-TEST-STRATEGY\.md'
        (Get-Content -LiteralPath $scriptPath -Raw) | Should -Match '05-GIT-QUALITY\.md'
        (Get-Content -LiteralPath $scriptPath -Raw) | Should -Match '06-LLM-PROTOCOL\.md'
        (Get-Content -LiteralPath $scriptPath -Raw) | Should -Match '07-GANTT-ENTITY-GUIDE\.md'
    }

    It 'uses $SummaryLines = 80 for skill summary truncation' {
        (Get-Content -LiteralPath $scriptPath -Raw) | Should -Match '\$SummaryLines = 80'
    }

    It 'writes files with -NoNewline -Encoding utf8' {
        (Get-Content -LiteralPath $scriptPath -Raw) | Should -Match '-NoNewline -Encoding utf8'
    }

    It 'validates source files exist before regenerating' {
        (Get-Content -LiteralPath $scriptPath -Raw) | Should -Match 'Canonical source missing'
    }
}