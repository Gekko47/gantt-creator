#requires -Version 7
<#
.SYNOPSIS
    Pester tests for check-status.ps1
#>

BeforeAll {
    $repoRoot = Split-Path -Parent $PSScriptRoot
    $scriptPath = Join-Path $repoRoot 'scripts\check-status.ps1'
}

Describe 'check-status.ps1' {
    It 'exists and is readable' {
        (Test-Path $scriptPath) | Should -BeTrue
    }

    It 'contains GetRelativePath for path containment' {
        (Get-Content -LiteralPath $scriptPath -Raw) | Should -Match 'GetRelativePath'
    }

    It 'does NOT contain StartsWith on repoRootFull' {
        (Get-Content -LiteralPath $scriptPath -Raw) | Should -Not -Match '\.StartsWith\(\$repoRootFull'
    }

    It 'rejects wildcards in candidate paths' {
        (Get-Content -LiteralPath $scriptPath -Raw) | Should -Match 'Test-Path.*-LiteralPath'
    }

    It 'uses git -C for repo-scoped git commands' {
        (Get-Content -LiteralPath $scriptPath -Raw) | Should -Match 'git -C \$repoRoot'
    }
}