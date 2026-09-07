#requires -Version 7
<#
.SYNOPSIS
    Pester tests for lint-ci.ps1: the actionlint SHA-256 pin in the local
    script must be identical to the one the CI workflow enforces (integrity
    parity). A text assertion is the right shape here: the contract under
    test is that two files carry the same pin.
#>

Describe 'lint-ci.ps1' {
    BeforeAll {
        $repoRoot = Split-Path -Parent $PSScriptRoot
        $script:lintText = Get-Content -LiteralPath (Join-Path $repoRoot 'scripts\lint-ci.ps1') -Raw
        $script:ciText   = Get-Content -LiteralPath (Join-Path $repoRoot '.github\workflows\ci.yml') -Raw
        $script:hex64 = '[0-9a-f]{64}'
    }

    It 'pins an actionlint SHA-256' {
        $script:lintText | Should -Match $script:hex64
    }

    It 'uses the same SHA-256 pin as the CI workflow (anti-drift)' {
        $localHash = [regex]::Match($script:lintText, $script:hex64).Value
        $ciHash = [regex]::Match($script:ciText, $script:hex64).Value
        $localHash | Should -Not -BeNullOrEmpty
        $localHash | Should -Be $ciHash
    }
}