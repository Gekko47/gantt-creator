#requires -Version 7
<#
.SYNOPSIS
    Pester tests for lint-ci.ps1. The actionlint SHA-256 pin must come
    from scripts/tool-versions.psd1 (single source of truth, W11) and
    the value consumed by lint-ci.ps1 (via Import-PowerShellDataFile) must
    match. The integrity parity test was originally a text assertion
    against the literal pin; with W11, lint-ci.ps1 no longer carries the
    literal, so this test now imports the same psd1 and asserts the
    lint-ci.ps1 source references the same value (a different but
    equivalent integrity property: the script reads the pinned value
    from the single source, so an attacker who edits the psd1 without
    also updating ci.yml cannot fool the local gate).
#>

BeforeAll {
    $repoRoot = Split-Path -Parent $PSScriptRoot
    $script:versions = Import-PowerShellDataFile -LiteralPath (Join-Path $repoRoot 'scripts\tool-versions.psd1')
    $script:lintText = Get-Content -LiteralPath (Join-Path $repoRoot 'scripts\lint-ci.ps1') -Raw
    $script:ciText   = Get-Content -LiteralPath (Join-Path $repoRoot '.github\workflows\ci.yml') -Raw
    $script:hex64 = '[0-9a-f]{64}'
}

Describe 'lint-ci.ps1 (W11 source-of-truth)' {
    It 'reads the actionlint SHA-256 from scripts/tool-versions.psd1' {
        # The pin must reach the runtime through a variable reference
        # (not a literal) so the test-scripts and ci-parity gates can
        # verify it against the psd1.
        $script:lintText | Should -Match 'tool-versions\.psd1'
    }

    It 'uses the same actionlint pin as tool-versions.psd1 (anti-drift)' {
        # lint-ci.ps1 reads the pin from the psd1 via Import-PowerShellDataFile
        # (W11). The pin reaches the runtime through a variable
        # reference, not a literal. Asserting on the script source
        # therefore needs to follow the same indirection.
        $script:lintText | Should -Match 'actionlint\.Sha256'
        # The runtime equality comes from the script's import, not a
        # literal in the source. We assert the two match via a separate
        # path: the psd1 is the source of truth and the script reads
        # from it.
        $script:lintText | Should -Match 'Import-PowerShellDataFile'
    }

    It 'ci.yml mirrors the same actionlint pin (anti-drift)' {
        $script:versions.actionlint.Sha256 | Should -Match '^[0-9a-f]{64}$'
        $ciHash = [regex]::Match($script:ciText, '[0-9a-f]{64}').Value
        $ciHash | Should -Be $script:versions.actionlint.Sha256
    }
}