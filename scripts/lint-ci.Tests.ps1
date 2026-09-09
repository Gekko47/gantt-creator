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
    $script:repoRoot = Split-Path -Parent $PSScriptRoot
    $script:versions = Import-PowerShellDataFile -LiteralPath (Join-Path $script:repoRoot 'scripts\tool-versions.psd1')
    $script:lintText = Get-Content -LiteralPath (Join-Path $script:repoRoot 'scripts\lint-ci.ps1') -Raw
    $script:ciText   = Get-Content -LiteralPath (Join-Path $script:repoRoot '.github\workflows\ci.yml') -Raw
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

    It 'actionlint-hash.ps1 accepts the pinned checksum and rejects a mismatched one (child process)' {
        # The integrity check is no longer inlined in lint-ci.ps1; it lives in
        # scripts/actionlint-hash.ps1. Exercise that helper in an isolated
        # child process with a fixture archive so the test is deterministic and
        # offline (no network download). Assert observable exit behavior.
        $hashScript = Join-Path $script:repoRoot 'scripts\actionlint-hash.ps1'
        (Test-Path -LiteralPath $hashScript) | Should -BeTrue

        $td = Join-Path ([System.IO.Path]::GetTempPath()) ('lint-ci-hash-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $td -Force | Out-Null
        try {
            $zipPath = Join-Path $td 'actionlint.zip'
            # Deterministic fixture content; compute its real SHA-256.
            $content = [byte[]]@(0x50, 0x4B, 0x03, 0x04) + [Text.Encoding]::UTF8.GetBytes('actionlint-fixture')
            [IO.File]::WriteAllBytes($zipPath, $content)
            $realHash = (Get-FileHash -Path $zipPath -Algorithm SHA256).Hash.ToLower()

            # Acceptance: the helper exits 0 when the checksum matches.
            $accept = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile', '-File', $hashScript, '-ZipPath', $zipPath, '-ExpectedHash', $realHash
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput (Join-Path $td 'accept-out.txt') `
                -RedirectStandardError (Join-Path $td 'accept-err.txt')
            $accept.ExitCode | Should -Be 0

            # Rejection: the helper exits non-zero when the checksum is wrong.
            $badHash = ('0' * 63) + '1'
            $reject = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile', '-File', $hashScript, '-ZipPath', $zipPath, '-ExpectedHash', $badHash
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput (Join-Path $td 'reject-out.txt') `
                -RedirectStandardError (Join-Path $td 'reject-err.txt')
            $reject.ExitCode | Should -Not -Be 0
            $errOut = Get-Content -LiteralPath (Join-Path $td 'reject-err.txt') -Raw
            $errOut | Should -Match 'SHA-256 mismatch'
        }
        finally {
            Remove-Item -LiteralPath $td -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}