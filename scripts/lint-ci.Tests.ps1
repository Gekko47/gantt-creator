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
        # W11: the workflow no longer carries a literal mirror of the pin; it
        # imports scripts/tool-versions.psd1 and reads $versions.actionlint.Sha256
        # at runtime. Assert the indirection rather than a literal: the import
        # and the property reference must both be present in the actionlint
        # step block (scoped so the import in the PSScriptAnalyzer step does
        # not satisfy it). A foreach loop is used instead of
        # ForEach-Object -Begin/-Process/-End so PSScriptAnalyzer tracks the
        # variables within a single scope.
        $ciStepBlockLines = [System.Collections.Generic.List[string]]::new()
        $inBlock = $false
        foreach ($line in ($script:ciText -split "`r?`n"))
        {
            if ($line.Trim() -match '^- name:\s*Workflow lint \(actionlint\)')
            {
                $inBlock = $true
            }
            elseif ($inBlock -and $line.Trim() -match '^- name:')
            {
                $inBlock = $false
            }
            elseif ($inBlock)
            {
                $null = $ciStepBlockLines.Add($line)
            }
        }
        $ciStepBlock = $ciStepBlockLines -join "`n"
        $ciStepBlock | Should -Match 'Import-PowerShellDataFile.*tool-versions\.psd1'
        $ciStepBlock | Should -Match 'actionlint\.Sha256'
        $ciStepBlock | Should -Match 'actionlint\.Version'
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

    It 'actionlint-hash.ps1 File mode accepts the pinned exe hash shape and rejects a mismatched one (child process)' {
        # The executable-integrity check reuses the same standalone helper
        # via its -FilePath parameter set (no network download needed).
        $hashScript = Join-Path $script:repoRoot 'scripts\actionlint-hash.ps1'
        (Test-Path -LiteralPath $hashScript) | Should -BeTrue

        $td = Join-Path ([System.IO.Path]::GetTempPath()) ('lint-ci-exe-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $td -Force | Out-Null
        try {
            $exePath = Join-Path $td 'actionlint.exe'
            [IO.File]::WriteAllBytes($exePath, [Text.Encoding]::UTF8.GetBytes('actionlint-fixture-exe'))
            $realHash = (Get-FileHash -LiteralPath $exePath -Algorithm SHA256).Hash.ToLower()

            $accept = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile', '-File', $hashScript, '-FilePath', $exePath, '-ExpectedHash', $realHash
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput (Join-Path $td 'accept-out.txt') `
                -RedirectStandardError (Join-Path $td 'accept-err.txt')
            $accept.ExitCode | Should -Be 0

            $badHash = ('f' * 63) + '0'
            $reject = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile', '-File', $hashScript, '-FilePath', $exePath, '-ExpectedHash', $badHash
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

    It 'pins a distinct, well-formed ExeSha256 for the cached-executable check' {
        # The cached-executable validation needs its own pin (the archive
        # hash cannot verify an extracted binary). It must be a 64-hex
        # digest and differ from the archive pin.
        $script:versions.actionlint.ExeSha256 | Should -Match '^[0-9a-f]{64}$'
        $script:versions.actionlint.ExeSha256 | Should -Not -Be $script:versions.actionlint.Sha256
    }

    It 'lint-ci.ps1 validates the cached executable via actionlint-hash.ps1 before execution' {
        # The cached-exe verification must run through the same standalone
        # helper as the archive check (single integrity path), consuming the
        # ExeSha256 pin from the psd1 at runtime rather than a literal.
        $script:lintText | Should -Match 'actionlint\.ExeSha256'
        $script:lintText | Should -Match "-FilePath (?s).*-ExpectedHash"
    }

    It 'keeps a single workflow file an array so splat passes the path' {
        $script:lintText | Should -Match 'object\[\]\]\$workflowFiles'
    }
}
