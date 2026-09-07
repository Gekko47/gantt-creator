#requires -Version 7
<#
.SYNOPSIS
    Behavioural Pester tests for check-md-links.ps1 (isolated child-pwsh
    harness, same pattern as check-status.Tests.ps1).
#>

Describe 'check-md-links.ps1' {
    BeforeAll { $script:scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'scripts\check-md-links.ps1' }

    It 'exists and is readable' {
        (Test-Path -LiteralPath $script:scriptPath) | Should -BeTrue
    }

    Context 'isolated execution' {
        BeforeEach {
            $script:tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid())
            $script:harness  = Join-Path $script:tempRoot 'scripts'
            New-Item -ItemType Directory -Path $script:harness -Force | Out-Null
            New-Item -ItemType Directory -Path (Join-Path $script:tempRoot 'docs') -Force | Out-Null

            # Copy the script unchanged into the harness; its $PSScriptRoot
            # is $script:harness, so its repo root is $script:tempRoot and it
            # scans $script:tempRoot\docs. The default .github and AGENTS.md
            # roots do not exist in the harness and are skipped.
            Copy-Item $script:scriptPath $script:harness

            function Invoke-MdLinksHarness {
                $outFile = Join-Path $script:tempRoot 'out.txt'
                $errFile = Join-Path $script:tempRoot 'err.txt'
                $proc = Start-Process -FilePath pwsh -ArgumentList @(
                    '-NoProfile', '-File', (Join-Path $script:harness 'check-md-links.ps1')
                ) -NoNewWindow -Wait -PassThru `
                    -WorkingDirectory $script:tempRoot `
                    -RedirectStandardOutput $outFile -RedirectStandardError $errFile
                $combined = (Get-Content -LiteralPath $outFile -Raw) + (Get-Content -LiteralPath $errFile -Raw)
                return [pscustomobject]@{ Exit = $proc.ExitCode; Output = $combined }
            }
        }

        AfterEach {
            if ($script:tempRoot -and (Test-Path -LiteralPath $script:tempRoot)) {
                Remove-Item -LiteralPath $script:tempRoot -Recurse -Force -ErrorAction SilentlyContinue
            }
        }

        It 'exits 0 when all links resolve (plain and ./ relative)' {
            Set-Content -LiteralPath (Join-Path $script:tempRoot 'docs\b.md') -Value 'target' -Encoding utf8
            Set-Content -LiteralPath (Join-Path $script:tempRoot 'docs\a.md') -Value '[b](b.md) and [c](./b.md)' -Encoding utf8
            $r = Invoke-MdLinksHarness
            $r.Exit   | Should -Be 0
            $r.Output | Should -Match 'OK'
        }

        It 'exits 1 on a broken relative link' {
            Set-Content -LiteralPath (Join-Path $script:tempRoot 'docs\a.md') -Value '[missing](missing.md)' -Encoding utf8
            $r = Invoke-MdLinksHarness
            $r.Exit   | Should -Not -Be 0
            $r.Output | Should -Match 'BROKEN LINKS'
            $r.Output | Should -Match 'missing\.md'
        }

        It 'validates ./relative links (positive test for the leading-dot policy)' {
            # Before the policy change, a leading dot exempted a link from the
            # check entirely (false clean). './missing.md' must now fail.
            Set-Content -LiteralPath (Join-Path $script:tempRoot 'docs\a.md') -Value '[dot](./missing.md)' -Encoding utf8
            $r = Invoke-MdLinksHarness
            $r.Exit   | Should -Not -Be 0
            $r.Output | Should -Match '\./missing\.md'
        }
    }
}