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
        (Test-Path -LiteralPath $scriptPath) | Should -BeTrue
    }

    Context 'isolated execution (replaces source-text assertions)' {
        BeforeEach {
            $script:tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid())
            $script:harness  = Join-Path $script:tempRoot 'scripts'
            # `New-Item` does not create intermediate parents reliably,
            # so create $tempRoot explicitly before any child of it.
            New-Item -ItemType Directory -Path $script:tempRoot -Force | Out-Null
            New-Item -ItemType Directory -Path $script:harness  -Force | Out-Null

            # The script does `git -C $repoRoot rev-parse --verify <hash>`,
            # where $repoRoot = Split-Path -Parent $PSScriptRoot, so the
            # harness must be a real git repo with at least one commit.
            git -C $script:tempRoot init -q | Out-Null
            git -C $script:tempRoot config user.email 'test@example.com' | Out-Null
            git -C $script:tempRoot config user.name  'test'              | Out-Null
            'init' | Set-Content -LiteralPath (Join-Path $script:tempRoot 'README.md') -Encoding utf8
            git -C $script:tempRoot add README.md | Out-Null
            git -C $script:tempRoot commit -q -m 'init' | Out-Null
            $script:realHash = (& git -C $script:tempRoot rev-parse HEAD).Trim()

            # Roadmap file with R0.8 and R1.0 so the id-token check passes.
            $roadmap = Join-Path $script:tempRoot 'ROADMAP.md'
            @'
# Roadmap
| R0.8 | foo |
| R1.0 | bar |
'@ | Set-Content -LiteralPath $roadmap -Encoding utf8

            # Copy the script unchanged into the harness; its $PSScriptRoot
            # will be $script:harness and $repoRoot = $script:tempRoot.
            Copy-Item $scriptPath $script:harness

            # Helper defined at script scope so Pester's `It` blocks can
            # call it directly without a `$script:` indirection.
            function Invoke-CheckStatusHarness {
                param([string]$StatusBody)
                $statusFile = Join-Path $script:tempRoot 'STATUS.md'
                $statusBody | Set-Content -LiteralPath $statusFile -Encoding utf8
                $outFile = Join-Path $script:tempRoot 'out.txt'
                $errFile = Join-Path $script:tempRoot 'err.txt'
                $proc = Start-Process -FilePath pwsh -ArgumentList @(
                    '-NoProfile','-File',(Join-Path $script:harness 'check-status.ps1'),
                    '-StatusPath','STATUS.md',
                    '-RoadmapPath','ROADMAP.md'
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

        It 'exits 0 on a clean status file (valid hash, existing path, known roadmap id)' {
            # All three checks pass: real commit hash, real on-disk file,
            # R0.8 in the roadmap.  The "path" token here is STATUS.md
            # itself, which exists at $repoRoot/STATUS.md.
            $body = @"
# Status

References the commit ``$($script:realHash)`` and the file ``STATUS.md`` and the roadmap id ``R0.8``.
"@
            $r = Invoke-CheckStatusHarness $body
            $r.Exit   | Should -Be 0
            $r.Output | Should -Match 'OK'
        }

        It 'exits 1 when STATUS references a commit hash that does not resolve' {
            $badHash = '0000000000000000000000000000000000000000'
            $body = @"
# Status

References the bogus commit ``$badHash``.
"@
            $r = Invoke-CheckStatusHarness $body
            $r.Exit   | Should -Not -Be 0
            $r.Output | Should -Match "STATUS references commit '$badHash'"
        }

        It 'exits 1 when STATUS references a repo path that does not exist' {
            $body = @"
# Status

References the missing file ``nonexistent/path.md``.
"@
            $r = Invoke-CheckStatusHarness $body
            $r.Exit   | Should -Not -Be 0
            $r.Output | Should -Match "STATUS references path 'nonexistent/path.md'"
        }

        It 'exits 1 when STATUS references a roadmap id that is not in the roadmap' {
            $body = @"
# Status

References roadmap item ``R9.9`` which is absent.
"@
            $r = Invoke-CheckStatusHarness $body
            $r.Exit   | Should -Not -Be 0
            $r.Output | Should -Match "STATUS references roadmap item 'R9.9'"
        }

        It 'exits 1 when STATUS references a path that resolves outside the repository' {
            # Behavioural replacement for the removed text-level tripwire in
            # CheckStatusScriptTests: the containment check must reject '..'
            # traversal with the documented message.
            $body = @"
# Status

References the file ``..\outside.md``.
"@
            $r = Invoke-CheckStatusHarness $body
            $r.Exit   | Should -Not -Be 0
            $r.Output | Should -Match 'resolves outside the repository'
        }

        It 'exits 1 when STATUS references an absolute path outside the repository' {
            $body = @"
# Status

References an absolute location ``C:\windows\evil.md``.
"@
            $r = Invoke-CheckStatusHarness $body
            $r.Exit | Should -Not -Be 0
        }
    }
}