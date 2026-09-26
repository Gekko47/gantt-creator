#requires -Version 7
<#
.SYNOPSIS
    Pester tests for check-status.ps1
#>

Describe 'check-status.ps1' {
    BeforeAll { $script:scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'scripts\check-status.ps1' }

    It 'exists and is readable' {
        (Test-Path -LiteralPath $script:scriptPath) | Should -BeTrue
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
            Copy-Item $script:scriptPath $script:harness

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
            # R0.8 in the roadmap. The "path" token must contain a directory
            # separator for the path rule to engage, and the file must exist
            # in the harness repository before the gate runs. It is also
            # committed, because the path check requires a tracked file (a
            # clean CI checkout has nothing else) -- not merely one on disk.
            New-Item -ItemType Directory -Path (Join-Path $script:tempRoot 'docs') -Force | Out-Null
            'exists' | Set-Content -LiteralPath (Join-Path $script:tempRoot 'docs\exists.md') -Encoding utf8
            git -C $script:tempRoot add docs/exists.md | Out-Null
            git -C $script:tempRoot commit -q -m 'add docs/exists.md' | Out-Null
            $body = @"
# Status

References the commit ``$($script:realHash)`` and the file ``docs\exists.md`` and the roadmap id ``R0.8``.
"@
            $r = Invoke-CheckStatusHarness $body
            $r.Exit   | Should -Be 0
            $r.Output | Should -Match 'OK'
        }

        It 'exits 1 when STATUS references a path that exists on disk but is untracked by git' {
            # Regression (CI divergence, 2026-09-26): docs/STATUS.md claimed
            # `docs/GanttCreator_StageInspect_Code_Audit.md`, a git-ignored
            # working document. Test-Path passed locally, so pre-commit and
            # verify-quick were green, but a clean CI checkout never had the
            # file and the gate failed there. The path check must require git
            # to know the path, not merely the filesystem.
            $untracked = Join-Path $script:tempRoot 'docs\untracked.md'
            # The harness BeforeEach creates only $tempRoot and $tempRoot\scripts,
            # so the docs directory must be created here before writing into it.
            New-Item -ItemType Directory -Path (Join-Path $script:tempRoot 'docs') -Force | Out-Null
            'present but never added' | Set-Content -LiteralPath $untracked -Encoding utf8
            # Positive control: the file really is on disk, so the violation
            # below can only come from the tracking requirement.
            (Test-Path -LiteralPath $untracked) | Should -BeTrue
            (& git -C $script:tempRoot ls-files --error-unmatch -- 'docs/untracked.md' 2>$null) | Should -BeNullOrEmpty

            $body = @"
# Status

References the file ``docs\untracked.md``.
"@
            $r = Invoke-CheckStatusHarness $body
            $r.Exit   | Should -Not -Be 0
            # Literal substring, not a regex: the token carries a backslash and
            # '\u' is an invalid .NET regex escape, and the gate echoes the token
            # verbatim rather than a git-normalised path.
            $r.Output.Contains("STATUS references path 'docs\untracked.md' which is not tracked by git") | Should -BeTrue
            # It must NOT be reported as merely nonexistent: the file is on disk.
            $r.Output.Contains("STATUS references path 'docs\untracked.md' which does not exist") | Should -BeFalse
        }

        It 'exits 0 for a path that exists on disk and is tracked by git' {
            # Positive control for the tracking requirement: a committed file
            # must still pass, or the new rule would reject every real path.
            New-Item -ItemType Directory -Path (Join-Path $script:tempRoot 'docs') -Force | Out-Null
            'committed' | Set-Content -LiteralPath (Join-Path $script:tempRoot 'docs\tracked.md') -Encoding utf8
            git -C $script:tempRoot add docs/tracked.md | Out-Null
            git -C $script:tempRoot commit -q -m 'add tracked doc' | Out-Null
            (& git -C $script:tempRoot ls-files --error-unmatch -- 'docs/tracked.md').Trim() | Should -Be 'docs/tracked.md'

            $body = @"
# Status

References the file ``docs\tracked.md``.
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

        It 'exits 1 when STATUS references a generated artifact file' {
            $body = @"
# Status

References the local artifact ``scripts/_artifacts/office-evidence/office-20260923-223234842.trx``.
"@
            $r = Invoke-CheckStatusHarness $body
            $r.Exit   | Should -Not -Be 0
            $r.Output | Should -Match "STATUS references generated/ignored artifact path 'scripts/_artifacts/office-evidence/office-20260923-223234842.trx'"
        }

        It 'exits 1 when STATUS references a generated artifact directory' {
            $body = @"
# Status

References the local artifact directory ``scripts/_artifacts/mutation/``.
"@
            $r = Invoke-CheckStatusHarness $body
            $r.Exit   | Should -Not -Be 0
            $r.Output | Should -Match "STATUS references generated/ignored artifact path 'scripts/_artifacts/mutation/'"
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

        It 'exits 1 when STATUS references a POSIX-style parent path that resolves outside the repository' {
            # Same containment contract as the backslash form: the normalised
            # relative-path check must reject '../' traversal too, on whatever
            # host the gate runs (the check must not depend on the platform
            # separator token used in the status file).
            $body = @"
# Status

References the file ``../outside.md``.
"@
            $r = Invoke-CheckStatusHarness $body
            $r.Exit   | Should -Not -Be 0
            $r.Output | Should -Match 'resolves outside the repository'
        }

        It 'exits 1 when STATUS references a path containing wildcard metacharacters' {
            # Positive test for the wildcard-rejection validator: a path-shaped
            # token (it has a directory separator and an extension) carrying ?
            # or [] must be rejected before any filesystem test, with the
            # documented message, and must not be counted as verified.
            $body = @"
# Status

References the glob-style path ``docs?[a]/file.md``.
"@
            $r = Invoke-CheckStatusHarness $body
            $r.Exit   | Should -Not -Be 0
            # Assert on a literal substring: the token itself carries ? and [],
            # which are regex metacharacters for -Match and wildcard metacharacters
            # for -like, so use String.Contains (literal) instead.
            $r.Output.Contains("STATUS references path 'docs?[a]/file.md' contains wildcard metacharacters") | Should -BeTrue
        }

        It 'exits 0 when a wildcard token is present but is not path-shaped' {
            # The wildcard rule must not fire on non-path tokens (e.g. attribute
            # annotations like [Fact]); the shared predicate must classify them
            # out of the path check entirely so they cannot trigger a false
            # wildcard violation.
            $body = @"
# Status

References the attribute annotation ``[Fact]`` and the roadmap id ``R0.8``.
"@
            $r = Invoke-CheckStatusHarness $body
            $r.Exit   | Should -Be 0
            $r.Output | Should -Match 'OK'
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