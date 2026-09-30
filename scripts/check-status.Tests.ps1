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
                param(
                    [string]$StatusBody,
                    [string]$WorkItemsRelativePath = 'WORK-ITEMS',
                    [string]$WorkingDirectory
                )
                $statusFile = Join-Path $script:tempRoot 'STATUS.md'
                $StatusBody | Set-Content -LiteralPath $statusFile -Encoding utf8
                $outFile = Join-Path $script:tempRoot 'out.txt'
                $errFile = Join-Path $script:tempRoot 'err.txt'
                # The working directory defaults to the harness repository root,
                # which is what every other test here runs from. A test that needs
                # to prove the gate is independent of the caller's location passes
                # an explicit directory instead.
                $cwd = if ($WorkingDirectory) { $WorkingDirectory } else { $script:tempRoot }
                $proc = Start-Process -FilePath pwsh -ArgumentList @(
                    '-NoProfile','-File',(Join-Path $script:harness 'check-status.ps1'),
                    '-StatusPath','STATUS.md',
                    '-RoadmapPath','ROADMAP.md',
                    '-WorkItemsPath',$WorkItemsRelativePath
                ) -NoNewWindow -Wait -PassThru `
                    -WorkingDirectory $cwd `
                    -RedirectStandardOutput $outFile -RedirectStandardError $errFile
                $combined = (Get-Content -LiteralPath $outFile -Raw) + (Get-Content -LiteralPath $errFile -Raw)
                return [pscustomobject]@{ Exit = $proc.ExitCode; Output = $combined }
            }

            # Writes one work item whose only fenced block is the supplied
            # command, so a test states just the command under test. Named
            # Invoke- rather than Set- because PSScriptAnalyzer treats a Set-
            # verb as state-changing and demands a ShouldProcess parameter a
            # test helper has no use for; Invoke- matches the sibling helper.
            function Invoke-WorkItemCommand {
                param([string]$Command)
                $dir = Join-Path $script:tempRoot 'WORK-ITEMS'
                New-Item -ItemType Directory -Path $dir -Force | Out-Null
                @"
# Work item

## Evidence

``````powershell
$Command
``````
"@ | Set-Content -LiteralPath (Join-Path $dir 'R9.9-fixture.md') -Encoding utf8
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

        It 'exits 0 for a tracked path written with a redundant parent segment' {
            # The gate matches the canonical root-relative path, not the raw
            # backticked text. git resolves a pathspec for us -- `ls-files --
            # docs/../docs/tracked.md` prints `docs/tracked.md` -- so comparing the
            # printed name against the raw token failed `-contains` and reported a
            # genuinely tracked path as untracked. Both spellings name the same
            # tracked file, and a backticked path is prose, not a git pathspec.
            New-Item -ItemType Directory -Path (Join-Path $script:tempRoot 'docs') -Force | Out-Null
            'committed' | Set-Content -LiteralPath (Join-Path $script:tempRoot 'docs\tracked.md') -Encoding utf8
            git -C $script:tempRoot add docs/tracked.md | Out-Null
            git -C $script:tempRoot commit -q -m 'add tracked doc' | Out-Null

            # Non-vacuity: the raw token is genuinely not the name git prints.
            $printed = (& git -C $script:tempRoot ls-files -- 'docs/../docs/tracked.md').Trim()
            $printed | Should -Be 'docs/tracked.md'
            $printed | Should -Not -Be 'docs/../docs/tracked.md'

            $body = @"
# Status

References the file ``docs/../docs/tracked.md``.
"@
            $r = Invoke-CheckStatusHarness $body
            $r.Exit   | Should -Be 0
            $r.Output | Should -Match 'OK'
        }

        It 'exits 0 for a tracked directory written with a trailing separator' {
            # A directory token may end in a separator. The tracked-directory test
            # appends its own '/', so a trailing one on the token would have
            # produced 'docs//' and matched nothing -- a false "untracked"
            # rejection for a real project directory.
            New-Item -ItemType Directory -Path (Join-Path $script:tempRoot 'docs') -Force | Out-Null
            'committed' | Set-Content -LiteralPath (Join-Path $script:tempRoot 'docs\tracked.md') -Encoding utf8
            git -C $script:tempRoot add docs/tracked.md | Out-Null
            git -C $script:tempRoot commit -q -m 'add tracked doc' | Out-Null

            $body = @"
# Status

References the directory ``docs/``.
"@
            $r = Invoke-CheckStatusHarness $body
            $r.Exit   | Should -Be 0
            $r.Output | Should -Match 'OK'
        }

        It 'exits 0 when STATUS references a path that is staged but not yet committed' {
            # The gate is a pre-commit hook, so it validates the tree of the commit
            # about to be made -- the index. A new file that STATUS names in the same
            # commit that introduces it must NOT be rejected: that was a false
            # rejection of a correct change, caused by reading HEAD instead. The
            # positive controls below prove the index and HEAD genuinely disagree,
            # so this is not a vacuous assertion.
            New-Item -ItemType Directory -Path (Join-Path $script:tempRoot 'docs') -Force | Out-Null
            $staged = Join-Path $script:tempRoot 'docs\staged-only.md'
            'staged, never committed' | Set-Content -LiteralPath $staged -Encoding utf8
            git -C $script:tempRoot add docs/staged-only.md | Out-Null
            (& git -C $script:tempRoot ls-files --error-unmatch -- 'docs/staged-only.md').Trim()
                | Should -Be 'docs/staged-only.md'
            (& git -C $script:tempRoot ls-tree -r --name-only HEAD -- 'docs/staged-only.md')
                | Should -BeNullOrEmpty
            (Test-Path -LiteralPath $staged) | Should -BeTrue

            $body = @"
# Status

References the file ``docs\staged-only.md``.
"@
            $r = Invoke-CheckStatusHarness $body
            $r.Exit   | Should -Be 0
        }

        It 'exits 1 when STATUS references a path that exists only in the working tree' {
            # The counterpart that must still fail: on disk, but never added to git,
            # so no checkout -- local or CI -- will ever have it.
            New-Item -ItemType Directory -Path (Join-Path $script:tempRoot 'docs') -Force | Out-Null
            $untracked = Join-Path $script:tempRoot 'docs\never-added.md'
            'never added' | Set-Content -LiteralPath $untracked -Encoding utf8
            (Test-Path -LiteralPath $untracked) | Should -BeTrue
            (& git -C $script:tempRoot ls-files --error-unmatch -- 'docs/never-added.md' 2>$null)
                | Should -BeNullOrEmpty

            $body = @"
# Status

References the file ``docs\never-added.md``.
"@
            $r = Invoke-CheckStatusHarness $body
            $r.Exit   | Should -Not -Be 0
            $r.Output.Contains("STATUS references path 'docs\never-added.md' which is not tracked by git") | Should -BeTrue
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

        # The suffix regression. The previous regex was `\bR\d+\.\d+\b`, which
        # cannot match R4.7A -- "7" and "A" are both word characters, so no
        # boundary exists between them. Every letter-suffixed row introduced in
        # roadmap revision 10 would therefore have been invisible to this gate.
        It 'exits 1 when STATUS references a letter-suffixed roadmap id that is absent' {
            @'
# Roadmap
| R0.8 | foo |
| R1.0 | bar |
'@ | Set-Content -LiteralPath (Join-Path $script:tempRoot 'ROADMAP.md') -Encoding utf8

            $body = @"
# Status

References roadmap item ``R4.7A`` which is absent.
"@
            $r = Invoke-CheckStatusHarness $body
            $r.Exit   | Should -Not -Be 0
            $r.Output | Should -Match "STATUS references roadmap item 'R4\.7A' which is absent"
        }

        It 'exits 0 when STATUS references a letter-suffixed roadmap id that is present' {
            @'
# Roadmap
| R0.8 | foo |
| R1.0 | bar |
| R4.7A | bar |
'@ | Set-Content -LiteralPath (Join-Path $script:tempRoot 'ROADMAP.md') -Encoding utf8

            # The guide is required for a suffixed ID, so its presence is part
            # of "this row is properly defined" rather than a separate concern.
            $guide = Join-Path $script:tempRoot 'WORK-ITEMS'
            New-Item -ItemType Directory -Path $guide -Force | Out-Null
            '# Work item' | Set-Content -LiteralPath (Join-Path $guide 'R4.7A-identity-and-hierarchy.md') -Encoding utf8

            $body = @"
# Status

References roadmap item ``R4.7A`` which is present.
"@
            $r = Invoke-CheckStatusHarness $body
            $r.Exit   | Should -Be 0
        }

        # A suffixed row with no guide is a row nobody can implement. This
        # positive test is what keeps the new branch non-vacuous: deleting the
        # `if (-not $guide)` block fails it.
        It 'exits 1 when a referenced letter-suffixed roadmap id has no work-item guide' {
            @'
# Roadmap
| R0.8 | foo |
| R1.0 | bar |
| R4.7A | bar |
'@ | Set-Content -LiteralPath (Join-Path $script:tempRoot 'ROADMAP.md') -Encoding utf8

            $body = @"
# Status

References roadmap item ``R4.7A`` which is present but has no guide.
"@
            $r = Invoke-CheckStatusHarness $body
            $r.Exit   | Should -Not -Be 0
            $r.Output | Should -Match "references suffixed roadmap item 'R4\.7A', which has no work-item guide"
        }

        It 'exits 0 when a referenced letter-suffixed roadmap id has a work-item guide' {
            @'
# Roadmap
| R0.8 | foo |
| R1.0 | bar |
| R4.7A | bar |
'@ | Set-Content -LiteralPath (Join-Path $script:tempRoot 'ROADMAP.md') -Encoding utf8

            $guide = Join-Path $script:tempRoot 'WORK-ITEMS'
            New-Item -ItemType Directory -Path $guide -Force | Out-Null
            '# Work item' | Set-Content -LiteralPath (Join-Path $guide 'R4.7A-identity-and-hierarchy.md') -Encoding utf8

            $body = @"
# Status

References roadmap item ``R4.7A`` which is present and has a guide.
"@
            $r = Invoke-CheckStatusHarness $body
            $r.Exit   | Should -Be 0
        }

        # The guide lookup resolved the guides directory against the caller's
        # working directory rather than the repository root, so every suffixed ID
        # reported "no work-item guide" whenever the gate ran from anywhere but the
        # root. The pre-commit hook and CI both run it from the root, which is why
        # every other test here passed. This runs it from an unrelated directory
        # with the guide directory still reachable only via the repository root.
        It 'finds a suffixed roadmap items guide when run from outside the repository root' {
            @'
# Roadmap
| R0.8 | foo |
| R1.0 | bar |
| R4.7A | bar |
'@ | Set-Content -LiteralPath (Join-Path $script:tempRoot 'ROADMAP.md') -Encoding utf8

            $guide = Join-Path $script:tempRoot 'WORK-ITEMS'
            New-Item -ItemType Directory -Path $guide -Force | Out-Null
            '# Work item' | Set-Content -LiteralPath (Join-Path $guide 'R4.7A-identity-and-hierarchy.md') -Encoding utf8

            # A directory that is NOT the repository root and holds no guides of
            # its own, so a lookup relative to the working directory finds nothing.
            $elsewhere = Join-Path $script:tempRoot 'elsewhere'
            New-Item -ItemType Directory -Path $elsewhere -Force | Out-Null
            # Non-vacuity: the guide really is outside the working directory, and
            # really is reachable from the repository root.
            (Resolve-Path -LiteralPath $guide).Path.StartsWith((Resolve-Path -LiteralPath $elsewhere).Path) | Should -BeFalse
            (Test-Path -LiteralPath (Join-Path $guide 'R4.7A-identity-and-hierarchy.md')) | Should -BeTrue

            $body = @"
# Status

References roadmap item ``R4.7A`` which is present and has a guide.
"@
            $r = Invoke-CheckStatusHarness $body 'WORK-ITEMS' $elsewhere
            $r.Exit   | Should -Be 0
            $r.Output | Should -Not -Match 'no work-item guide'
        }

        # Guards the deliberate scoping: the guide rule is enforced only for
        # suffixed IDs, so pre-existing unsuffixed rows are not retro-required.
        It 'does not require a work-item guide for an unsuffixed roadmap id' {
            @'
# Roadmap
| R0.8 | foo |
| R1.0 | bar |
'@ | Set-Content -LiteralPath (Join-Path $script:tempRoot 'ROADMAP.md') -Encoding utf8

            $body = @"
# Status

References roadmap item ``R1.0`` which is present and has no guide.
"@
            $r = Invoke-CheckStatusHarness $body
            $r.Exit   | Should -Be 0
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

        Context 'work-item evidence commands' {
            # The defect this rule exists for: 73 work items carried an evidence
            # command naming a solution file that does not exist, so the command
            # could not run and the evidence it promised was never produced.

            It 'exits 1 when a work-item evidence command names a path that does not exist' {
                Invoke-WorkItemCommand 'dotnet build src/GanttCreator.slnx /p:Configuration=Release /warnaserror'

                $r = Invoke-CheckStatusHarness "# Status`n"

                $r.Exit | Should -Not -Be 0
                $r.Output.Contains("names 'src/GanttCreator.slnx' in an evidence command") | Should -BeTrue
            }

            It 'exits 0 when the work-item evidence command names a path that exists' {
                # Positive control: a real path must still pass, or the rule would
                # reject every legitimate evidence block.
                'committed' | Set-Content -LiteralPath (Join-Path $script:tempRoot 'GanttCreator.slnx') -Encoding utf8
                Invoke-WorkItemCommand 'dotnet build GanttCreator.slnx /p:Configuration=Release /warnaserror'

                $r = Invoke-CheckStatusHarness "# Status`n"

                $r.Exit | Should -Be 0
            }

            It 'ignores a non-existent path named on a comment line' {
                # A comment is a note, not a command. This is what lets a work item
                # record the path it is deliberately correcting (R3.12 does exactly
                # this) without tripping the gate.
                #
                # The real command names GanttCreator.slnx, which must exist: bare
                # filenames are checked too (that is what catches a command naming a
                # file that cannot resolve), so a missing one would be a genuine
                # violation rather than a comment false positive. The bad path here is
                # the one confined to the comment.
                'committed' | Set-Content -LiteralPath (Join-Path $script:tempRoot 'GanttCreator.slnx') -Encoding utf8
                Invoke-WorkItemCommand "# src/GanttCreator.slnx is wrong; use GanttCreator.slnx`ndotnet build GanttCreator.slnx"

                $r = Invoke-CheckStatusHarness "# Status`n"

                $r.Exit | Should -Be 0
                $r.Output | Should -Not -Match "src/GanttCreator.slnx"
            }

            It 'exits 1 when a work-item evidence command names a bare filename that does not exist' {
                # The bare-filename case: `dotnet build GanttCreator.slnx` names a
                # file at the repository root, so the file argument must be checked
                # rather than skipped. Without this, the most common shape of evidence
                # command is unchecked -- which is exactly how 73 work items came to
                # name a non-existent solution path.
                Invoke-WorkItemCommand 'dotnet build GanttCreator.slnx'

                $r = Invoke-CheckStatusHarness "# Status`n"

                $r.Exit | Should -Not -Be 0
                $r.Output.Contains("names 'GanttCreator.slnx' in an evidence command") | Should -BeTrue
            }

            It 'exits 0 when a work-item evidence command names a URL' {
                # A URL is not a repository path. The `//authority/...` fragment that
                # the tokeniser sees carries separators and would otherwise be joined
                # onto the repo root and tested as a local path that never exists.
                Invoke-WorkItemCommand 'pwsh -c "Invoke-WebRequest https://example.com/build.json"'

                $r = Invoke-CheckStatusHarness "# Status`n"

                $r.Exit | Should -Be 0
                $r.Output | Should -Not -Match 'example\.com'
            }

            It 'ignores a non-existent path named in prose outside a fenced block' {
                # Only commands are checked. Prose may discuss a path that does not
                # exist -- an item describing a defect, for instance.
                $dir = Join-Path $script:tempRoot 'WORK-ITEMS'
                New-Item -ItemType Directory -Path $dir -Force | Out-Null
                @'
# Work item

Earlier revisions named ``src/GanttCreator.slnx``, which does not exist.
'@ | Set-Content -LiteralPath (Join-Path $dir 'R9.9-fixture.md') -Encoding utf8

                $r = Invoke-CheckStatusHarness "# Status`n"

                $r.Exit | Should -Be 0
            }

            It 'exits 0 when a work-item evidence command names an existing bare filename' {
                # A bare filename is resolved against the repository root, so
                # `dotnet format README.md` is checked rather than skipped. The
                # harness repository has a committed README.md, so this token
                # resolves and the command is legal -- the counterpart tests pin
                # the missing-filename and root-resolution behaviour.
                Invoke-WorkItemCommand 'dotnet format README.md'

                $r = Invoke-CheckStatusHarness "# Status`n"

                $r.Exit | Should -Be 0
            }

            It 'exits 0 when the work-items directory does not exist' {
                # The rule is additive: a repository without work items must not
                # fail the status gate for that reason alone.
                $r = Invoke-CheckStatusHarness "# Status`n" 'NO-SUCH-DIRECTORY'

                $r.Exit | Should -Be 0
            }
        }
    }
}
