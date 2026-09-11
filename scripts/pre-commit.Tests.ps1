#requires -Version 7
<#
.SYNOPSIS
    Pester tests for pre-commit.ps1 and install-pre-commit.ps1
#>

Describe 'pre-commit.ps1' {
    BeforeAll { $script:scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'scripts\pre-commit.ps1' }

    It 'exists and is readable' {
        (Test-Path -LiteralPath $script:scriptPath) | Should -BeTrue
    }

    It 'runs only the fast deterministic gates, not full verify-quick' {
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        # Strip the XML doc comment block (<# ... #>) so mentions in
        # .SYNOPSIS/.DESCRIPTION do not trip the negative match.
        $codeOnly = $raw -replace '(?s)<#.*?#>', ''
        $codeOnly | Should -Match 'check-cline-skills\.ps1'
        $codeOnly | Should -Match 'check-status\.ps1'
        $codeOnly | Should -Match 'check-md-links\.ps1'
        $codeOnly | Should -Not -Match 'check-skill-summary\.ps1'
        # Must not INVOKE verify-quick.ps1 in the executable portion.
        $codeOnly | Should -Not -Match '(pwsh|\$\(|&)\s?.*verify-quick\.ps1'
    }

    It 'fails fast with a non-zero exit and the blocking message' {
        # The script must surface a 'Commit blocked' line on failure.
        # Exit-code propagation is covered by the behavioral tests
        # (Start-Process invocation) later in this file.
        $content = Get-Content -LiteralPath $script:scriptPath -Raw
        $content | Should -Match 'Commit blocked'
    }

    Context 'gate behavior (with stub checkers)' {
        BeforeAll {
            # Helper: write a checker stub that appends its name to the log
            # and optionally exits non-zero. Encoding is utf8NoBOM so the
            # #requires line and PowerShell parsing are happy on Windows.
            function Write-PreCommitStubChecker {
                param(
                    [string] $Path,
                    [string] $Name,
                    [string] $LogPath,
                    [int]    $ExitCode = 0
                )
                $body = @"
`$ErrorActionPreference = 'Continue'
Add-Content -LiteralPath '$LogPath' -Value '$Name'
if ($ExitCode -ne 0) { Write-Host 'simulated failure' }
exit $ExitCode
"@
                Set-Content -LiteralPath $Path -Value $body -Encoding utf8NoBOM
            }

            # Helper: write a checker stub that fails via a terminating
            # Write-Error, the pattern real checkers in this repo use
            # (see check-cline-skills.ps1). Locks in the fix for the
            # $LASTEXITCODE-reset bug where piping through ForEach-Object
            # used to mask this kind of failure.
            function Write-PreCommitTerminatingErrorStubChecker {
                param(
                    [string] $Path,
                    [string] $Name,
                    [string] $LogPath,
                    [int]    $ExitCode = 1
                )
                $body = @"
`$ErrorActionPreference = 'Stop'
Add-Content -LiteralPath '$LogPath' -Value '$Name'
Write-Error 'simulated terminating failure'
exit $ExitCode
"@
                Set-Content -LiteralPath $Path -Value $body -Encoding utf8NoBOM
            }
        }

        BeforeEach {
            $script:tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid())
            $script:baseTempDir = $script:tempDir
            $script:scriptsDir = Join-Path $script:tempDir 'scripts'
            $script:invocationLog = Join-Path $script:tempDir 'invocation.log'
            New-Item -ItemType Directory -Path $script:scriptsDir -Force | Out-Null

            # The dirty-tree guard treats a git failure as a blocking
            # condition, so the stub-checker tests must run inside a real
            # (clean) git repo and exercise the normal checker loop, not
            # rely on the git-failure bypass a non-repo directory would
            # trigger. The repo is clean (empty-tree commit), so the guard
            # passes and the checkers run.
            git -C $script:tempDir init -q
            git -C $script:tempDir config user.email 'test@local'
            git -C $script:tempDir config user.name 'test'
            git -C $script:tempDir commit -q -m 'init' --allow-empty

            # Copy the real pre-commit.ps1 verbatim. We do NOT dot-source it
            # or override its private $checks variable; the test exercises
            # the same code that runs in production, with only the checker
            # scripts replaced by stubs.
            $script:copyScriptPath = Join-Path $script:scriptsDir 'pre-commit.ps1'
            Copy-Item -LiteralPath $script:scriptPath -Destination $script:copyScriptPath -Force
        }

        AfterEach {
            # Tests that reassign $script:tempDir (the fixture-repo cases)
            # orphan the directory BeforeEach created; remove both so neither
            # leaks into the next test or the temp folder.
            foreach ($dir in @($script:tempDir, $script:baseTempDir)) {
                if ($dir -and (Test-Path -LiteralPath $dir)) {
                    Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
        }

        It 'invokes all three checkers in declared order and exits 0 when all pass' {
            Write-PreCommitStubChecker -Path (Join-Path $script:scriptsDir 'check-cline-skills.ps1')  -Name 'check-cline-skills.ps1'   -LogPath $script:invocationLog -ExitCode 0
            Write-PreCommitStubChecker -Path (Join-Path $script:scriptsDir 'check-status.ps1')        -Name 'check-status.ps1'         -LogPath $script:invocationLog -ExitCode 0
            Write-PreCommitStubChecker -Path (Join-Path $script:scriptsDir 'check-md-links.ps1')      -Name 'check-md-links.ps1'       -LogPath $script:invocationLog -ExitCode 0

            # Invoke via Start-Process so all output streams (incl.
            # Write-Host, which Pester's test host otherwise intercepts)
            # land in files we can assert on. The exit code comes back
            # through the process object, avoiding $LASTEXITCODE races.
            $script:stdoutFile = Join-Path $script:tempDir 'stdout.txt'
            $script:stderrFile = Join-Path $script:tempDir 'stderr.txt'
            $script:proc = Start-Process -FilePath pwsh -ArgumentList @('-NoProfile','-File',$script:copyScriptPath) -NoNewWindow -Wait -PassThru -RedirectStandardOutput $script:stdoutFile -RedirectStandardError $script:stderrFile
            $script:exitCode = $script:proc.ExitCode
            $script:output = (Get-Content -LiteralPath $script:stdoutFile -Raw) + (Get-Content -LiteralPath $script:stderrFile -Raw)

            $script:exitCode | Should -Be 0
            $script:output | Should -Match 'pre-commit: quick gates PASS'
            Get-Content -LiteralPath $script:invocationLog | Should -Be @(
                'check-cline-skills.ps1',
                'check-status.ps1',
                'check-md-links.ps1'
            )
        }

        It 'exits non-zero with the blocking message and short-circuits on first failure' {
            # First stub passes; second fails; third is staged but must never run.
            Write-PreCommitStubChecker -Path (Join-Path $script:scriptsDir 'check-cline-skills.ps1')  -Name 'check-cline-skills.ps1'   -LogPath $script:invocationLog -ExitCode 0
            Write-PreCommitStubChecker -Path (Join-Path $script:scriptsDir 'check-status.ps1')        -Name 'check-status.ps1'         -LogPath $script:invocationLog -ExitCode 7
            Write-PreCommitStubChecker -Path (Join-Path $script:scriptsDir 'check-md-links.ps1')      -Name 'check-md-links.ps1'       -LogPath $script:invocationLog -ExitCode 0

            $script:stdoutFile = Join-Path $script:tempDir 'stdout.txt'
            $script:stderrFile = Join-Path $script:tempDir 'stderr.txt'
            $script:proc = Start-Process -FilePath pwsh -ArgumentList @('-NoProfile','-File',$script:copyScriptPath) -NoNewWindow -Wait -PassThru -RedirectStandardOutput $script:stdoutFile -RedirectStandardError $script:stderrFile
            $script:exitCode = $script:proc.ExitCode
            $script:output = (Get-Content -LiteralPath $script:stdoutFile -Raw) + (Get-Content -LiteralPath $script:stderrFile -Raw)

            # Observable gate behavior: any non-zero exit blocks the commit.
            # (The exact exit code is whatever PowerShell reports for a
            # terminating error from the failing checker; asserting the
            # non-zero property is the stable contract.)
            $script:exitCode | Should -Not -Be 0
            $script:output | Should -Match 'Commit blocked'

            # Fail-fast proof: the 2nd checker ran, the 3rd did not.
            Get-Content -LiteralPath $script:invocationLog | Should -Be @(
                'check-cline-skills.ps1',
                'check-status.ps1'
            )
        }

        It 'emits the blocking message when a checker fails via a terminating Write-Error' {
            # Regression guard for the $LASTEXITCODE-reset bug: before the
            # fix, pre-commit.ps1 piped the checker's output through
            # ForEach-Object { Write-Host ... }, which reset $LASTEXITCODE
            # to 0 before the if-branch could fire. The first checker
            # passes; the second fails via a terminating Write-Error (the
            # pattern used by check-cline-skills.ps1 today); the third is
            # staged and must never run.
            Write-PreCommitStubChecker               -Path (Join-Path $script:scriptsDir 'check-cline-skills.ps1')  -Name 'check-cline-skills.ps1'  -LogPath $script:invocationLog -ExitCode 0
            Write-PreCommitTerminatingErrorStubChecker -Path (Join-Path $script:scriptsDir 'check-status.ps1')        -Name 'check-status.ps1'        -LogPath $script:invocationLog -ExitCode 1
            Write-PreCommitStubChecker               -Path (Join-Path $script:scriptsDir 'check-md-links.ps1')      -Name 'check-md-links.ps1'      -LogPath $script:invocationLog -ExitCode 0

            $script:stdoutFile = Join-Path $script:tempDir 'stdout.txt'
            $script:stderrFile = Join-Path $script:tempDir 'stderr.txt'
            $script:proc = Start-Process -FilePath pwsh -ArgumentList @('-NoProfile','-File',$script:copyScriptPath) -NoNewWindow -Wait -PassThru -RedirectStandardOutput $script:stdoutFile -RedirectStandardError $script:stderrFile
            $script:exitCode = $script:proc.ExitCode
            $script:output = (Get-Content -LiteralPath $script:stdoutFile -Raw) + (Get-Content -LiteralPath $script:stderrFile -Raw)

            $script:exitCode | Should -Not -Be 0
            $script:output | Should -Match 'Commit blocked'

            # Fail-fast proof: the 2nd checker ran (and failed), the 3rd did not.
            Get-Content -LiteralPath $script:invocationLog | Should -Be @(
                'check-cline-skills.ps1',
                'check-status.ps1'
            )
        }

        It 'refuses to run when an inspected root has unstaged changes' {
            # Isolated fixture repo so this test does not depend on (or dirty)
            # the real working tree. pre-commit.ps1 computes $repoRoot as the
            # parent of $scriptRoot (= Split-Path $PSCommandPath), so the copied
            # script must live at $repoRoot/scripts/pre-commit.ps1.
            $script:tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid())
            $script:scriptsDir = Join-Path $script:tempDir 'scripts'
            New-Item -ItemType Directory -Path $script:scriptsDir -Force | Out-Null

            # Init a real git repo with a clean baseline: the guard compares
            # working tree vs index, so both must start identical.
            git -C $script:tempDir init -q
            git -C $script:tempDir config user.email 'test@local'
            git -C $script:tempDir config user.name 'test'
            New-Item -ItemType Directory -Path (Join-Path $script:tempDir 'docs') -Force | Out-Null
            New-Item -ItemType Directory -Path (Join-Path $script:tempDir '.cline\skills') -Force | Out-Null
            New-Item -ItemType Directory -Path (Join-Path $script:tempDir '.github') -Force | Out-Null
            Set-Content -LiteralPath (Join-Path $script:tempDir 'AGENTS.md') -Value 'initial'
            Set-Content -LiteralPath (Join-Path $script:tempDir 'docs\x.md') -Value 'x'
            Set-Content -LiteralPath (Join-Path $script:tempDir '.cline\skills\y.md') -Value 'y'
            Set-Content -LiteralPath (Join-Path $script:tempDir '.github\ci.yml') -Value 'z'
            git -C $script:tempDir add -A
            git -C $script:tempDir commit -q -m 'init'

            # Dirty an inspected root WITHOUT staging it.
            Set-Content -LiteralPath (Join-Path $script:tempDir 'AGENTS.md') -Value 'changed'

            # Copy the real pre-commit.ps1 into scripts/ so $PSCommandPath
            # resolves to $repoRoot/scripts/pre-commit.ps1, matching production.
            $script:copyScriptPath = Join-Path $script:scriptsDir 'pre-commit.ps1'
            Copy-Item -LiteralPath $script:scriptPath -Destination $script:copyScriptPath -Force

            $script:stdoutFile = Join-Path $script:tempDir 'stdout.txt'
            $script:stderrFile = Join-Path $script:tempDir 'stderr.txt'
            $script:proc = Start-Process -FilePath pwsh -ArgumentList @('-NoProfile','-File',$script:copyScriptPath) -NoNewWindow -Wait -PassThru -RedirectStandardOutput $script:stdoutFile -RedirectStandardError $script:stderrFile
            $script:exitCode = $script:proc.ExitCode
            $script:output = (Get-Content -LiteralPath $script:stdoutFile -Raw) + (Get-Content -LiteralPath $script:stderrFile -Raw)

            $script:exitCode | Should -Not -Be 0
            $script:output | Should -Match 'Commit blocked'
        }

        It 'runs normally when the working tree is clean under the inspected roots' {
            # Mirror of the dirty case but with NO unstaged change: the guard
            # must NOT fire and the gate must run through to the checkers.
            $script:tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid())
            $script:scriptsDir = Join-Path $script:tempDir 'scripts'
            $script:invocationLog = Join-Path $script:tempDir 'invocation.log'
            New-Item -ItemType Directory -Path $script:scriptsDir -Force | Out-Null

            git -C $script:tempDir init -q
            git -C $script:tempDir config user.email 'test@local'
            git -C $script:tempDir config user.name 'test'
            New-Item -ItemType Directory -Path (Join-Path $script:tempDir 'docs') -Force | Out-Null
            New-Item -ItemType Directory -Path (Join-Path $script:tempDir '.cline\skills') -Force | Out-Null
            New-Item -ItemType Directory -Path (Join-Path $script:tempDir '.github') -Force | Out-Null
            Set-Content -LiteralPath (Join-Path $script:tempDir 'AGENTS.md') -Value 'initial'
            Set-Content -LiteralPath (Join-Path $script:tempDir 'docs\x.md') -Value 'x'
            Set-Content -LiteralPath (Join-Path $script:tempDir '.cline\skills\y.md') -Value 'y'
            Set-Content -LiteralPath (Join-Path $script:tempDir '.github\ci.yml') -Value 'z'
            git -C $script:tempDir add -A
            git -C $script:tempDir commit -q -m 'init'

            # No unstaged change this time. Stub checkers are required because
            # the guard passes and the gate proceeds to the checker loop.
            Write-PreCommitStubChecker -Path (Join-Path $script:scriptsDir 'check-cline-skills.ps1')  -Name 'check-cline-skills.ps1'  -LogPath $script:invocationLog -ExitCode 0
            Write-PreCommitStubChecker -Path (Join-Path $script:scriptsDir 'check-status.ps1')        -Name 'check-status.ps1'        -LogPath $script:invocationLog -ExitCode 0
            Write-PreCommitStubChecker -Path (Join-Path $script:scriptsDir 'check-md-links.ps1')      -Name 'check-md-links.ps1'      -LogPath $script:invocationLog -ExitCode 0

            $script:copyScriptPath = Join-Path $script:scriptsDir 'pre-commit.ps1'
            Copy-Item -LiteralPath $script:scriptPath -Destination $script:copyScriptPath -Force

            $script:stdoutFile = Join-Path $script:tempDir 'stdout.txt'
            $script:stderrFile = Join-Path $script:tempDir 'stderr.txt'
            $script:proc = Start-Process -FilePath pwsh -ArgumentList @('-NoProfile','-File',$script:copyScriptPath) -NoNewWindow -Wait -PassThru -RedirectStandardOutput $script:stdoutFile -RedirectStandardError $script:stderrFile
            $script:exitCode = $script:proc.ExitCode
            $script:output = (Get-Content -LiteralPath $script:stdoutFile -Raw) + (Get-Content -LiteralPath $script:stderrFile -Raw)

            $script:exitCode | Should -Be 0
            $script:output | Should -Match 'pre-commit: quick gates PASS'
            $script:output | Should -Not -Match 'Commit blocked'
        }

        It 'allows the commit when all inspected-root changes are staged (no unstaged/untracked)' {
            # Regression guard for the staged-changes bug: the guard must use
            # `git diff --name-only` (unstaged) + untracked, NOT `git status
            # --porcelain` which would also flag staged changes and make the
            # hook block every commit touching an inspected root. Staged
            # changes (working tree == index) must pass.
            $script:tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid())
            $script:scriptsDir = Join-Path $script:tempDir 'scripts'
            $script:invocationLog = Join-Path $script:tempDir 'invocation.log'
            New-Item -ItemType Directory -Path $script:scriptsDir -Force | Out-Null

            git -C $script:tempDir init -q
            git -C $script:tempDir config user.email 'test@local'
            git -C $script:tempDir config user.name 'test'
            New-Item -ItemType Directory -Path (Join-Path $script:tempDir 'docs') -Force | Out-Null
            New-Item -ItemType Directory -Path (Join-Path $script:tempDir '.cline\skills') -Force | Out-Null
            New-Item -ItemType Directory -Path (Join-Path $script:tempDir '.github') -Force | Out-Null
            Set-Content -LiteralPath (Join-Path $script:tempDir 'AGENTS.md') -Value 'initial'
            Set-Content -LiteralPath (Join-Path $script:tempDir 'docs\x.md') -Value 'x'
            Set-Content -LiteralPath (Join-Path $script:tempDir '.cline\skills\y.md') -Value 'y'
            Set-Content -LiteralPath (Join-Path $script:tempDir '.github\ci.yml') -Value 'z'
            git -C $script:tempDir add -A
            git -C $script:tempDir commit -q -m 'init'

            # Make a change and STAGE it (working tree == index). No unstaged
            # or untracked changes remain.
            Set-Content -LiteralPath (Join-Path $script:tempDir 'AGENTS.md') -Value 'changed'
            git -C $script:tempDir add AGENTS.md

            # Stub checkers: the guard should pass and the gate should run.
            Write-PreCommitStubChecker -Path (Join-Path $script:scriptsDir 'check-cline-skills.ps1')  -Name 'check-cline-skills.ps1'  -LogPath $script:invocationLog -ExitCode 0
            Write-PreCommitStubChecker -Path (Join-Path $script:scriptsDir 'check-status.ps1')        -Name 'check-status.ps1'        -LogPath $script:invocationLog -ExitCode 0
            Write-PreCommitStubChecker -Path (Join-Path $script:scriptsDir 'check-md-links.ps1')      -Name 'check-md-links.ps1'      -LogPath $script:invocationLog -ExitCode 0

            $script:copyScriptPath = Join-Path $script:scriptsDir 'pre-commit.ps1'
            Copy-Item -LiteralPath $script:scriptPath -Destination $script:copyScriptPath -Force

            $script:stdoutFile = Join-Path $script:tempDir 'stdout.txt'
            $script:stderrFile = Join-Path $script:tempDir 'stderr.txt'
            $script:proc = Start-Process -FilePath pwsh -ArgumentList @('-NoProfile','-File',$script:copyScriptPath) -NoNewWindow -Wait -PassThru -RedirectStandardOutput $script:stdoutFile -RedirectStandardError $script:stderrFile
            $script:exitCode = $script:proc.ExitCode
            $script:output = (Get-Content -LiteralPath $script:stdoutFile -Raw) + (Get-Content -LiteralPath $script:stderrFile -Raw)

            $script:exitCode | Should -Be 0
            $script:output | Should -Match 'pre-commit: quick gates PASS'
            $script:output | Should -Not -Match 'Commit blocked'
        }
    }
}

Describe 'install-pre-commit.ps1' {
    BeforeAll { $script:installScriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'scripts\install-pre-commit.ps1' }

    It 'exists and is readable' {
        (Test-Path -LiteralPath $script:installScriptPath) | Should -BeTrue
    }

    It 'invokes the cross-platform PowerShell 7 (pwsh) for the hook shim' {
        $content = Get-Content -LiteralPath $script:installScriptPath -Raw
        $content | Should -Match 'exec pwsh'
        $content | Should -Match 'core\.hooksPath'
    }

    It 'writes the shim as LF with UTF-8 without BOM' {
        $content = Get-Content -LiteralPath $script:installScriptPath -Raw
        $content | Should -Match 'New-Object System\.Text\.UTF8Encoding'
        $content | Should -Match '\$contentLf'
        $content | Should -Match 'WriteAllText'
    }

    It 'installs the hook shim with POSIX executable bit on non-Windows' {
        # End-to-end verification that the installer produces a hook file with
        # the executable bit set on POSIX hosts. On Windows the executable bit
        # is not required (git for Windows runs hooks via its bundled sh), so
        # we guard the assertion and still verify the hook file is created.
        $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid())
        $scriptsDir = Join-Path $tempDir 'scripts'
        New-Item -ItemType Directory -Path $scriptsDir -Force | Out-Null

        # The installer expects scripts/pre-commit.ps1 to exist relative to
        # the repo root, so plant a minimal stub in the temp repo.
        $stubContent = @'
`$ErrorActionPreference = 'Stop'
Write-Host 'pre-commit stub'
exit 0
'@
        Set-Content -LiteralPath (Join-Path $scriptsDir 'pre-commit.ps1') -Value $stubContent -Encoding utf8NoBOM

        git -C $tempDir init -q
        git -C $tempDir config user.email 'test@local'
        git -C $tempDir config user.name 'test'
        git -C $tempDir commit -q -m 'init' --allow-empty

        # Run the installer against the temp repo. The installer reads its
        # own path to derive $scriptRoot and $repoRoot, so we invoke it from
        # a location where Split-Path -Parent gives us the temp repo root.
        # Simplest approach: copy the installer into the temp repo's scripts/
        # folder and run it from there.
        $installerCopy = Join-Path $scriptsDir 'install-pre-commit.ps1'
        Copy-Item -LiteralPath $script:installScriptPath -Destination $installerCopy -Force

        $stdoutFile = Join-Path $tempDir 'stdout.txt'
        $stderrFile = Join-Path $tempDir 'stderr.txt'
        $proc = Start-Process -FilePath pwsh -ArgumentList @('-NoProfile', '-File', $installerCopy) -NoNewWindow -Wait -PassThru -RedirectStandardOutput $stdoutFile -RedirectStandardError $stderrFile
        $exitCode = $proc.ExitCode
        $output = (Get-Content -LiteralPath $stdoutFile -Raw) + (Get-Content -LiteralPath $stderrFile -Raw)

        $exitCode | Should -Be 0
        $output | Should -Match 'pre-commit hook installed'

        # The installer writes the shim to .githooks/pre-commit relative to
        # the repo root (which is the parent of scripts/).
        $githooksDir = Join-Path $tempDir '.githooks'
        $hookFile = Join-Path $githooksDir 'pre-commit'
        $hookFile | Should -Exist

        if (-not $IsWindows) {
            # On POSIX hosts the installer runs chmod +x; verify the bit is set.
            $attrs = Get-Item -LiteralPath $hookFile
            ($attrs.Attributes -band [System.IO.FileAttributes]::Executable) | Should -Be ([System.IO.FileAttributes]::Executable)
        }

        Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'marks the hook shim executable on POSIX hosts' {
        # Windows hosts do not need the executable bit (git for Windows runs
        # hooks via its bundled sh); POSIX hosts execute the file directly, so
        # the installer must chmod +x it there, guarded by !$IsWindows. The
        # branch cannot run on this Windows test host, hence the source-level
        # assertion on the guard.
        $content = Get-Content -LiteralPath $script:installScriptPath -Raw
        $content | Should -Match 'chmod \+x \$hookFile'
        $content | Should -Match 'IsWindows'
    }
}
