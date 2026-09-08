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
        $codeOnly | Should -Match 'check-skill-summary\.ps1'
        $codeOnly | Should -Match 'check-status\.ps1'
        $codeOnly | Should -Match 'check-md-links\.ps1'
        # Must not INVOKE verify-quick.ps1 in the executable portion.
        $codeOnly | Should -Not -Match '(pwsh|\$\(|&)\s.*verify-quick\.ps1'
    }

    It 'fails fast with a non-zero exit and the blocking message' {
        # The script must surface a 'Commit blocked' line and propagate
        # a non-zero exit code. The exact token used to do that is an
        # implementation detail (previously xit $LASTEXITCODE; now
        # xit $checkerExit after the $LASTEXITCODE-reset fix).
        $content = Get-Content -LiteralPath $script:scriptPath -Raw
        $content | Should -Match 'Commit blocked'
        $content | Should -Match 'exit \$checkerExit'
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
            $script:scriptsDir = Join-Path $script:tempDir 'scripts'
            $script:invocationLog = Join-Path $script:tempDir 'invocation.log'
            New-Item -ItemType Directory -Path $script:scriptsDir -Force | Out-Null

            # Copy the real pre-commit.ps1 verbatim. We do NOT dot-source it
            # or override its private $checks variable; the test exercises
            # the same code that runs in production, with only the checker
            # scripts replaced by stubs.
            $script:copyScriptPath = Join-Path $script:scriptsDir 'pre-commit.ps1'
            Copy-Item -LiteralPath $script:scriptPath -Destination $script:copyScriptPath -Force
        }

        AfterEach {
            if ($script:tempDir -and (Test-Path -LiteralPath $script:tempDir)) {
                Remove-Item -LiteralPath $script:tempDir -Recurse -Force
            }
        }

        It 'invokes all four checkers in declared order and exits 0 when all pass' {
            Write-PreCommitStubChecker -Path (Join-Path $script:scriptsDir 'check-cline-skills.ps1')  -Name 'check-cline-skills.ps1'   -LogPath $script:invocationLog -ExitCode 0
            Write-PreCommitStubChecker -Path (Join-Path $script:scriptsDir 'check-skill-summary.ps1') -Name 'check-skill-summary.ps1'  -LogPath $script:invocationLog -ExitCode 0
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
                'check-skill-summary.ps1',
                'check-status.ps1',
                'check-md-links.ps1'
            )
        }

        It 'exits non-zero with the blocking message and short-circuits on first failure' {
            # First two pass; third fails; fourth is staged but must never run.
            Write-PreCommitStubChecker -Path (Join-Path $script:scriptsDir 'check-cline-skills.ps1')  -Name 'check-cline-skills.ps1'   -LogPath $script:invocationLog -ExitCode 0
            Write-PreCommitStubChecker -Path (Join-Path $script:scriptsDir 'check-skill-summary.ps1') -Name 'check-skill-summary.ps1'  -LogPath $script:invocationLog -ExitCode 0
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

            # Fail-fast proof: the 3rd checker ran, the 4th did not.
            Get-Content -LiteralPath $script:invocationLog | Should -Be @(
                'check-cline-skills.ps1',
                'check-skill-summary.ps1',
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
            Write-PreCommitTerminatingErrorStubChecker -Path (Join-Path $script:scriptsDir 'check-skill-summary.ps1') -Name 'check-skill-summary.ps1' -LogPath $script:invocationLog -ExitCode 1
            Write-PreCommitStubChecker               -Path (Join-Path $script:scriptsDir 'check-status.ps1')        -Name 'check-status.ps1'        -LogPath $script:invocationLog -ExitCode 0

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
                'check-skill-summary.ps1'
            )
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
}
