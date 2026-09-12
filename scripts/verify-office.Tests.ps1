#requires -Version 7
<#
.SYNOPSIS
    Pester tests for verify-office.ps1 (L12 watchdog, no --blame-hang-timeout).
#>

Describe 'verify-office.ps1' {
    BeforeAll {
        $script:scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'scripts\verify-office.ps1'
        # Harness shim: a fake `dotnet` that the watchdog is pointed at by
        # absolute path. $env:VERIFY_OFFICE_STUB_BEHAVIOR controls the stub:
        # 'pass' exits 0 at once; 'hang' sleeps past the harness deadline so
        # the watchdog soft-time-out fires.
        $script:stubDir = Join-Path ([System.IO.Path]::GetTempPath()) ('verify-office-stub-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $script:stubDir -Force | Out-Null
        @'
@echo off
if "%VERIFY_OFFICE_STUB_BEHAVIOR%"=="hang" (
  powershell -NoProfile -Command "Start-Sleep -Seconds 10"
)
exit /b 0
'@ | Set-Content -LiteralPath (Join-Path $script:stubDir 'dotnet.cmd') -Encoding utf8NoBOM
@'
@if "%VERIFY_OFFICE_TASKKILL_EXIT%"=="1" (
  exit /b 7
)
exit /b 0
'@ | Set-Content -LiteralPath (Join-Path $script:stubDir 'taskkill.cmd') -Encoding utf8NoBOM

    }

    AfterAll {
        Remove-Item -LiteralPath $script:stubDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'exists and is readable' {
        (Test-Path -LiteralPath $script:scriptPath) | Should -BeTrue
    }

    It 'does not use --blame-hang-timeout (L12)' {
        # The blame-hang collector aborts the testhost ~0.2 s after start
        # on this host; the script must not pass the flag. Strip full-line
        # comments so a comment mentioning the flag does not trip the match.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Not -Match 'blame-hang-timeout'
    }

    It 'timeout path kills the dotnet-test process tree with taskkill /T' {
        # The watchdog must kill the whole dotnet-test tree, not just the
        # launcher. Strip full-line comments so a comment mentioning the
        # command does not trip the match.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match 'taskkill /PID \$testProc\.Id /T /F'
        $codeOnly | Should -Not -Match 'Stop-Process -InputObject \$testProc'
    }

    It 'timeout path sweeps harness-owned Office processes' {
        # On timeout the sweep must run as a safety net for orphaned Office
        # processes that were not in the before-snapshot, passing the owned
        # tree and the fixture's own recorded PIDs.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match 'Remove-HarnessOwnedOfficeProcesses \$officeBeforeSnapshot \$ownedTreePids \$fixtureOwnedPids'
    }

    It 'sweep uses taskkill /T /F by PID, never Stop-Process' {
        # The sweep must kill only processes positively identified as
        # harness-owned, and it must use taskkill /T /F by PID rather than
        # Stop-Process.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match 'taskkill /PID \$processId /T /F 2>\$null'
        $codeOnly | Should -Not -Match 'Stop-Process -InputObject \$processId'
    }

    It 'sweep takes the fixture-owned PIDs as the primary ownership signal' {
        # The fixture-to-shell handoff (owned-PID manifest) is the primary
        # ownership test; parentage is only the fallback.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match 'FixtureOwnedPids'
        $codeOnly | Should -Match 'GANTTCREATOR_OWNED_PIDS_PATH'
        $codeOnly | Should -Match 'owned-office-pids.json'
        $codeOnly | Should -Match 'recorded by OfficeFixture'
        # The old "not present before test run = kill" inference must be gone.
        $codeOnly | Should -Not -Match 'not present before test run'
    }

    It 'sweep refuses to kill when neither the fixture PID nor the parent is known' {
        # The safe direction: an Office process that is neither recorded by
        # the fixture nor a child of the owned tree is skipped, never killed.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match 'Skipping Office process PID'
    }

    It 'the owned-PID manifest path is derived from the evidence directory, never committed' {
        # The manifest must live under the ignored scripts/_artifacts/ tree
        # and be deleted before the test run starts.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match 'Join-Path \$evidence ''owned-office-pids.json'''
        $codeOnly | Should -Match 'Remove-Item -LiteralPath \$ownedPidsPath'
    }


    It 'captures an Office process snapshot before starting dotnet test' {
        # The sweep must distinguish harness-owned Office processes from
        # user-owned ones by comparing against a before-snapshot.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match "Get-Process -Name 'EXCEL','POWERPNT' -ErrorAction SilentlyContinue"
        $codeOnly | Should -Match '\$officeBeforeSnapshot'
    }

    It 'positive control: the taskkill assertion fires on a flagged stub' {
        # Ensure the negative assertion above can detect a regression where
        # taskkill /T is removed from the timeout path.
        $flagged = "if (-not `$testProc.HasExited) { Stop-Process -InputObject `$testProc -Force }`n"
        $codeOnly = $flagged -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Not -Match 'taskkill /PID \$testProc\.Id /T /F'
    }


    It 'positive control: the --blame-hang-timeout assertion fires on a flagged stub' {
        $flagged = "dotnet test --blame-hang-timeout 600`n"
        $codeOnly = $flagged -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match 'blame-hang-timeout' -Because 'the negative assertion above must be able to detect the flag'
    }

    Context 'watchdog deadline (isolated harness)' {
        BeforeEach {
            # Registry/build steps would run for real, so strip them and point
            # the watchdog at the stub. Literal String.Replace is used (not
            # regex) so the patches are deterministic and cannot silently
            # no-op if the source shifts.
            $script:harnessRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('verify-office-' + [guid]::NewGuid().ToString('N'))
            $script:harnessScripts = Join-Path $script:harnessRoot 'scripts'
            New-Item -ItemType Directory -Path $script:harnessScripts -Force | Out-Null
            $body = Get-Content -LiteralPath $script:scriptPath -Raw

            # Resolve the stub path once, before any replacement uses it.
            # The stub is a .cmd file executed via cmd /c, since Start-Process
            # cannot reliably run .cmd files directly as -FilePath.
            $stubCmd = Join-Path $script:stubDir 'dotnet.cmd'

            $body = $body.Replace(
                '$cfg = Get-ItemProperty ''HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'' -ErrorAction SilentlyContinue',
                '$cfg = @{ ProductReleaseIds = ''harness''; Platform = ''x64''; VersionToReport = ''16.0''; ClientCulture = ''en-US'' }'
            )
            $body = $body.Replace(
                'dotnet build $Solution -c $Configuration --no-restore -warnaserror',
                'Write-Host ''harness: skip build''; $LASTEXITCODE = 0'
            )
            # Use cmd /c to execute the .cmd stub, since Start-Process
            # cannot reliably run .cmd files directly as -FilePath.
            $body = $body.Replace(
                '$dotnetExe = ''dotnet''',
                '$dotnetExe = ''cmd'''
            )
            $body = $body.Replace('$watchdogIntervalSeconds = 5', '$watchdogIntervalSeconds = 1')
            $body = $body.Replace('$cleanupDeadlineSeconds = 30', '$cleanupDeadlineSeconds = 1')
            # Replace the entire $testArgs multi-line block with cmd /c <stub>.
            # Use a regex for the replacement since we need to match a
            # multi-line block. The pattern matches from $testArgs = @(
            # through the closing ) on its own line.
            $body = $body -replace '(?ms)\$testArgs = @\([^)]*\)\s*\n', "`$testArgs = @('/c', '$stubCmd')`n"
            $script:harnessScript = Join-Path $script:harnessScripts 'verify-office.ps1'
            $body | Set-Content -LiteralPath $script:harnessScript -Encoding utf8NoBOM
        }

        AfterEach {
            Remove-Item -LiteralPath $script:harnessRoot -Recurse -Force -ErrorAction SilentlyContinue
        }

        It 'soft-times-out with exit 124 and a TIMEOUT diagnostic when dotnet hangs' {
            $outFile = Join-Path $script:harnessRoot 'out.txt'
            $errFile = Join-Path $script:harnessRoot 'err.txt'
            $sep = [System.IO.Path]::PathSeparator
            $envArgs = @(
                '-NoProfile', '-File', $script:harnessScript,
                '-DeadlineSeconds', '2'
            )
            $proc = Start-Process -FilePath pwsh -ArgumentList $envArgs -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile `
                -Environment @{ VERIFY_OFFICE_STUB_BEHAVIOR = 'hang'; PATH = ($script:stubDir + $sep + $env:PATH) }
            $output = (Get-Content -LiteralPath $outFile -Raw) + (Get-Content -LiteralPath $errFile -Raw)
            $proc.ExitCode | Should -Be 124
            $output | Should -Match 'TIMEOUT'
        }

        It 'reports cleanup failure when taskkill fails and the process tree remains active' {
            $outFile = Join-Path $script:harnessRoot 'out.txt'
            $errFile = Join-Path $script:harnessRoot 'err.txt'
            $sep = [System.IO.Path]::PathSeparator
            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile', '-File', $script:harnessScript,
                '-DeadlineSeconds', '2'
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile `
                -Environment @{ VERIFY_OFFICE_STUB_BEHAVIOR = 'hang'; VERIFY_OFFICE_TASKKILL_EXIT = '1'; PATH = ($script:stubDir + $sep + $env:PATH) }
            $output = (Get-Content -LiteralPath $outFile -Raw) + (Get-Content -LiteralPath $errFile -Raw)
            $proc.ExitCode | Should -Be 124
            $output | Should -Match 'taskkill failed to stop the dotnet-test process tree'
            $output | Should -Match 'Cleanup failed: dotnet-test process tree.*still active'
            $output | Should -Match 'TIMEOUT'
        }

        It 'passes through exit 0 when dotnet exits promptly (positive control)' {
            $outFile = Join-Path $script:harnessRoot 'out.txt'
            $errFile = Join-Path $script:harnessRoot 'err.txt'
            $envArgs = @(
                '-NoProfile', '-File', $script:harnessScript,
                '-DeadlineSeconds', '60'
            )
            $sep = [System.IO.Path]::PathSeparator
            $proc = Start-Process -FilePath pwsh -ArgumentList $envArgs -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile `
                -Environment @{ VERIFY_OFFICE_STUB_BEHAVIOR = 'pass'; PATH = ($script:stubDir + $sep + $env:PATH) }
            $output = (Get-Content -LiteralPath $outFile -Raw) + (Get-Content -LiteralPath $errFile -Raw)
            $proc.ExitCode | Should -Be 0
            $output | Should -Match 'verify-office: PASS'
        }
    }
}
