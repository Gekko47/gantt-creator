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
rem A real run WRITES its trx. VERIFY_OFFICE_STUB_TRX names where, so a test can model
rem "this run produced results" instead of pre-creating a stale file that the harness
rem would then archive as if it came from this run.
if defined VERIFY_OFFICE_STUB_TRX (
  for %%I in ("%VERIFY_OFFICE_STUB_TRX%") do if not exist "%%~dpI" mkdir "%%~dpI"
  echo ^<TestRun^>harness-fake^</TestRun^> > "%VERIFY_OFFICE_STUB_TRX%"
)
rem A real run also writes ONE forced-kill record per fixture teardown, and the
rem gate now fails closed when none was written. VERIFY_OFFICE_STUB_KILLS models
rem the count: unset or 0 is a clean teardown, 1 is the normal one-kill teardown,
rem 2 is the positive case for the per-teardown ceiling, and 'none' models the
rem handoff breaking so no record exists at all.
rem
rem VERIFY_OFFICE_STUB_RCW models the application reference count the teardown
rem recorded: unset or 0 is the last-managed-holder case, 1 is a leaked reference,
rem and 'absent' models a fixture that never wrote the field, which the gate must
rem read as unmeasured rather than clean.
if not defined VERIFY_OFFICE_STUB_KILLS set VERIFY_OFFICE_STUB_KILLS=0
if not defined VERIFY_OFFICE_STUB_RCW set VERIFY_OFFICE_STUB_RCW=0
set "STUB_RCW_FRAGMENT=,"Rcw":%VERIFY_OFFICE_STUB_RCW%"
if /i "%VERIFY_OFFICE_STUB_RCW%"=="absent" set "STUB_RCW_FRAGMENT="
if /i not "%VERIFY_OFFICE_STUB_KILLS%"=="none" (
  if defined GANTTCREATOR_FORCED_KILLS_PATH (
    for %%I in ("%GANTTCREATOR_FORCED_KILLS_PATH%") do if not exist "%%~dpI" mkdir "%%~dpI"
    echo {"Test":"stub-teardown","Count":%VERIFY_OFFICE_STUB_KILLS%,"GracefulExitMs":0,"Quit":"returned","BodyMs":1234%STUB_RCW_FRAGMENT%}>>"%GANTTCREATOR_FORCED_KILLS_PATH%"
  )
)
rem VERIFY_OFFICE_STUB_CLEAN=1 adds a second record whose teardown exited
rem gracefully, so a test can prove the clean list is reported too.
if "%VERIFY_OFFICE_STUB_CLEAN%"=="1" (
  echo {"Test":"stub-clean","Count":0,"GracefulExitMs":0,"Quit":"returned","BodyMs":900%STUB_RCW_FRAGMENT%}>>"%GANTTCREATOR_FORCED_KILLS_PATH%"
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
        # tree and the fixture's own recorded PIDs. Named parameters: the
        # fixture signal is a PID -> start-time map, not a bare int array.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match 'Remove-HarnessOwnedOfficeProcesses -BeforeSnapshot \$officeBeforeSnapshot -OwnedTreePids \$ownedTreePids -FixtureManifest \$fixtureManifest'
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
        $codeOnly | Should -Match 'FixtureManifest'
        $codeOnly | Should -Match 'Test-ManifestProcessIdentity'
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

    It 'reads the owned-PID manifest line by line, not as one JSON document' {
        # The fixture appends one JSON object per line. Converting-JSON on the
        # whole file fails on concatenated objects ("Additional text
        # encountered after finished reading JSON content"), which is what
        # silently emptied the sweep's primary ownership signal.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match 'foreach \(\$line in @\(Get-Content -LiteralPath \$Path'
        # Single-quoted so `$ownedPidsPath` stays literal: in a double-quoted
        # PowerShell string the backslash is not an escape character, so the
        # variable would expand to empty and the pattern could never match the
        # forbidden command, making the negative assertion vacuous.
        $codeOnly | Should -Not -Match 'Get-Content -LiteralPath \$ownedPidsPath -Raw \| ConvertFrom-Json'
    }

    It 'positive control: the whole-file manifest assertion fires on a flagged stub' {
        # Proves the preceding negative assertion can actually detect the
        # regression it guards, rather than passing because its pattern is dead.
        $flagged = "Get-Content -LiteralPath `$ownedPidsPath -Raw | ConvertFrom-Json`n"
        $codeOnly = $flagged -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match 'Get-Content -LiteralPath \$ownedPidsPath -Raw \| ConvertFrom-Json'
    }

    It 'requires the recorded start time to match before killing a manifest PID' {
        # A PID alone is not identity: Windows recycles process IDs, so a stale
        # manifest can name an unrelated process. The start time is the second
        # half of the proof, and a mismatch must skip rather than kill.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match 'Test-ManifestProcessIdentity'
        $codeOnly | Should -Match 'StartTimeUtcTicks'
        $codeOnly | Should -Match '\$liveTicks -ne \$recordedTicks'
        $codeOnly | Should -Match 'PID was recycled'
    }

    It 'refuses to treat a manifest record with no start time as owned' {
        # Fail-safe direction: an unverifiable record is skipped, never killed.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match 'if \(\$recordedTicks -le 0\)'
        $codeOnly | Should -Match 'carries no start time'
    }

    It 'sweeps harness-owned Office processes before the build, not only on timeout' {
        # A stray from a previous run holds the packed XLL, and the build's
        # ExcelDnaPack step fails first -- before any post-test sweep could run.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $preflightIndex = $codeOnly.IndexOf('preflight: sweeping')
        $buildIndex = $codeOnly.IndexOf('dotnet build $Solution')
        $preflightIndex | Should -BeGreaterThan 0
        $buildIndex | Should -BeGreaterThan 0
        $preflightIndex | Should -BeLessThan $buildIndex
    }

    It 'sweeps on the non-timeout exit paths too' {
        # The sweep used to run only inside the timeout branch, so a green run
        # that leaked an Excel left it holding the XLL for the next run.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $saveTrx = $codeOnly.LastIndexOf('Save-Trx')
        $postRunSweep = $codeOnly.IndexOf('$finalManifest = Read-OwnedPidManifest')
        $postRunSweep | Should -BeGreaterThan $saveTrx
    }

    It 'fails the gate when a stray Office process had to be killed' {
        # A killed stray is a real leak, not a warning: it means a test left an
        # Excel holding the packed XLL. It must not pass silently.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match '\$straysKilled -gt 0'
        $codeOnly | Should -Match 'survived the test run and had to be killed'
        $codeOnly | Should -Match 'exit 3'
    }

    It 'gates the forced-exit signal per teardown, not by an absolute run total' {
        # One fixture owns one Excel process, so one teardown contributes at most
        # one forced exit. The run total therefore counts TEST BODIES that hit the
        # shutdown race and tracks suite size -- 38 at 63 bodies, 40 at 65 and 67,
        # 41 on the R5.22 tree -- which is why the retired absolute ceiling had to
        # be raised six times and always to the highest observation, leaving it no
        # margin.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match '\$MaxForcedExitsPerTeardown = \d+'
        $codeOnly | Should -Match '\$count -gt \$MaxForcedExitsPerTeardown'
        $codeOnly | Should -Match 'forced more than \$MaxForcedExitsPerTeardown exit\(s\)'
        $codeOnly | Should -Match 'exit 4'
        # The absolute total must not gate: it can only fail on suite size.
        $codeOnly | Should -Not -Match '\$forcedKills -gt \$MaxForcedKills\b'
        $codeOnly | Should -Not -Match 'exceeds the ratchet ceiling'
    }

    It 'pins the per-teardown ceiling at 1' {
        # The invariant, not a baseline: a fixture owns one process, so a teardown
        # that needed two kills owned two processes or escalated twice.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match '\$MaxForcedExitsPerTeardown = 1'
    }

    It 'gates the application reference count at zero and fails closed when it was not measured' {
        # Measured 2026-11-25 over 68 live teardowns: every record released the
        # Application proxy to 0 references (22 clean, 46 with one forced kill),
        # so the kill is the host's residency and the count is the only leak this
        # gate can observe. A build that stopped writing the field would otherwise
        # read as "no leaks", which is why the unmeasured case fails closed.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match "Properties\['Rcw'\]"
        $codeOnly | Should -Match '\$rcwMeasured -eq 0'
        $codeOnly | Should -Match 'leak signal was NOT measured'
        $codeOnly | Should -Match '\$rcwLeaks\.Count -gt 0'
        $codeOnly | Should -Match 'left an application reference behind'
        $codeOnly | Should -Match 'exit 7'
    }

    It 'builds each shard term as a self-contained category conjunction' {
        # A bare `FullyQualifiedName~Class` OR list matches by substring and ignores the
        # category: a rehearsal split on 2026-11-25 ran 143 tests against the gate's 103 and
        # pulled in the excluded TransferSoak body. Each term now carries its own predicate,
        # so the expression needs no grouping (`&` binds tighter than `|`).
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match 'Category=OfficeIntegration&FullyQualifiedName~\$_'
        # The bare form must not survive beside it.
        $codeOnly | Should -Not -Match 'ForEach-Object \{ "FullyQualifiedName~\$_" \}'
    }

    It 'fails the gate when a shard ran a test that discovery did not list' {
        # The other direction from the missing-test check: a filter wider than the gate would
        # otherwise be judged on a different suite than it reports.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match '\$unexpectedTests\.Count -gt 0'
        $codeOnly | Should -Match 'discovery did not list'
        $codeOnly | Should -Match 'exit 6'
    }

    It 'keeps a Theory in the discovered set even though the host lists it with arguments' {
        # A [Theory] is listed ONLY as one row per data set, e.g.
        # `..._state(artefactName: "Shape")`. A discovery pattern that required the line to end
        # after the method silently dropped every Theory, and the live two-shard run of
        # 2026-11-25 then reported the method as "ran but was not listed".
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match ([regex]::Escape('(?:\(.*\))?'))
        $codeOnly | Should -Match '\$listedTestNames'
        $codeOnly | Should -Match 'gate discovered \$\(\$listedTestNames\.Count\)'
    }

    It 'logs each shard''s own duration and counts distinct methods in the coverage line' {
        # A one-shard run and a two-shard run are only comparable if each shard's wall time is
        # in the report, and the coverage line must not count a [Theory]'s rows as tests: the
        # 2026-11-25 two-shard run reported "discovered 104" beside "all 106 discovered".
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match 'shard \{0\}: exit \{1\} after'
        $codeOnly | Should -Match 'coverage: all \$\(\$listedTestNames\.Count\)'
        $codeOnly | Should -Not -Match 'coverage: all \$\(\$listedTests\.Count\)'
    }

    It 'canonicalises both name sets before comparing them' {
        # A [Theory] is spelled differently by the two sources: --list-tests lists each data
        # row as `Method(arg: value)` while the trx also carries the bare method name. The
        # first version stripped only one side and reported a Theory as "ran but was not
        # listed" on the live two-shard run of 2026-11-25.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match '\$executedNames = @\(\$executedTests'
        $codeOnly | Should -Match '\$listedNames = @\(\$listedTests'
        $codeOnly | Should -Match '\$listedNames \| Where-Object \{ \$executedNames -notcontains \$_ \}'
        $codeOnly | Should -Match '\$executedNames \| Where-Object \{ \$listedNames -notcontains \$_ \}'
    }

    It 'reports the per-body host launch cost beside the reference count' {
        # Recorded since 2026-11-25 but read by nothing, so the P12 decision had to open the raw
        # JSONL. Reported, never gated: it is context for the next host question, not a threshold.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match "Properties\['HostMs'\]"
        $codeOnly | Should -Match 'Host-launch signal:'
        $codeOnly | Should -Match '\$hostSumMs'
        $codeOnly | Should -Match '\$hostMaxMs'
    }

    It 'fails closed when the test host lists no tests, unless the harness opts out' {
        # The coverage proof depends on the discovered list, so an empty list means the gate cannot
        # prove which tests it ran. It used to degrade to one unsharded pass with a warning, which
        # is exactly the false green the proof exists to stop.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match '\[switch\]\$AllowUnprovenDiscovery'
        $codeOnly | Should -Match 'if \(-not \$AllowUnprovenDiscovery\)'
        $codeOnly | Should -Match 'An unproven run is not a pass'
        $codeOnly | Should -Match 'NO coverage proof'
        $codeOnly | Should -Match 'exit 6'
    }

    It 'positive control: the over-coverage verdict is detected when it is absent' {
        # Proves the two assertions above can fail: a source carrying only the missing-test
        # check must pass neither of them.
        $flagged = "if (`$missingTests.Count -gt 0) { exit 6 }`n"
        $codeOnly = $flagged -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Not -Match '\$unexpectedTests\.Count -gt 0'
        $codeOnly | Should -Not -Match 'discovery did not list'
    }

    It 'positive control: the reference-count verdict is detected when it is absent' {
        # Proves the negative assertions above can detect a regression: a source
        # with no reference-count idiom must fail every one of them.
        $flagged = "if (`$teardowns -eq 0) { exit 5 }`n"
        $codeOnly = $flagged -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Not -Match "Properties\['Rcw'\]"
        $codeOnly | Should -Not -Match '\$rcwLeaks\.Count -gt 0'
        $codeOnly | Should -Not -Match 'left an application reference behind'
    }

    It 'fails closed when no teardown record was written' {
        # An unmeasured signal is not a zero signal: the retired total read a
        # missing log as 0 kills and passed the gate.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match '\$teardowns -eq 0'
        $codeOnly | Should -Match 'was NOT measured'
        $codeOnly | Should -Match 'exit 5'
    }

    It 'excludes a deliberate escalation from the forced-exit signal' {
        # A test that forces a kill to prove the escalation path works must not
        # count against the ratchet, or the ceiling can never be reached.
        $fixture = Get-Content -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'tests\GanttCreator.Office.IntegrationTests\OfficeFixture.cs') -Raw
        $fixture | Should -Match 'SuppressLeakSignal'
        $fixture | Should -Match 'if \(SuppressLeakSignal\)'
    }

    It 'positive control: the retired total ceiling is detected if it comes back' {
        # Proves the negative assertion above is not dead: the forbidden
        # comparison is recognised when it is present.
        $flagged = "if (`$forcedKills -gt `$MaxForcedKills) { exit 4 }`n"
        $codeOnly = $flagged -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match '\$forcedKills -gt \$MaxForcedKills\b'
    }

    It 'positive control: the fail-closed assertion detects a written message' {
        $flagged = "if (`$teardowns -eq 0) { Write-Host 'the signal was NOT measured'; exit 5 }`n"
        $codeOnly = $flagged -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match 'was NOT measured'
    }

    It 'names the clean teardowns beside the forced ones' {
        # The clean list is the control side of the measurement, so it is part of
        # the report: R5.18n D4 is closed on the finding that the two sides differ
        # by an explicit release, not by an unreleased proxy chain.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match 'clean teardowns'
        $codeOnly | Should -Match '\$cleanTeardowns\.Add'
        $codeOnly | Should -Match 'StringComparer\]::Ordinal'
    }

    It 'reports what Quit() did in each recorded teardown' {
        # A Quit that returns while the process survives and a Quit the host
        # refuses are different defects, and the count alone cannot tell them
        # apart. R5.18n D4 needs the outcome in the run log.
        $raw = Get-Content -LiteralPath $script:scriptPath -Raw
        $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Match '\$quitOutcomes'
        $codeOnly | Should -Match "Properties\['Quit'\]"
        $codeOnly | Should -Match 'by Quit\(\) outcome'
    }

    It 'positive control: the fail-on-stray assertion fires on a flagged stub' {
        # Ensures the negative assertion above can detect a regression where the
        # stray verdict is removed.
        $flagged = "if (`$straysKilled -gt 0) { }`n"
        $codeOnly = $flagged -replace '(?m)^\s*#.*$', ''
        $codeOnly | Should -Not -Match 'survived the test run and had to be killed'
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
            # The stub `dotnet` prints no test names, so the real fail-closed guard would end every
            # case below with exit 6. The harness therefore opts out explicitly: this script keeps
            # the guard but also accepts the documented switch, and the guard itself is driven
            # end-to-end by the case that runs $script:guardScript (the same body, unpatched).
            $script:guardScript = Join-Path $script:harnessScripts 'verify-office-guard.ps1'
            $body | Set-Content -LiteralPath $script:guardScript -Encoding utf8NoBOM
            # Disable exactly the fail-closed branch, leaving the WARN path and every verdict after
            # it intact. Replacing the outer condition instead would have re-enabled the guard.
            $body = $body.Replace(
                'if (-not $AllowUnprovenDiscovery) {',
                'if ($false) {'
            )
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

        It 'fails closed with exit 6 when the host lists no tests and no opt-out was passed' {
            # Drives the guard the harness otherwise patches out: the same generated body, unpatched.
            # A stub `dotnet` lists no tests, which is the exact condition the guard refuses.
            $outFile = Join-Path $script:harnessRoot 'out.txt'
            $errFile = Join-Path $script:harnessRoot 'err.txt'
            $sep = [System.IO.Path]::PathSeparator
            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile', '-File', $script:guardScript,
                '-DeadlineSeconds', '60'
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile `
                -Environment @{ VERIFY_OFFICE_STUB_BEHAVIOR = 'pass'; PATH = ($script:stubDir + $sep + $env:PATH) }
            $output = (Get-Content -LiteralPath $outFile -Raw) + (Get-Content -LiteralPath $errFile -Raw)
            $proc.ExitCode | Should -Be 6
            $output | Should -Match 'An unproven run is not a pass'
            $output | Should -Not -Match 'verify-office: PASS'
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

        It 'archives the per-run trx before the verdict (positive control)' {
            # dotnet test overwrites office.trx every run; without the archive
            # a failing run's results are destroyed by the next invocation.
            # The stub WRITES the trx, because that is what a real run does - a
            # pre-created file would be indistinguishable from a stale one.
            $trxSrc = Join-Path $script:harnessRoot 'tests\GanttCreator.Office.IntegrationTests\TestResults\office.trx'
            $outFile = Join-Path $script:harnessRoot 'out.txt'
            $errFile = Join-Path $script:harnessRoot 'err.txt'
            $sep = [System.IO.Path]::PathSeparator
            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile', '-File', $script:harnessScript,
                '-DeadlineSeconds', '60'
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile `
                -Environment @{
                    VERIFY_OFFICE_STUB_BEHAVIOR = 'pass'
                    VERIFY_OFFICE_STUB_TRX = $trxSrc
                    PATH = ($script:stubDir + $sep + $env:PATH)
                }
            $output = (Get-Content -LiteralPath $outFile -Raw) + (Get-Content -LiteralPath $errFile -Raw)
            $proc.ExitCode | Should -Be 0
            $output | Should -Match 'Archived trx'
            $archives = @(Get-ChildItem -LiteralPath (Join-Path $script:harnessScripts '_artifacts\office-evidence') -Filter 'office-*.trx' -ErrorAction SilentlyContinue)
            $archives.Count | Should -Be 1
            Get-Content -LiteralPath $archives[0].FullName -Raw | Should -BeLike '*harness-fake*'
        }

        It 'refuses to archive a trx this run did not produce' {
            # Measured 2026-11-09: a timed-out gate archived a trx dated four days
            # earlier, and reading it produced a false diagnosis. A file left by a
            # previous run must be named as absent, not presented as this run's.
            $trxSrc = Join-Path $script:harnessRoot 'tests\GanttCreator.Office.IntegrationTests\TestResults\office.trx'
            New-Item -ItemType Directory -Path (Split-Path -Parent $trxSrc) -Force | Out-Null
            Set-Content -LiteralPath $trxSrc -Value '<TestRun>stale-from-a-previous-run</TestRun>'
            $outFile = Join-Path $script:harnessRoot 'out.txt'
            $errFile = Join-Path $script:harnessRoot 'err.txt'
            $sep = [System.IO.Path]::PathSeparator
            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile', '-File', $script:harnessScript,
                '-DeadlineSeconds', '60'
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile `
                -Environment @{
                    VERIFY_OFFICE_STUB_BEHAVIOR = 'pass'
                    PATH = ($script:stubDir + $sep + $env:PATH)
                }
            $output = (Get-Content -LiteralPath $outFile -Raw) + (Get-Content -LiteralPath $errFile -Raw)
            $proc.ExitCode | Should -Be 0
            $output | Should -Match 'no trx was produced by this run'
            $output | Should -Not -Match 'Archived trx'
            $archives = @(Get-ChildItem -LiteralPath (Join-Path $script:harnessScripts '_artifacts\office-evidence') -Filter 'office-*.trx' -ErrorAction SilentlyContinue)
            $archives.Count | Should -Be 0
        }

        It 'invokes the trx archive on both verdict paths (timeout and normal)' {
            # One function definition + one call on the timeout path (before
            # exit 124) + one call covering the fail/pass verdict. The count
            # fires if a call site is deleted.
            $raw = Get-Content -LiteralPath $script:scriptPath -Raw
            $codeOnly = $raw -replace '(?m)^\s*#.*$', ''
            ([regex]::Matches($codeOnly, 'Save-Trx')).Count | Should -Be 3
        }

        It 'reports the forced-exit signal and passes when every teardown forced one exit' {
            # The measured shape of an accepted run: one record, one forced exit.
            # This is the positive control for the per-teardown ceiling, which is
            # the case the retired absolute total failed on. It is ALSO the
            # positive control for the leak gate's corrected meaning: a teardown
            # that needed a kill but released the proxy to 0 references passes.
            $outFile = Join-Path $script:harnessRoot 'out.txt'
            $errFile = Join-Path $script:harnessRoot 'err.txt'
            $sep = [System.IO.Path]::PathSeparator
            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile', '-File', $script:harnessScript,
                '-DeadlineSeconds', '60'
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile `
                -Environment @{ VERIFY_OFFICE_STUB_BEHAVIOR = 'pass'; VERIFY_OFFICE_STUB_KILLS = '1'; PATH = ($script:stubDir + $sep + $env:PATH) }
            $output = (Get-Content -LiteralPath $outFile -Raw) + (Get-Content -LiteralPath $errFile -Raw)
            $proc.ExitCode | Should -Be 0
            $output | Should -Match '1 forced exit\(s\) across 1 recorded teardown\(s\)'
            $output | Should -Match 'Reference-count signal: 1 of 1 teardown\(s\) reported an application reference count, and 0 left one behind'
            $output | Should -Match 'verify-office: PASS'
        }

        It 'fails with exit 7 when a teardown left an application reference behind' {
            # The leak gate itself. A non-zero count means the fixture was not the
            # last managed holder, which no amount of host behaviour explains.
            $outFile = Join-Path $script:harnessRoot 'out.txt'
            $errFile = Join-Path $script:harnessRoot 'err.txt'
            $sep = [System.IO.Path]::PathSeparator
            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile', '-File', $script:harnessScript,
                '-DeadlineSeconds', '60'
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile `
                -Environment @{ VERIFY_OFFICE_STUB_BEHAVIOR = 'pass'; VERIFY_OFFICE_STUB_KILLS = '1'; VERIFY_OFFICE_STUB_RCW = '1'; PATH = ($script:stubDir + $sep + $env:PATH) }
            $output = (Get-Content -LiteralPath $outFile -Raw) + (Get-Content -LiteralPath $errFile -Raw)
            $proc.ExitCode | Should -Be 7
            $output | Should -Match 'stub-teardown left 1 application reference\(s\)'
            $output | Should -Match 'left an application reference behind'
        }

        It 'fails closed when no teardown reported a reference count' {
            # Same rule as the missing-log case: an unmeasured leak signal must not
            # read as a clean one.
            $outFile = Join-Path $script:harnessRoot 'out.txt'
            $errFile = Join-Path $script:harnessRoot 'err.txt'
            $sep = [System.IO.Path]::PathSeparator
            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile', '-File', $script:harnessScript,
                '-DeadlineSeconds', '60'
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile `
                -Environment @{ VERIFY_OFFICE_STUB_BEHAVIOR = 'pass'; VERIFY_OFFICE_STUB_KILLS = '1'; VERIFY_OFFICE_STUB_RCW = 'absent'; PATH = ($script:stubDir + $sep + $env:PATH) }
            $output = (Get-Content -LiteralPath $outFile -Raw) + (Get-Content -LiteralPath $errFile -Raw)
            $proc.ExitCode | Should -Be 5
            $output | Should -Match 'leak signal was NOT measured'
        }

        It 'fails when one teardown forced more than one exit' {
            # One fixture owns one Excel process, so two kills from one teardown
            # is the anomaly the per-teardown ceiling exists to catch.
            $outFile = Join-Path $script:harnessRoot 'out.txt'
            $errFile = Join-Path $script:harnessRoot 'err.txt'
            $sep = [System.IO.Path]::PathSeparator
            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile', '-File', $script:harnessScript,
                '-DeadlineSeconds', '60'
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile `
                -Environment @{ VERIFY_OFFICE_STUB_BEHAVIOR = 'pass'; VERIFY_OFFICE_STUB_KILLS = '2'; PATH = ($script:stubDir + $sep + $env:PATH) }
            $output = (Get-Content -LiteralPath $outFile -Raw) + (Get-Content -LiteralPath $errFile -Raw)
            $proc.ExitCode | Should -Be 4
            $output | Should -Match 'forced more than 1 exit\(s\)'
        }

        It 'fails closed when no teardown record was written' {
            # The retired total read a missing log as zero kills and passed. An
            # unmeasured signal must not be a zero signal.
            $outFile = Join-Path $script:harnessRoot 'out.txt'
            $errFile = Join-Path $script:harnessRoot 'err.txt'
            $sep = [System.IO.Path]::PathSeparator
            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile', '-File', $script:harnessScript,
                '-DeadlineSeconds', '60'
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile `
                -Environment @{ VERIFY_OFFICE_STUB_BEHAVIOR = 'pass'; VERIFY_OFFICE_STUB_KILLS = 'none'; PATH = ($script:stubDir + $sep + $env:PATH) }
            $output = (Get-Content -LiteralPath $outFile -Raw) + (Get-Content -LiteralPath $errFile -Raw)
            $proc.ExitCode | Should -Be 5
            $output | Should -Match 'was NOT measured'
        }

        It 'reports the clean teardowns in the same run as the forced ones' {
            $outFile = Join-Path $script:harnessRoot 'out.txt'
            $errFile = Join-Path $script:harnessRoot 'err.txt'
            $sep = [System.IO.Path]::PathSeparator
            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile', '-File', $script:harnessScript,
                '-DeadlineSeconds', '60'
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile `
                -Environment @{ VERIFY_OFFICE_STUB_BEHAVIOR = 'pass'; VERIFY_OFFICE_STUB_KILLS = '1'; VERIFY_OFFICE_STUB_CLEAN = '1'; PATH = ($script:stubDir + $sep + $env:PATH) }
            $output = (Get-Content -LiteralPath $outFile -Raw) + (Get-Content -LiteralPath $errFile -Raw)
            $proc.ExitCode | Should -Be 0
            $output | Should -Match '1 forced exit\(s\) across 2 recorded teardown\(s\)'
            $output | Should -Match 'clean teardowns'
            $output | Should -Match 'stub-clean'
        }

        It 'reports the Quit() outcome the records carry' {
            $outFile = Join-Path $script:harnessRoot 'out.txt'
            $errFile = Join-Path $script:harnessRoot 'err.txt'
            $sep = [System.IO.Path]::PathSeparator
            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile', '-File', $script:harnessScript,
                '-DeadlineSeconds', '60'
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile `
                -Environment @{ VERIFY_OFFICE_STUB_BEHAVIOR = 'pass'; VERIFY_OFFICE_STUB_KILLS = '1'; PATH = ($script:stubDir + $sep + $env:PATH) }
            $output = (Get-Content -LiteralPath $outFile -Raw) + (Get-Content -LiteralPath $errFile -Raw)
            $proc.ExitCode | Should -Be 0
            $output | Should -Match 'by Quit\(\) outcome'
            $output | Should -Match 'returned'
        }
    }
}
