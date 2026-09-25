using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

// artifact-source: verify-quick.ps1 -> 'publish AddIn (packed XLL)'
// Produced by the 'publish AddIn (packed XLL)' step of
// scripts/verify-quick.ps1 and scripts/verify.ps1 at
// src/GanttCreator.AddIn/bin/Release/net10.0-windows/publish/ (the
// .vscode 'test' task depends on 'publish-addin' for the same reason).
// docs/02-ARCHITECTURE.md build-pipeline artifact contract.

/// <summary>
/// Proves the <see cref="OfficeFixture"/> launches and tears down Excel
/// cleanly. These tests exercise the real COM lifecycle: create an
/// Excel Application, verify it is running, dispose the fixture, and
/// confirm the owned process exits with no orphan.
///
/// Tagged <c>[Trait("Category","OfficeIntegration")]</c> so
/// <c>verify-quick.ps1</c> and <c>verify.ps1</c> exclude them; run via
/// <c>pwsh ./scripts/verify-office.ps1</c> on the self-hosted runner
/// (per ADR-0001).
///
/// All awaited calls use <c>ConfigureAwait(true)</c> because xUnit's
/// analyzer (xUnit1030) requires it for test parallelization, and
/// <c>true</c> satisfies both xUnit1030 and CA2007. Awaits are written
/// as explicit calls (never <c>await using</c>) so CA2007 can observe
/// the <c>ConfigureAwait(true)</c> on the disposal path.
/// </summary>
public class OfficeFixtureTests
{
    private readonly ITestOutputHelper _output;

    public OfficeFixtureTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Excel_launches_and_reports_process_id()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);

            Assert.NotNull(fixture.Excel);
            Assert.True(fixture.ProcessId != 0,
                "Excel launched but process ID was not captured.");

            _output.WriteLine(
                $"Excel launched: PID={fixture.ProcessId}, " +
                $"Version={fixture.Excel.Version}");
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Excel_teardown_leaves_no_orphan_process()
    {
        var pid = 0;
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            pid = fixture.ProcessId;
            Assert.True(pid != 0,
                "Excel launched but process ID was not captured.");
            _output.WriteLine($"Excel launched, PID={pid}");
        }
        finally
        {
            // DisposeAsync runs here, before the orphan poll below.
            await fixture.DisposeAsync().ConfigureAwait(true);
        }

        // Poll until the process exits or the deadline elapses.
        var sw = Stopwatch.StartNew();
        bool exited = false;
        while (sw.Elapsed.TotalSeconds < 10)
        {
            try
            {
                using var proc = Process.GetProcessById(pid);
                if (proc.HasExited) { exited = true; break; }
            }
            catch (ArgumentException)
            {
                exited = true; // process ID not found
                break;
            }
            await Task.Delay(100).ConfigureAwait(true);
        }

        Assert.True(exited,
            $"Excel process {pid} still running after teardown (orphan).");
        _output.WriteLine($"Excel process {pid} exited cleanly after teardown.");
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task CreateWorkbook_returns_a_non_null_workbook_and_leaves_no_orphan()
    {
        var fixture = new OfficeFixture();
        var pid = 0;
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            pid = fixture.ProcessId;
            Assert.True(pid != 0,
                "Excel launched but process ID was not captured.");

            var workbook = fixture.CreateWorkbook();
            Assert.NotNull(workbook);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }

        // Assert the owned Excel process exited after teardown.
        var sw = Stopwatch.StartNew();
        bool exited = false;
        while (sw.Elapsed.TotalSeconds < 10)
        {
            try
            {
                using var proc = Process.GetProcessById(pid);
                if (proc.HasExited) { exited = true; break; }
            }
            catch (ArgumentException)
            {
                exited = true;
                break;
            }
            await Task.Delay(100).ConfigureAwait(true);
        }

        Assert.True(exited,
            $"Excel process {pid} still running after teardown (orphan).");
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task CreateWorkbook_closes_cleanly_when_caller_closes_workbook_first()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);

            var workbook = fixture.CreateWorkbook();
            workbook.Close(SaveChanges: false);
            fixture.MarkWorkbookClosed(workbook);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task CreateWorkbook_can_be_called_twice_and_both_teardown_cleanly()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);

            var first = fixture.CreateWorkbook();
            var second = fixture.CreateWorkbook();
            Assert.NotNull(first);
            Assert.NotNull(second);
            Assert.NotSame(first, second);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Excel_open_close_five_times_no_orphan()
    {
        var pids = new List<int>();

        for (int i = 1; i <= 5; i++)
        {
            var fixture = new OfficeFixture();
            try
            {
                await fixture.InitializeAsync().ConfigureAwait(true);

                var pid = fixture.ProcessId;
                Assert.True(pid != 0,
                    $"Cycle {i}: Excel launched but process ID was not captured.");

                pids.Add(pid);
                _output.WriteLine($"Cycle {i}: Excel launched, PID={pid}");
            }
            finally
            {
                // Dispose runs at the end of each cycle, not at test end.
                await fixture.DisposeAsync().ConfigureAwait(true);
            }
        }

        // After all five cycles, verify every owned process exited.
        // Poll with a deadline rather than asserting immediately, because
        // Excel may take a moment to fully exit after Quit().
        var sw = Stopwatch.StartNew();
        bool anyOrphan;
        do
        {
            anyOrphan = false;
            foreach (var pid in pids)
            {
                try
                {
                    using var proc = Process.GetProcessById(pid);
                    if (!proc.HasExited) { anyOrphan = true; break; }
                }
                catch (ArgumentException)
                {
                    // This PID has exited.
                }
            }
            if (!anyOrphan) break;
            await Task.Delay(200).ConfigureAwait(true);
        } while (sw.Elapsed.TotalSeconds < 15);

        Assert.False(anyOrphan,
            $"One or more Excel processes survived five open/close " +
            $"cycles (orphans: {string.Join(", ", pids)}).");
        _output.WriteLine("All five Excel processes exited cleanly.");
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Packed_XLL_loads_once_and_produces_an_attributable_open_record()
    {
        // R1.6 attributable gate (work item R1.6 D3/D5): load the packaged XLL
        // into real Excel via Application.RegisterXLL, so real Excel-DNA
        // initialization runs, and assert the session-bearing packed-XLL 'open'
        // record that proves this load actually initialised.
        //
        // Scope note: the session token attributes the open record to one load,
        // but AutoClose still produces no 'close' record on the automation path
        // (D6 finding) -- the closing assertion is the owned-process exit. Full
        // close-callback proof remains pending.
        //
        // Why one cycle and not five: the wait below is the test's whole cost
        // (~22 s per cycle, the deferred initialization the log comment above
        // records), and it proves the same thing every time -- that a load
        // initialises and writes one new session token. Repeating it five times
        // bought nothing on the close side, which is unobservable here. Five-cycle
        // open/close no-orphan coverage is already carried by
        // Excel_open_close_five_times_no_orphan, and add-in reopen/session
        // re-arming is pinned by the D2 contract test in AddInHostTests.
        var xllPath = ResolvePackedXllPath();
        var logPath = GetAddInLogPath();
        var fixture = new OfficeFixture();
        int pid;

        // Negative control: seed the seen set with every session already in the
        // log, so the wait cannot be satisfied by a foreign record that predates
        // this test -- the record must belong to this load.
        var seenSessions = PackedOpenSessions(TryReadLog(logPath));
        _output.WriteLine(
            $"Pre-existing packed-XLL sessions in the log: {seenSessions.Count}.");

        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);

            pid = fixture.ProcessId;
            Assert.True(pid != 0, "Excel launched but process ID was not captured.");

            var registered = fixture.RegisterXll(xllPath);
            Assert.True(registered,
                $"Application.RegisterXLL returned false for '{xllPath}'.");

            // No open record is expected yet: initialization is deferred until
            // Quit (see the observation note below).
        }
        finally
        {
            // Quit is what makes the deferred AutoOpen record land.
            await fixture.DisposeAsync().ConfigureAwait(true);
        }

        var newSession = await WaitForNewPackedOpenSessionAsync(logPath, seenSessions, 1)
            .ConfigureAwait(true);
        _output.WriteLine(
            $"PID={pid}; RegisterXLL=True; session-bearing packed-XLL 'open' record " +
            $"observed after quit (session={newSession}).");

        // The load's Excel must be gone; a survivor here is what holds the packed
        // XLL and breaks the next run's publish step.
        var sw = Stopwatch.StartNew();
        var exited = false;
        while (sw.Elapsed.TotalSeconds < 120)
        {
            try
            {
                using var proc = Process.GetProcessById(pid);
                if (!proc.HasExited) { await Task.Delay(200).ConfigureAwait(true); continue; }
                exited = true;
            }
            catch (ArgumentException)
            {
                exited = true;
            }

            break;
        }

        Assert.True(exited, $"Excel process {pid} survived the XLL load (orphan).");
        _output.WriteLine(
            $"Excel process {pid} exited cleanly after {sw.Elapsed.TotalSeconds:F1} s.");
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public void RecordOwnedProcessId_writes_the_owned_pid_to_the_manifest()
    {
        // Pure unit test of the fixture-to-shell handoff: no Excel, no COM.
        var dir = Directory.CreateTempSubdirectory();
        var manifestPath = Path.Combine(dir.FullName, "owned-office-pids.json");
        var previous = Environment.GetEnvironmentVariable("GANTTCREATOR_OWNED_PIDS_PATH");
        try
        {
            Environment.SetEnvironmentVariable("GANTTCREATOR_OWNED_PIDS_PATH", manifestPath);

            OfficeFixture.RecordOwnedProcessIdForTest(111);
            OfficeFixture.RecordOwnedProcessIdForTest(222);

            var lines = File.ReadAllLines(manifestPath);
            Assert.Equal(2, lines.Length);
            Assert.Contains("111", lines[0], StringComparison.Ordinal);
            Assert.Contains("222", lines[1], StringComparison.Ordinal);
        }
        finally
        {
            Environment.SetEnvironmentVariable("GANTTCREATOR_OWNED_PIDS_PATH", previous);
            try { Directory.Delete(dir.FullName, true); }
            catch (IOException) { }
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public void RecordOwnedProcessId_is_a_noop_when_no_manifest_path_is_set()
    {
        var previous = Environment.GetEnvironmentVariable("GANTTCREATOR_OWNED_PIDS_PATH");
        try
        {
            Environment.SetEnvironmentVariable("GANTTCREATOR_OWNED_PIDS_PATH", null);
            // Must not throw, and must not write anywhere.
            OfficeFixture.RecordOwnedProcessIdForTest(333);
        }
        finally
        {
            Environment.SetEnvironmentVariable("GANTTCREATOR_OWNED_PIDS_PATH", previous);
        }
    }

    [Fact]
    public void ReadStartTimeUtcTicks_reads_a_live_process_start_time()
    {
        // The manifest's identity proof is the PID plus the start time. This
        // process is certainly alive, so its start time must be readable and
        // non-zero; the sweep compares against exactly this value.
        using var self = System.Diagnostics.Process.GetCurrentProcess();

        long ticks = OfficeFixture.ReadStartTimeUtcTicks(self.Id);

        Assert.NotEqual(0, ticks);
        Assert.Equal(self.StartTime.ToUniversalTime().Ticks, ticks);
    }

    [Fact]
    public void ReadStartTimeUtcTicks_returns_zero_for_a_process_that_does_not_exist()
    {
        // Fail-safe: an unidentifiable process records zero, and the sweep's
        // safe direction is to skip a zero rather than kill it.
        Assert.Equal(0, OfficeFixture.ReadStartTimeUtcTicks(int.MaxValue));
    }

    [Fact]
    public void KillOwnedProcess_never_targets_a_zero_process_id()
    {
        // The default seam refuses a zero PID before touching any process, so a
        // fixture that never identified its Excel cannot be handed a kill. The
        // base method is called directly: the recording subclass's override
        // would bypass the guard this pins.
        var fixture = new OfficeFixture();

        Assert.False(fixture.KillOwnedProcess(0));
    }

    [Fact]
    public void KillOwnedProcess_refuses_a_recycled_pid_whose_start_time_no_longer_matches()
    {
        // A PID alone is not identity: Windows recycles process IDs, so a PID
        // recorded at launch can name a different process by teardown. Killing on
        // the PID alone would terminate a process this fixture never started, so
        // the default kill must refuse when the recorded start time disagrees.
        // The current process is used because it is certainly alive; the check is
        // driven through the identity seam rather than a real kill so the test
        // cannot terminate the test host.
        var fixture = new OfficeFixture();
        using var self = Process.GetCurrentProcess();
        fixture.OwnedProcessIdsForTest = [self.Id];

        // A deliberately wrong recorded start time: the live process reports its
        // real one, so the two halves of the identity proof disagree.
        fixture.RememberOwnedProcessStartTimeForTest(self.Id, long.MaxValue);

        Assert.False(fixture.IsOwnedProcessIdentityForTest(self.Id));
        Assert.False(fixture.KillOwnedProcess(self.Id));
    }

    [Fact]
    public void KillOwnedProcess_refuses_a_pid_with_no_recorded_start_time()
    {
        // Fail-safe direction, matching the shell sweep: a PID with no verifiable
        // start time is not proven to be the harness's, so it is skipped rather
        // than killed. int.MaxValue does not exist, so adoption records zero.
        var fixture = new OfficeFixture();
        fixture.OwnedProcessIdsForTest = [int.MaxValue];

        Assert.False(fixture.IsOwnedProcessIdentityForTest(int.MaxValue));
        Assert.False(fixture.KillOwnedProcess(int.MaxValue));
    }

    [Fact]
    public void KillOwnedProcess_still_accepts_a_verified_owned_pid()
    {
        // The hardening must not disable the escalation that keeps a leaked Excel
        // from holding the packed XLL. A PID whose recorded start time still matches
        // the live process is proven to be ours, so the identity check passes and
        // only the PID guard remains.
        var fixture = new OfficeFixture();
        using var self = Process.GetCurrentProcess();
        fixture.OwnedProcessIdsForTest = [self.Id];

        Assert.True(fixture.IsOwnedProcessIdentityForTest(self.Id));
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Teardown_escalates_to_a_forced_kill_when_the_owned_process_survives()
    {
        // The leak this covers: unreleased COM proxies keep the Application's
        // reference count above zero, so Quit() returns without terminating the
        // process, and that process holds the packed XLL open. Teardown must
        // escalate to a kill of its own PID rather than waiting out the deadline
        // and leaving the orphan behind.
        //
        // Liveness is driven through the IsProcessRunning seam and the process
        // only stops reporting alive once the kill is issued, so the test needs no
        // real process to keep alive and cannot deadlock on a child pipe.
        var fixture = new StubbornFixture { LivePid = 4242 };

        await fixture.DisposeStubbornAsync().ConfigureAwait(true);

        Assert.Equal([4242], fixture.KilledProcessIds);
        Assert.Equal(1, fixture.ForcedKillCount);
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Teardown_reports_a_process_that_outlives_the_forced_kill()
    {
        // A kill that does not take effect must not be tolerated silently: the
        // surviving process would hold the packed XLL and break the next run.
        var fixture = new UnkillableFixture { LivePid = 77 };

        InvalidOperationException failure = await Assert.ThrowsAsync<InvalidOperationException>(
            () => fixture.DisposeStubbornAsync()).ConfigureAwait(true);

        Assert.Contains("77", failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_com_scope_tracks_proxies_and_releases_them_on_dispose()
    {
        // The scope is the mechanism for retiring the leak at source. It must
        // hold what it was given, hand it back for use, and empty itself on
        // dispose; a proxy left tracked would keep the reference count up and
        // reintroduce the very leak it exists to prevent. A non-COM argument is
        // fine here: ReleaseComObject rejects it and the scope swallows that, so
        // a plain object still proves the track/dispose bookkeeping.
        var scope = new OfficeFixture.ComScope();

        var tracked = scope.Track(new object());

        Assert.NotNull(tracked);
        Assert.Equal(1, scope.TrackedCount);

        scope.Dispose();

        Assert.Equal(0, scope.TrackedCount);
    }

    [Fact]
    public void A_com_scope_rejects_a_null_proxy() =>
        Assert.Throws<ArgumentNullException>(
            () => new OfficeFixture.ComScope().Track<object>(null!));

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Teardown_does_not_escalate_when_the_process_exits_on_its_own()
    {
        // The common case must stay unforced: a process that exits on its own
        // never reaches the kill, so ForcedKillCount stays zero and the leak
        // signal is not polluted.
        var fixture = new WellBehavedFixture { LivePid = 88 };

        await fixture.DisposeStubbornAsync().ConfigureAwait(true);

        Assert.Equal(0, fixture.ForcedKillCount);
    }

    /// <summary>
    /// A fixture whose owned process never exits on its own, so the escalation
    /// path runs without a real Excel and without a real kill.
    /// </summary>
    private class StubbornFixture : OfficeFixture
    {
        public List<int> KilledProcessIds { get; } = [];

        public int LivePid { get; set; }

        public StubbornFixture()
        {
            // This fixture exists to force an escalation, so its kill is
            // deliberate and must not count against the leak ratchet.
            SuppressLeakSignal = true;
        }

        /// <summary>
        /// Shortened so the escalation test reaches the kill without first
        /// spending the production 10 s window.
        /// </summary>
        internal override TimeSpan GracefulExitTimeout => TimeSpan.FromMilliseconds(150);

        internal override TimeSpan ForcedExitTimeout => TimeSpan.FromMilliseconds(150);

        public async Task DisposeStubbornAsync()
        {
            OwnedProcessIdsForTest = [LivePid];
            await DisposeAsync().ConfigureAwait(true);
        }

        /// <summary>Always reports the process alive, until the kill is issued.</summary>
        internal override bool IsProcessRunning(int processId) => !KilledProcessIds.Contains(processId);

        /// <summary>Records the kill; the liveness seam reacts to the record.</summary>
        internal override bool KillOwnedProcess(int processId)
        {
            KilledProcessIds.Add(processId);
            return processId != 0;
        }
    }

    /// <summary>A fixture whose kill never takes effect, so a survivor is reported.</summary>
    private sealed class UnkillableFixture : StubbornFixture
    {
        internal override bool KillOwnedProcess(int processId)
        {
            KilledProcessIds.Add(processId);
            return true;
        }

        internal override bool IsProcessRunning(int processId) => true;
    }

    /// <summary>A fixture whose process exits on its own, so no escalation happens.</summary>
    private sealed class WellBehavedFixture : StubbornFixture
    {
        internal override bool IsProcessRunning(int processId) => false;
    }

    [Fact]
    public void A_manifest_record_carries_the_start_time_that_identifies_its_process()
    {
        // A PID alone is not identity: Windows recycles PIDs between runs, so the
        // manifest must record the start time the sweep later re-checks.
        var dir = Directory.CreateTempSubdirectory();
        var manifestPath = Path.Combine(dir.FullName, "owned-office-pids.json");
        var previous = Environment.GetEnvironmentVariable("GANTTCREATOR_OWNED_PIDS_PATH");
        try
        {
            Environment.SetEnvironmentVariable("GANTTCREATOR_OWNED_PIDS_PATH", manifestPath);
            using var self = System.Diagnostics.Process.GetCurrentProcess();

            OfficeFixture.RecordOwnedProcessIdForTest(self.Id);

            var line = File.ReadAllLines(manifestPath).Single();
            using var record = System.Text.Json.JsonDocument.Parse(line);
            Assert.Equal(self.Id, record.RootElement.GetProperty("ProcessId").GetInt32());
            Assert.Equal(
                self.StartTime.ToUniversalTime().Ticks,
                record.RootElement.GetProperty("StartTimeUtcTicks").GetInt64());
        }
        finally
        {
            Environment.SetEnvironmentVariable("GANTTCREATOR_OWNED_PIDS_PATH", previous);
            try { Directory.Delete(dir.FullName, true); }
            catch (IOException) { }
        }
    }

    [Fact]
    public void Each_manifest_line_is_independently_parseable_json()
    {
        // The reader parses line by line. A whole-file ConvertFrom-Json fails on
        // concatenated objects, which is what silently emptied the sweep's primary
        // ownership signal before the start-time change.
        var dir = Directory.CreateTempSubdirectory();
        var manifestPath = Path.Combine(dir.FullName, "owned-office-pids.json");
        var previous = Environment.GetEnvironmentVariable("GANTTCREATOR_OWNED_PIDS_PATH");
        try
        {
            Environment.SetEnvironmentVariable("GANTTCREATOR_OWNED_PIDS_PATH", manifestPath);
            OfficeFixture.RecordOwnedProcessIdForTest(11);
            OfficeFixture.RecordOwnedProcessIdForTest(22);
            OfficeFixture.RecordOwnedProcessIdForTest(33);

            var parsed = File.ReadAllLines(manifestPath)
                .Select(line => System.Text.Json.JsonDocument.Parse(line).RootElement.GetProperty("ProcessId").GetInt32())
                .ToArray();

            Assert.Equal([11, 22, 33], parsed);
        }
        finally
        {
            Environment.SetEnvironmentVariable("GANTTCREATOR_OWNED_PIDS_PATH", previous);
            try { Directory.Delete(dir.FullName, true); }
            catch (IOException) { }
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public void RecordOwnedProcessId_is_a_noop_for_zero_process_id()
    {
        var dir = Directory.CreateTempSubdirectory();
        var manifestPath = Path.Combine(dir.FullName, "owned-office-pids.json");
        var previous = Environment.GetEnvironmentVariable("GANTTCREATOR_OWNED_PIDS_PATH");
        try
        {
            Environment.SetEnvironmentVariable("GANTTCREATOR_OWNED_PIDS_PATH", manifestPath);
            OfficeFixture.RecordOwnedProcessIdForTest(0);
            Assert.False(File.Exists(manifestPath));
        }
        finally
        {
            Environment.SetEnvironmentVariable("GANTTCREATOR_OWNED_PIDS_PATH", previous);
            try { Directory.Delete(dir.FullName, true); }
            catch (IOException) { }
        }
    }

    /// <summary>
    /// Resolves the packed XLL the publish pipeline produces. This test never
    /// builds artifacts (build-pipeline traceability): the XLL must already
    /// exist from <c>scripts/verify-quick.ps1</c>'s "publish AddIn (packed
    /// XLL)" step (or an explicit <c>dotnet publish</c> of the AddIn project).
    /// </summary>
    internal static string ResolvePackedXllPath()
    {
        // Walk up from the test host's base directory to the repository root,
        // identified by the solution file, so the path survives relocation.
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null && !File.Exists(Path.Combine(dir.FullName, "GanttCreator.slnx")))
        {
            dir = dir.Parent;
        }

        Assert.NotNull(dir);
        var xllPath = Path.Combine(
            dir.FullName, "src", "GanttCreator.AddIn", "bin", "Release",
            "net10.0-windows", "publish", "GanttCreator.AddIn-AddIn64-packed.xll");
        Assert.True(
            File.Exists(xllPath),
            "The packed XLL was not found. Produce it first with " +
            "'pwsh ./scripts/verify-quick.ps1' (its 'publish AddIn (packed XLL)' step) " +
            $"or 'dotnet publish src/GanttCreator.AddIn -c Release'. Expected at: {xllPath}");
        return xllPath;
    }

    /// <summary>
    /// The add-in rolling log path: the production location
    /// <c>AddInLogFactory</c> writes to (<c>%LOCALAPPDATA%\GanttCreator\logs</c>).
    /// </summary>
    internal static string GetAddInLogPath() =>
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "GanttCreator", "logs", "gantt-creator-addin.log");

    /// <summary>
    /// Extracts the session tokens of packed-XLL <c>open</c> records: lines
    /// carrying the lifecycle record shape whose <c>xll=</c> field names the
    /// packed XLL and whose <c>session=</c> field carries a non-empty token.
    /// Pure and culture-free so the non-Office unit tests can pin it against
    /// the real record shape.
    /// </summary>
    internal static HashSet<string> PackedOpenSessions(string? logContent)
    {
        var sessions = new HashSet<string>(StringComparer.Ordinal);
        if (string.IsNullOrEmpty(logContent))
        {
            return sessions;
        }

        foreach (var line in logContent.Split('\n'))
        {
            if (!line.Contains(" open addin-version=", StringComparison.Ordinal) ||
                !line.Contains("xll=GanttCreator.AddIn-AddIn64-packed.xll", StringComparison.Ordinal))
            {
                continue;
            }

            var marker = "session=";
            var start = line.IndexOf(marker, StringComparison.Ordinal);
            if (start < 0)
            {
                continue;
            }

            var token = line.Substring(start + marker.Length).Trim();
            if (!string.IsNullOrEmpty(token))
            {
                sessions.Add(token);
            }
        }

        return sessions;
    }

    /// <summary>
    /// Counts the AutoOpen <c>open</c> records written for the packed XLL:
    /// lines carrying the lifecycle record shape whose <c>xll=</c> field names
    /// the packed XLL. Pure and culture-free so the non-Office unit tests can
    /// pin it against the real record shape.
    /// </summary>
    internal static int CountPackedOpenRecords(string? logContent) =>
        PackedOpenSessions(logContent).Count;

    /// <summary>
    /// Reads the log file without throwing while Excel holds it open
    /// (RollingLog opens its writer with read sharing). A transient read
    /// failure yields an empty string; the caller polls with a deadline.
    /// </summary>
    private static string TryReadLog(string path)
    {
        try
        {
            // FileShare.ReadWrite: Windows share checks are bidirectional, so a
            // reader requesting FileShare.Read (what File.ReadAllText uses) is
            // rejected while any Excel instance still holds RollingLog's writer
            // handle (FileAccess.Write). The swallowed IOException would
            // silently zero the pre-existing session baseline (Five_cycles XLL
            // failure, R2.7a risk closure).
            using var stream = new FileStream(
                path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = new StreamReader(stream);
            return reader.ReadToEnd();
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return string.Empty;
        }
    }

    /// <summary>
    /// Polls the add-in log until a session-bearing packed-XLL <c>open</c>
    /// record appears whose session token was not seen before this cycle.
    /// The token attributes the record to one load: a foreign session (for
    /// example an interleaved record from another Excel instance) does not
    /// satisfy the wait, and a missing record fails with a diagnostic. The
    /// deadline is 60 s after the Excel instance quit. Returns the new
    /// session token. The exact per-cycle session is recorded in the test
    /// output for the human evidence review.
    /// </summary>
    private static async Task<string> WaitForNewPackedOpenSessionAsync(
        string logPath, HashSet<string> seenSessions, int cycle)
    {
        var sw = Stopwatch.StartNew();
        while (sw.Elapsed.TotalSeconds < 90)
        {
            var sessions = PackedOpenSessions(TryReadLog(logPath));
            sessions.ExceptWith(seenSessions);
            if (sessions.Count > 0)
            {
                return sessions.OrderBy(s => s, StringComparer.Ordinal).First();
            }

            await Task.Delay(500).ConfigureAwait(true);
        }

        Assert.Fail(
            $"Cycle {cycle}: no new session-bearing packed-XLL 'open' record appeared in " +
            $"'{logPath}' within 90 s after the Excel instance quit (sessions before: {seenSessions.Count}). " +
            "RegisterXLL returned true, so the XLL loaded; in the 2026-09-17 session-token probe " +
            "the matching record materialized within ~30 s of RegisterXLL — its absence after 90 s " +
            "means no attributable open record appeared for this load.");
        throw new InvalidOperationException("Unreachable: Assert.Fail throws.");
    }

    // Non-Office unit tests for the R1.6 log matcher. These carry no
    // OfficeIntegration trait, so verify-quick's test step runs them; they
    // pin the matcher against the real RollingLog record shape (timestamp-
    // prefixed lines, as produced by the 2026-09-16 RegisterXLL spike).

    [Fact]
    public void CountPackedOpenRecords_counts_the_real_packed_xll_open_record_shape()
    {
        const string content =
            "2026-09-16T21:44:19.135Z open addin-version=0.0.0+local.abc123 " +
            "excel-version=16.0 process=x64 xll=GanttCreator.AddIn-AddIn64-packed.xll session=s99-g0h1i2j3k4l5\r\n";

        Assert.Equal(1, CountPackedOpenRecords(content));
    }

    [Fact]
    public void PackedOpenSessions_ignores_sessionless_records_other_xlls_and_close_records()
    {
        // Negative control: a sessionless legacy open record, another XLL's
        // session-bearing record, a Diagnostics record, and a close record
        // must all be rejected — only the packed XLL's session-bearing open
        // record counts.
        const string content =
            "2026-09-16T21:44:19.135Z open addin-version=0.0.0+local.abc123 " +
            "excel-version=16.0 process=x64 xll=GanttCreator.AddIn-AddIn64-packed.xll\r\n" +
            "2026-09-16T07:28:47.543Z open addin-version=0.0.0+local.abc123 " +
            "excel-version=16.0 process=x64 xll=GanttCreator.AddIn-AddIn64.xll session=s1-000000000000\r\n" +
            "2026-09-16T07:28:58.769Z Diagnostics: addin-version=0.0.0 " +
            "excel-version=16.0 process=x64 xll=GanttCreator.AddIn-AddIn64-packed.xll\r\n" +
            "2026-09-16T21:44:19.135Z close session=s1-000000000000\r\n";

        Assert.Equal(0, CountPackedOpenRecords(content));
        Assert.Empty(PackedOpenSessions(content));
    }

    [Fact]
    public void PackedOpenSessions_collects_each_session_once_and_ignores_empty_tokens()
    {
        const string one =
            "2026-09-16T21:44:19.135Z open addin-version=0.0.0 xll=GanttCreator.AddIn-AddIn64-packed.xll session=s1-g0h1i2j3k4l5\r\n";
        const string emptyToken =
            "2026-09-16T21:44:19.135Z open addin-version=0.0.0 xll=GanttCreator.AddIn-AddIn64-packed.xll session=\r\n";
        var sessions = PackedOpenSessions(one + one + emptyToken);
        Assert.Equal(new HashSet<string>(["s1-g0h1i2j3k4l5"], StringComparer.Ordinal), sessions);

        Assert.Equal(0, CountPackedOpenRecords(null));
        Assert.Equal(0, CountPackedOpenRecords(string.Empty));
        Assert.Empty(PackedOpenSessions(null));
        Assert.Empty(PackedOpenSessions(string.Empty));
    }

    [Fact]
    public void TryReadLog_reads_the_log_while_a_writer_holds_it_open()
    {
        // Regression (bidirectional Windows share checks; R2.7a Five_cycles
        // XLL failure): the poll reads the add-in log while an Excel
        // instance still holds RollingLog's writer handle (FileAccess.Write).
        // A reader opening with FileShare.Read is rejected by that write
        // access, the exception is swallowed to string.Empty, and the
        // pre-existing session baseline silently becomes 0.
        var path = Path.Combine(
            Path.GetTempPath(), $"gantt-creator-logread-{Guid.NewGuid():N}.log");
        try
        {
            const string record =
                "2026-09-23T00:00:00.000Z open addin-version=0.0.0 " +
                "xll=GanttCreator.AddIn-AddIn64-packed.xll session=s-read-probe\r\n";
            File.WriteAllText(path, record);

            using (var writer = new FileStream(
                path, FileMode.Append, FileAccess.Write, FileShare.Read))
            {
                _ = writer;
                var content = TryReadLog(path);
                Assert.Contains(
                    "session=s-read-probe", content, StringComparison.Ordinal);
            }
        }
        finally
        {
            try
            {
                File.Delete(path);
            }
            catch (IOException)
            {
                // Best-effort cleanup.
            }
            catch (UnauthorizedAccessException)
            {
                // Best-effort cleanup.
            }
        }
    }
}
