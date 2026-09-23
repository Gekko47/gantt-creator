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
    public async Task Five_cycles_registering_the_packed_XLL_leave_no_orphan_and_show_session_open_record_growth()
    {
        // R1.6 attributable gate (work item R1.6 D3/D5): the five open/close
        // cycles run with the add-in's packed XLL loaded via
        // Application.RegisterXLL, so real Excel-DNA initialization runs
        // inside Excel each cycle. Each cycle asserts RegisterXLL success,
        // the appearance of a session-bearing packed-XLL 'open' record for
        // that cycle's session token, and owned-process exit.
        //
        // Scope note: the session token correlates the open record to one
        // load, but AutoClose still produces no 'close' record on the
        // automation path (D6 finding) — the closing assertion is the owned
        // orphan poll. Full close-callback proof remains pending.
        //
        // Observed Office behaviour (2026-09-16 spike, three runs, Microsoft
        // 365 16.0.20326 x64; 2026-09-17 session-token probe): RegisterXLL
        // returns true immediately, and the matching 'open' log record
        // materializes ~5-30 s after RegisterXLL, while the Excel instance is
        // still alive or during shutdown. The wait therefore runs AFTER
        // DisposeAsync (Quit + exit poll): once the process has exited, the
        // flushed record must be on disk.
        var xllPath = ResolvePackedXllPath();
        var logPath = GetAddInLogPath();
        var pids = new List<int>();
        // Negative control: seed the seen set with every session already in
        // the log, so cycle 1 cannot be satisfied by a foreign record that
        // predates the test — each cycle must produce a NEW session token.
        var seenSessions = PackedOpenSessions(TryReadLog(logPath));
        _output.WriteLine(
            $"Pre-existing packed-XLL sessions in the log: {seenSessions.Count}.");

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

                var registered = fixture.RegisterXll(xllPath);
                Assert.True(registered,
                    $"Cycle {i}: Application.RegisterXLL returned false for '{xllPath}'.");

                // No open record is expected yet: initialization is deferred
                // until Quit (see the observation note above).
            }
            finally
            {
                // Dispose runs at the end of each cycle, not at test end;
                // Quit is what makes the deferred AutoOpen record land.
                await fixture.DisposeAsync().ConfigureAwait(true);
            }

            var newSession = await WaitForNewPackedOpenSessionAsync(
                logPath, seenSessions, i).ConfigureAwait(true);
            seenSessions.Add(newSession);
            _output.WriteLine(
                $"Cycle {i}: PID={pids[i - 1]}; RegisterXLL=True; session-bearing packed-XLL 'open' " +
                $"record observed after quit (session={newSession}; sessions so far: {seenSessions.Count}).");
        }

        // After all five cycles, verify every owned process exited. Poll with
        // a deadline rather than asserting immediately, because Excel may take
        // a moment to fully exit after Quit().
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
        // Evidence 2026-09-23 (verify-office run 22:07, archive
        // office-20260923-221210367.trx): all five cycle PIDs exited
        // cleanly on their own, but XLL-loaded instances outlive the
        // plain test's 15 s window — every PID was still alive at 15 s
        // and confirmed dead when checked minutes later. Deadline widened
        // to 120 s so the observed minutes-scale clearance is covered with
        // margin (well under the 600 s suite deadline); the zero-survivor
        // assertion is unchanged.
        } while (sw.Elapsed.TotalSeconds < 120);

        Assert.False(anyOrphan,
            $"One or more Excel processes survived five XLL-loaded open/close " +
            $"cycles (orphans: {string.Join(", ", pids)}).");
        _output.WriteLine(
            $"All five XLL-loaded Excel processes exited cleanly after {sw.Elapsed.TotalSeconds:F1} s; " +
            $"total packed-XLL open records: {CountPackedOpenRecords(TryReadLog(logPath))}.");
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
