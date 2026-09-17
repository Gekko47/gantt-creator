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
    public async Task Five_cycles_registering_the_packed_XLL_leave_no_orphan_and_show_shared_open_record_growth()
    {
        // R1.6 smoke test (work item R1.6 D3): the five open/close cycles run
        // with the add-in's packed XLL loaded via Application.RegisterXLL, so
        // real Excel-DNA initialization runs inside Excel each cycle.
        //
        // Narrowed smoke-test scope: each cycle asserts RegisterXLL success,
        // growth in the shared packed-XLL open-record count, and owned-process
        // exit. A growing shared count does NOT attribute a record to one
        // owned instance, and growth after Quit does NOT establish that
        // AutoClose ran (deferred initialization is an unverified inference,
        // not an established fact). The final orphan poll is the only closing
        // assertion; AutoClose host evidence remains pending (see the pending
        // work item below).
        //
        // Observed Office behaviour (2026-09-16 spike, three runs, Microsoft
        // 365 16.0.20326 x64): RegisterXLL returns true immediately, but the
        // matching 'open' log record materializes only when the Excel instance
        // quits. The wait therefore runs AFTER DisposeAsync (Quit + exit
        // poll): once the process has exited, the flushed record must be on
        // disk. AutoClose produces no 'close' record on this path.
        var xllPath = ResolvePackedXllPath();
        var logPath = GetAddInLogPath();
        var pids = new List<int>();
        var packedOpenRecordsBefore = CountPackedOpenRecords(TryReadLog(logPath));

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

            var countAfter = await WaitForNewPackedOpenRecordAsync(
                logPath, packedOpenRecordsBefore, i).ConfigureAwait(true);
            _output.WriteLine(
                $"Cycle {i}: PID={pids[i - 1]}; RegisterXLL=True; shared packed-XLL 'open' " +
                $"record count grew after quit (shared records: {packedOpenRecordsBefore} -> {countAfter}).");
            packedOpenRecordsBefore = countAfter;
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
        } while (sw.Elapsed.TotalSeconds < 15);

        Assert.False(anyOrphan,
            $"One or more Excel processes survived five XLL-loaded open/close " +
            $"cycles (orphans: {string.Join(", ", pids)}).");
        _output.WriteLine(
            "All five XLL-loaded Excel processes exited cleanly; " +
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
    /// Counts the AutoOpen <c>open</c> records written for the packed XLL:
    /// lines carrying the lifecycle record shape whose <c>xll=</c> field names
    /// the packed XLL. Pure and culture-free so the non-Office unit tests can
    /// pin it against the real record shape.
    /// </summary>
    internal static int CountPackedOpenRecords(string? logContent)
    {
        if (string.IsNullOrEmpty(logContent))
        {
            return 0;
        }

        var count = 0;
        foreach (var line in logContent.Split('\n'))
        {
            if (line.Contains(" open addin-version=", StringComparison.Ordinal) &&
                line.Contains("xll=GanttCreator.AddIn-AddIn64-packed.xll", StringComparison.Ordinal))
            {
                count++;
            }
        }

        return count;
    }

    /// <summary>
    /// Reads the log file without throwing while Excel holds it open
    /// (RollingLog opens its writer with read sharing). A transient read
    /// failure yields an empty string; the caller polls with a deadline.
    /// </summary>
    private static string TryReadLog(string path)
    {
        try
        {
            return File.ReadAllText(path);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return string.Empty;
        }
    }

    /// <summary>
    /// Polls the add-in log until the shared packed-XLL <c>open</c> record
    /// count grows. This is a smoke signal only: a growing shared count does
    /// NOT attribute a record to one owned instance and does NOT establish
    /// that AutoClose ran. Returns the new total record count. The deadline is
    /// 60 s after the Excel instance quit: the spiked record-growth behavior
    /// materializes during shutdown, observed 0–20 s after <c>Quit()</c>
    /// across all 2026-09-16 spike and gate runs, so 60 s gives a comfortable
    /// margin over the measured worst case. Tolerates an interleaved record
    /// from another session by asserting the count grew rather than grew by
    /// exactly one; the exact per-cycle count is recorded in the test output
    /// for the human evidence review.
    /// </summary>
    private static async Task<int> WaitForNewPackedOpenRecordAsync(
        string logPath, int previousCount, int cycle)
    {
        var sw = Stopwatch.StartNew();
        while (sw.Elapsed.TotalSeconds < 60)
        {
            var count = CountPackedOpenRecords(TryReadLog(logPath));
            if (count > previousCount)
            {
                return count;
            }

            await Task.Delay(500).ConfigureAwait(true);
        }

        Assert.Fail(
            $"Cycle {cycle}: the shared packed-XLL 'open' record count did not grow in " +
            $"'{logPath}' within 60 s after the Excel instance quit (records before: {previousCount}). " +
            "RegisterXLL returned true, so the XLL loaded; in the 2026-09-16 spike runs " +
            "(0-20 s after Quit across six loaded instances) the matching record " +
            "materialized during shutdown - its absence after 60 s means no new " +
            "shared 'open' record appeared, not proof that a lifecycle callback " +
            "was missed by this instance.");
        return previousCount; // Unreachable: Assert.Fail throws.
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
            "excel-version=16.0 process=x64 xll=GanttCreator.AddIn-AddIn64-packed.xll\r\n";

        Assert.Equal(1, CountPackedOpenRecords(content));
    }

    [Fact]
    public void CountPackedOpenRecords_ignores_other_xll_names_and_record_kinds()
    {
        const string content =
            "2026-09-16T07:28:47.543Z open addin-version=0.0.0+local.abc123 " +
            "excel-version=16.0 process=x64 xll=GanttCreator.AddIn-AddIn64.xll\r\n" +
            "2026-09-16T07:28:58.769Z Diagnostics: addin-version=0.0.0 " +
            "excel-version=16.0 process=x64 xll=GanttCreator.AddIn-AddIn64-packed.xll\r\n" +
            "2026-09-16T21:44:19.135Z close\r\n";

        Assert.Equal(0, CountPackedOpenRecords(content));
    }

    [Fact]
    public void CountPackedOpenRecords_counts_each_record_once_and_handles_null_or_empty()
    {
        var one = "2026-09-16T21:44:19.135Z open addin-version=0.0.0 xll=GanttCreator.AddIn-AddIn64-packed.xll\r\n";
        Assert.Equal(2, CountPackedOpenRecords(one + one));
        Assert.Equal(0, CountPackedOpenRecords(null));
        Assert.Equal(0, CountPackedOpenRecords(string.Empty));
    }
}
