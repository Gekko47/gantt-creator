using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

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
}
