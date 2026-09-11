using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// xUnit fixture that safely launches and tears down a real Excel
/// Application instance for Office-integration tests. This is the
/// reusable harness for R1.6's "repeat Excel open/close five times with
/// no add-in error or owned orphan process" gate.
///
/// Lifecycle:
/// - <see cref="InitializeAsync"/>: creates an Excel Application
///   (Visible=false, DisplayAlerts=false, ScreenUpdating=false) and
///   captures the new process ID by diffing the EXCEL.EXE process list
///   before and after launch.
/// - <see cref="DisposeAsync"/>: closes every open workbook without
///   saving, calls Quit(), releases each COM proxy through
///   <see cref="Marshal.ReleaseComObject"/> (one release per proxy,
///   never chained), forces a GC pass to drop Runtime Callable
///   Wrappers, then polls the owned process ID until it exits or a
///   10&nbsp;second deadline elapses.
///
/// Every COM proxy is held in a local variable and released exactly
/// once, per the COM ownership rules in AGENTS.md. No arbitrary
/// <c>Thread.Sleep</c> calls &mdash; the orphan check polls a named
/// observable condition (process exited) with a deadline.
/// </summary>
internal sealed class OfficeFixture : IAsyncLifetime
{
    private Application? _excel;
    private Workbooks? _workbooks;
    private int _excelProcessId;
    private bool _disposed;

    /// <summary>
    /// The live Excel Application instance. Valid only between
    /// <see cref="InitializeAsync"/> and <see cref="DisposeAsync"/>.
    /// </summary>
    /// <exception cref="InvalidOperationException">
    /// Thrown if the fixture has not been initialized yet.
    /// </exception>
    public Application Excel => _excel
        ?? throw new InvalidOperationException(
            "Excel not initialized. Ensure InitializeAsync has run.");

    /// <summary>
    /// Process ID of the owned Excel instance, captured at launch.
    /// Zero if the process could not be identified (orphan check is
    /// best-effort in that case).
    /// </summary>
    public int ProcessId => _excelProcessId;

    /// <inheritdoc />
    public Task InitializeAsync()
    {
        // Snapshot existing Excel processes so we can identify the one
        // this fixture creates.
        var before = Process.GetProcessesByName("EXCEL")
            .Select(p => p.Id)
            .ToHashSet();

        var app = new Application
        {
            Visible = false,
            DisplayAlerts = false,
            ScreenUpdating = false,
        };
        _excel = app;

        // Identify the new EXCEL.EXE process by diffing the snapshot.
        var after = Process.GetProcessesByName("EXCEL")
            .Select(p => p.Id)
            .Where(id => !before.Contains(id))
            .ToList();

        if (after.Count == 1)
        {
            _excelProcessId = after[0];
        }

        return Task.CompletedTask;
    }

    /// <inheritdoc />
    public async Task DisposeAsync()
    {
        if (_disposed) return;
        _disposed = true;

        if (_excel == null) return;

        try
        {
            _workbooks = _excel.Workbooks;

            // Close every open workbook without saving. The Excel COM
            // collection is 1-based; iterate backwards so index shifts
            // from removals do not skip entries.
            var count = _workbooks.Count;
            for (int i = count; i >= 1; i--)
            {
                Workbook wb = _workbooks[i];
                wb.Close(SaveChanges: false);
                Marshal.ReleaseComObject(wb);
            }

            Marshal.ReleaseComObject(_workbooks);
            _workbooks = null;

            _excel.Quit();
            Marshal.ReleaseComObject(_excel);
            _excel = null;
        }
        catch (COMException)
        {
            // Excel may already be shutting down from a prior failure.
            // Best-effort cleanup: drop references and continue to the
            // GC + orphan-poll phase.
            _workbooks = null;
            _excel = null;
        }

        // Force GC to release any lingering Runtime Callable Wrappers
        // that would otherwise keep Excel alive past Quit().
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        await PollForProcessExitAsync(_excelProcessId);
    }

    /// <summary>
    /// Polls the owned Excel process until it exits or the deadline
    /// elapses. Never throws; the test itself asserts on the result.
    /// </summary>
    private static async Task PollForProcessExitAsync(int processId)
    {
        if (processId == 0) return;

        var sw = Stopwatch.StartNew();
        while (sw.Elapsed.TotalSeconds < 10)
        {
            try
            {
                using var proc = Process.GetProcessById(processId);
                if (proc.HasExited) return;
            }
            catch (ArgumentException)
            {
                // Process ID not found &mdash; it exited.
                return;
            }

            await Task.Delay(100).ConfigureAwait(true);
        }
    }
}
