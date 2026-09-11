using System.Diagnostics;
using System.Runtime.ExceptionServices;
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

        var app = new Application();
        _excel = app;

        // Configure Excel properties in a guarded path that can clean up
        // if any setter throws.
        try
        {
            app.Visible = false;
            app.DisplayAlerts = false;
            app.ScreenUpdating = false;
        }
        catch
        {
            // If property configuration fails, ensure we clean up the
            // Application instance we just created.
            // CA1031 is narrowed below: Quit/ReleaseComObject failures during
            // guarded cleanup must not mask the original property-setter error.
#pragma warning disable CA1031
            try { app.Quit(); } catch { }
#pragma warning restore CA1031
            Marshal.ReleaseComObject(app);
            throw;
        }

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

        Exception? cleanupException = null;

        try
        {
            try
            {
                _workbooks = _excel.Workbooks;

                // Close every open workbook without saving. The Excel COM
                // collection is 1-based; iterate backwards so index shifts
                // from removals do not skip entries.
                var count = _workbooks.Count;
                for (int i = count; i >= 1; i--)
                {
                    Workbook? wb = null;
                    try
                    {
                        wb = _workbooks[i];
                        wb.Close(SaveChanges: false);
                    }
                    finally
                    {
                        if (wb != null)
                        {
                            Marshal.ReleaseComObject(wb);
                        }
                    }
                }
            }
            finally
            {
                // Release the Workbooks collection proxy even if workbook
                // cleanup failed.
                if (_workbooks != null)
                {
                    try
                    {
                        Marshal.ReleaseComObject(_workbooks);
                    }
                    finally
                    {
                        _workbooks = null;
                    }
                }
            }
        }
        catch (COMException ex)
        {
            // Excel may already be shutting down from a prior failure.
            // Best-effort cleanup: capture the exception and continue to the
            // GC + orphan-poll phase.
            cleanupException = ex;
            _workbooks = null;
        }
        catch (ArgumentException ex)
        {
            // Workbook index out of range or similar argument issues during
            // cleanup. Best-effort: capture and continue.
            cleanupException = ex;
            _workbooks = null;
        }
        catch (InvalidOperationException ex)
        {
            // Excel application in invalid state during cleanup.
            // Best-effort: capture and continue.
            cleanupException = ex;
            _workbooks = null;
        }
        catch (NotImplementedException ex)
        {
            // COM method not implemented. Best-effort: capture and continue.
            cleanupException = ex;
            _workbooks = null;
        }
        catch (NotSupportedException ex)
        {
            // COM method not supported. Best-effort: capture and continue.
            cleanupException = ex;
            _workbooks = null;
        }

        // Always attempt application shutdown independently, even if
        // workbook cleanup threw. Preserve the first cleanup error, then
        // execute Quit() and ReleaseComObject in a separate phase.
        try
        {
            _excel.Quit();
        }
        catch (COMException ex)
        {
            // Excel may already be shutting down; capture if this is the
            // first error we've seen.
            if (cleanupException == null)
            {
                cleanupException = ex;
            }
        }
#pragma warning disable CA1031 // General-completed: intentionally aggregated; first failure preserved and rethrown via ExceptionDispatchInfo.
        catch (Exception ex)
        {
            if (cleanupException == null)
            {
                cleanupException = ex;
            }
        }
#pragma warning restore CA1031

        // Always attempt to release the Excel Application proxy, even if
        // Quit() failed or was not attempted.
        try
        {
            if (_excel != null)
            {
                Marshal.ReleaseComObject(_excel);
            }
        }
        catch (COMException ex)
        {
            if (cleanupException == null)
            {
                cleanupException = ex;
            }
        }
#pragma warning disable CA1031 // General-completed: intentionally aggregated; first failure preserved and rethrown via ExceptionDispatchInfo.
        catch (Exception ex)
        {
            if (cleanupException == null)
            {
                cleanupException = ex;
            }
        }
#pragma warning restore CA1031
        finally
        {
            _excel = null;
        }

        // Force GC to release any lingering Runtime Callable Wrappers
        // that would otherwise keep Excel alive past Quit().
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        await PollForProcessExitAsync(_excelProcessId).ConfigureAwait(true);

        // Report any cleanup exception after all cleanup attempts complete.
        // Use ExceptionDispatchInfo to preserve the original stack trace
        // from the workbook or Quit() cleanup failure.
        if (cleanupException != null)
        {
            ExceptionDispatchInfo.Capture(cleanupException).Throw();
        }
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
