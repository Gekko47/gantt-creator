using System.Diagnostics;
using System.Globalization;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Text.Json;
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
    /// Workbooks created through <see cref="CreateWorkbook"/>. Closed and
    /// released here during teardown, before the existing collection sweep, so
    /// a test that forgets to clean up still leaves no orphan and no double
    /// close. Each proxy is released exactly once (COM ownership, AGENTS.md).
    /// </summary>
    private List<Workbook>? _createdWorkbooks;

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

    /// <summary>
    /// Loads an XLL code resource into the owned Excel instance via
    /// <c>Application.RegisterXLL</c> (work item R1.6 D3). Excel loads the
    /// resource and registers its entry points; in the 2026-09-16 automation
    /// spike, the matching <c>open</c> record in the add-in rolling log
    /// materialized after the Excel instance quit. Callback timing and record
    /// ownership were not established; treat the record as a shared-log
    /// growth signal, not per-instance lifecycle proof.
    /// </summary>
    /// <param name="xllPath">Full path to the XLL to load.</param>
    /// <returns>
    /// The <c>RegisterXLL</c> Boolean result: <see langword="true"/> when the
    /// code resource was loaded successfully.
    /// </returns>
    /// <remarks>
    /// COM ownership: <c>RegisterXLL</c> returns a Boolean and crosses no COM
    /// proxy; the call runs on the fixture's held <c>Application</c> proxy,
    /// which <see cref="DisposeAsync"/> owns and releases exactly once. The
    /// XLL is not deregistered here — Excel unloads it at <c>Quit()</c>,
    /// which is what the five-cycle orphan gate observes.
    /// </remarks>
    public bool RegisterXll(string xllPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(xllPath);
        Application excel = Excel;
        return excel.RegisterXLL(xllPath);
    }

    /// <summary>
    /// Creates a new blank workbook in the owned Excel instance and tracks it
    /// for deterministic teardown. The returned <see cref="Workbook"/> proxy
    /// is owned by the caller for the test's lifetime; <see cref="DisposeAsync"/>
    /// closes and releases it if the test does not, so a test that forgets to
    /// clean up still leaves no orphan and no double close.
    /// </summary>
    /// <remarks>
    /// <c>Workbooks.Add</c> creates the default sheet set (a blank worksheet
    /// plus, on some builds, a chart sheet). The Initialise-sheet command
    /// adopts the blank active worksheet, so a workbook created here is the
    /// exact pre-state the adopt-path integration test needs. COM ownership:
    /// the returned proxy crosses no collection boundary, is held in a local
    /// by the caller, and is released exactly once — here, on teardown.
    /// </remarks>
    /// <returns>The created <see cref="Workbook"/> proxy.</returns>
    public Workbook CreateWorkbook()
    {
        Workbooks workbooks = Excel.Workbooks;
        try
        {
            Workbook created = workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            if (_createdWorkbooks is null)
            {
                _createdWorkbooks = new List<Workbook>();
            }

            _createdWorkbooks.Add(created);
            return created;
        }
        finally
        {
            Marshal.ReleaseComObject(workbooks);
        }
    }

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

        // Report the owned PID to the verifying shell. On a genuine timeout
        // DisposeAsync never runs and the test host is killed, so the shell
        // needs a signal that outlives the process; the manifest file does.
        // Best-effort: this must never affect test pass/fail.
        RecordOwnedProcessId(_excelProcessId);

        return Task.CompletedTask;
    }

    /// <inheritdoc />
    public async Task DisposeAsync()
    {
        if (_disposed) return;
        _disposed = true;

        if (_excel == null) return;

        Exception? cleanupException = null;

        // Close and release each workbook handed out by CreateWorkbook before
        // the collection sweep. Once Close removes it from Workbooks, the sweep
        // sees only other workbooks that are still open. Release the tracked
        // RCW in finally even when Close fails; the sweep then obtains fresh
        // collection proxies and a fresh post-close count.
        if (_createdWorkbooks is not null)
        {
            foreach (Workbook wb in _createdWorkbooks)
            {
                try
                {
                    wb.Close(SaveChanges: false);
                }
                catch (COMException ex) when (ex.HResult == unchecked((int)0x80010108))
                {
                    // The caller already closed this workbook; the collection
                    // sweep remains responsible for workbooks still open in Excel.
                }
                catch (COMException ex)
                {
                    PreserveFirstCleanupException(ref cleanupException, ex);
                }
#pragma warning disable CA1031
                catch (Exception ex)
                {
                    if (cleanupException is null)
                    {
                        cleanupException = ex;
                    }
                }
#pragma warning restore CA1031
                finally
                {
                    try
                    {
                        Marshal.ReleaseComObject(wb);
                    }
                    catch (COMException ex)
                    {
                        if (cleanupException is null)
                        {
                            cleanupException = ex;
                        }
                    }
#pragma warning disable CA1031
                    catch (Exception ex)
                    {
                        if (cleanupException is null)
                        {
                            cleanupException = ex;
                        }
                    }
#pragma warning restore CA1031
                }
            }

            _createdWorkbooks.Clear();
        }

        try
        {
            try
            {
                _workbooks = _excel.Workbooks;

                // Close every workbook still open after the tracked-workbook
                // loop. Read Count only now, then iterate backwards so removals
                // do not skip entries.
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
            // Best-effort cleanup: capture the first exception and continue to the
            // GC + orphan-poll phase.
            cleanupException ??= ex;
            _workbooks = null;
        }
        catch (ArgumentException ex)
        {
            // Workbook index out of range or similar argument issues during
            // cleanup. Best-effort: capture the first error and continue.
            cleanupException ??= ex;
            _workbooks = null;
        }
        catch (InvalidOperationException ex)
        {
            // Excel application in invalid state during cleanup.
            // Best-effort: capture the first error and continue.
            cleanupException ??= ex;
            _workbooks = null;
        }
        catch (NotImplementedException ex)
        {
            // COM method not implemented. Best-effort: capture the first error
            // and continue.
            cleanupException ??= ex;
            _workbooks = null;
        }
        catch (NotSupportedException ex)
        {
            // COM method not supported. Best-effort: capture the first error
            // and continue.
            cleanupException ??= ex;
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

    private static void PreserveFirstCleanupException(ref Exception? cleanupException, Exception exception) =>
        cleanupException ??= exception;

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

    /// <summary>
    /// Appends the fixture's owned Excel process ID to the manifest file the
    /// verifying shell reads on timeout. The path is supplied by the shell via
    /// <c>$env:GANTTCREATOR_OWNED_PIDS_PATH</c>; when it is unset (for example
    /// when the fixture is exercised directly rather than through
    /// <c>verify-office.ps1</c>) this is a no-op.
    /// </summary>
    /// <param name="processId">The owned Excel process ID, or zero when it
    /// could not be identified.</param>
    private static void RecordOwnedProcessId(int processId)
    {
        if (processId == 0) return;

        var path = Environment.GetEnvironmentVariable("GANTTCREATOR_OWNED_PIDS_PATH");
        if (string.IsNullOrEmpty(path)) return;

        try
        {
            var entry = new OwnedProcessIdEntry { ProcessId = processId };
            var json = JsonSerializer.Serialize(entry);
            // Append, not overwrite: a run can launch more than one fixture
            // before a timeout, and each owned PID must survive.
            File.AppendAllText(path, json + Environment.NewLine);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
        {
            // Best-effort: the manifest is a safety net for the timeout sweep,
            // never a test gate. Swallow and move on.
            _ = ex; // silence unused-variable warning under warnings-as-errors
        }
    }

    /// <summary>
    /// One owned-PID record as written to the manifest file.
    /// </summary>
    private sealed class OwnedProcessIdEntry
    {
        public int ProcessId { get; set; }
    }

    /// <summary>
    /// Test-only entry point for <see cref="RecordOwnedProcessId"/>, which is
    /// private. Kept internal (not public) so the production surface stays
    /// narrow; the xUnit test in the same assembly calls this.
    /// </summary>
    internal static void RecordOwnedProcessIdForTest(int processId) =>
        RecordOwnedProcessId(processId);
}
