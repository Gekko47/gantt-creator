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
/// <remarks>
/// Not sealed: <see cref="KillOwnedProcess"/> is an <c>internal virtual</c> seam
/// so a contract test can drive the teardown-escalation path without spawning
/// Excel.
/// </remarks>
internal class OfficeFixture : IAsyncLifetime
{
    private Application? _excel;
    private Workbooks? _workbooks;
    private int _excelProcessId;
    private bool _disposed;

    /// <summary>
    /// Every EXCEL.EXE process this launch created, in discovery order. A single
    /// <c>new Application()</c> can start more than one process, so the first is
    /// the primary tracked PID and the rest are recorded too; otherwise those
    /// extra processes are neither verified at teardown nor handed to the
    /// verifying shell's sweep.
    /// </summary>
    private readonly List<int> _ownedProcessIds = [];

    /// <summary>
    /// How long the owned process gets to exit on its own after <c>Quit()</c> and
    /// a GC pass. The production bound is 10 s; a process that needs the full
    /// window is the unreleased-COM-proxy signal, not normal Excel lag. Virtual
    /// so the escalation test can shorten it instead of spending 10 s.
    /// </summary>
    internal virtual TimeSpan GracefulExitTimeout => TimeSpan.FromSeconds(10);

    /// <summary>
    /// How long the owned process gets to disappear after the forced kill. A kill
    /// is synchronous for the target process, so this only covers the OS
    /// reclaiming the handle.
    /// </summary>
    internal virtual TimeSpan ForcedExitTimeout => TimeSpan.FromSeconds(5);
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
    /// Every EXCEL.EXE process ID this fixture launched, primary first. Empty when
    /// no new process could be identified, in which case the orphan check and the
    /// shell sweep have nothing to act on.
    /// </summary>
    public IReadOnlyList<int> OwnedProcessIds => _ownedProcessIds;

    /// <summary>
    /// Seeds the owned-PID list for a test that drives teardown without launching
    /// Excel. Internal so the production launch path stays the only writer.
    /// </summary>
    internal List<int> OwnedProcessIdsForTest
    {
        get => _ownedProcessIds;
        set
        {
            _ownedProcessIds.Clear();
            _ownedProcessIds.AddRange(value);
            if (_ownedProcessIds.Count > 0)
            {
                _excelProcessId = _ownedProcessIds[0];
            }
        }
    }

    /// <summary>
    /// Number of times teardown had to escalate to a forced kill because the
    /// owned process survived the graceful quit, GC, and exit poll.
    /// </summary>
    /// <remarks>
    /// A non-zero count is the regression signal for unreleased COM proxies: the
    /// escalation stops the leak from breaking the gate, and this counter is what
    /// makes it visible. The goal is zero. A test that forces an escalation on
    /// purpose (to prove the escalation path works) sets
    /// <see cref="SuppressLeakSignal"/> so its deliberate kill is not counted
    /// against the ratchet.
    /// </remarks>
    public int ForcedKillCount { get; private set; }

    /// <summary>
    /// Whether this fixture's escalation count should be excluded from the leak
    /// signal. Set only by a test that forces a kill deliberately, never by a
    /// test that simply leaked.
    /// </summary>
    internal bool SuppressLeakSignal { get; set; }

    /// <summary>
    /// Records the escalation count to the shell so the gate can report the leak
    /// signal. The target is zero; a non-zero value means a test body left a COM
    /// proxy alive and the process only went away because it was killed.
    /// </summary>
    /// <remarks>
    /// Best-effort, like the PID handoff: it is diagnostics, never a pass/fail
    /// input, and a failure to write must not fail the test.
    /// </remarks>
    internal void ReportForcedKillCount()
    {
        if (SuppressLeakSignal)
        {
            return;
        }

        var path = Environment.GetEnvironmentVariable("GANTTCREATOR_FORCED_KILLS_PATH");
        if (string.IsNullOrEmpty(path))
        {
            return;
        }

        try
        {
            File.AppendAllText(path, $"{ForcedKillCount}{Environment.NewLine}");
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            _ = ex;
        }
    }

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

        // Identify the EXCEL.EXE processes this launch created by diffing the
        // snapshot. Record every one of them: the primary drives the orphan poll
        // and the escalation, and the rest are still handed to the shell so a
        // stray among them is swept rather than silently left behind.
        var after = Process.GetProcessesByName("EXCEL")
            .Select(p => p.Id)
            .Where(id => !before.Contains(id))
            .ToList();

        foreach (var id in after)
        {
            _ownedProcessIds.Add(id);
        }

        if (after.Count > 0)
        {
            _excelProcessId = after[0];
        }

        // Report the owned PIDs to the verifying shell. On a genuine timeout
        // DisposeAsync never runs and the test host is killed, so the shell
        // needs a signal that outlives the process; the manifest file does.
        // Best-effort: this must never affect test pass/fail.
        foreach (var id in _ownedProcessIds)
        {
            RecordOwnedProcessId(id);
        }

        return Task.CompletedTask;
    }

    /// <inheritdoc />
    public async Task DisposeAsync()
    {
        if (_disposed) return;
        _disposed = true;

        // A fixture that launched no Excel still owns the PIDs a test seeded
        // through OwnedProcessIdsForTest, so the owned-process exit check below
        // must still run for them. Only the COM teardown is skipped.
        if (_excel == null)
        {
            await EnsureOwnedProcessesExitedAsync().ConfigureAwait(true);
            ReportForcedKillCount();
            return;
        }

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

        // Every owned process this launch created gets the same treatment, not
        // just the primary: a second Excel process from one launch is equally
        // capable of holding the packed XLL open.
        await EnsureOwnedProcessesExitedAsync().ConfigureAwait(true);
        ReportForcedKillCount();

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
    /// A scoped COM reference-count release for integration test bodies.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Every test in this assembly acquires a chain of proxies --
    /// <c>workbook.Sheets</c>, <c>worksheet.ListObjects</c>,
    /// <c>listObject.ListRows</c>, <c>listRow.Range</c> -- and none of them were
    /// ever released. Unreleased proxies keep the Application's reference count
    /// above zero, so <c>Quit()</c> returns without terminating the process: that
    /// is the root cause of the leaked EXCEL.EXE. The fixture now force-kills a
    /// survivor, so the leak no longer breaks the gate, but the kill is a
    /// symptom, not a fix.
    /// </para>
    /// <para>
    /// A scope releases its proxies in reverse acquisition order on dispose, which
    /// matches the COM ownership rule the production adapters already follow. Use
    /// it as:
    /// <code>
    /// using var scope = new OfficeFixture.ComScope();
    /// var sheets = scope.Track(workbook.Sheets);
    /// </code>
    /// so a converted test stops contributing to the leak signal.
    /// </para>
    /// </remarks>
    internal sealed class ComScope : IDisposable
    {
        private readonly List<object> _tracked = [];

        /// <summary>Tracks a proxy for release when this scope is disposed.</summary>
        /// <typeparam name="T">The proxy type.</typeparam>
        /// <param name="proxy">The proxy to release later.</param>
        /// <returns>The same proxy, so the call can wrap an expression.</returns>
        /// <remarks>
        /// The constraint is <c>class</c>, not <see cref="System.MarshalByRefObject"/>:
        /// the PIA's Excel types are COM <em>interfaces</em> (<c>Range</c>,
        /// <c>ListObject</c>, ...), so the concrete RCW behind them is what
        /// <see cref="System.Runtime.InteropServices.Marshal.ReleaseComObject(object)"/>
        /// accepts. Releasing a plain object that was never a COM proxy throws
        /// <see cref="ArgumentException"/>, which <see cref="Dispose"/> swallows
        /// so a non-COM argument cannot fail teardown.
        /// </remarks>
        public T Track<T>(T proxy)
            where T : class
        {
            ArgumentNullException.ThrowIfNull(proxy);
            _tracked.Add(proxy);
            return proxy;
        }

        /// <summary>Gets the number of proxies currently tracked.</summary>
        public int TrackedCount => _tracked.Count;

        /// <inheritdoc />
        public void Dispose()
        {
            for (var index = _tracked.Count - 1; index >= 0; index--)
            {
                try
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_tracked[index]);
                }
                catch (ArgumentException)
                {
                    // Already released elsewhere in the test; releasing twice is
                    // not an error worth failing teardown over.
                }
            }

            _tracked.Clear();
        }
    }

    /// <summary>Reports whether a process is still running. Virtual so the teardown
    /// escalation can be driven deterministically without spawning a process that
    /// has to be kept alive for the duration of the test.
    /// </summary>
    /// <param name="processId">The process to probe; zero counts as not running.</param>
    /// <returns><see langword="true"/> when the process is alive.</returns>
    internal virtual bool IsProcessRunning(int processId)
    {
        if (processId == 0) return false;

        try
        {
            using var process = Process.GetProcessById(processId);
            return !process.HasExited;
        }
        catch (ArgumentException)
        {
            return false;
        }
    }

    /// <summary>
    /// Polls an owned Excel process until it exits or the deadline elapses.
    /// Never throws.
    /// </summary>
    /// <param name="processId">The process to poll; zero is treated as already gone.</param>
    /// <param name="timeout">How long to wait for exit.</param>
    /// <returns><see langword="true"/> when the process is gone (or was never identified).</returns>
    private async Task<bool> PollForProcessExitAsync(int processId, TimeSpan timeout)
    {
        if (processId == 0) return true;

        var sw = Stopwatch.StartNew();
        while (sw.Elapsed < timeout)
        {
            if (!IsProcessRunning(processId))
            {
                return true;
            }

            await Task.Delay(100).ConfigureAwait(true);
        }

        return !IsProcessRunning(processId);
    }

    /// <summary>
    /// Ensures every process this fixture launched is gone, escalating to a
    /// targeted kill, and reports any that survived as a cleanup failure.
    /// </summary>
    /// <remarks>
    /// A survivor is not tolerated: it holds the packed XLL and breaks the next
    /// run's publish step, so it must surface as a failure rather than a warning.
    /// </remarks>
    /// <exception cref="InvalidOperationException">Thrown when a process outlives both the graceful quit and the kill.</exception>
    private async Task EnsureOwnedProcessesExitedAsync()
    {
        var surviving = new List<int>();
        foreach (var ownedId in _ownedProcessIds)
        {
            if (!await EnsureOwnedProcessExitedAsync(ownedId).ConfigureAwait(true))
            {
                surviving.Add(ownedId);
            }
        }

        if (surviving.Count > 0)
        {
            throw new InvalidOperationException(
                $"Office Excel process(es) {string.Join(", ", surviving)} survived teardown.");
        }
    }

    /// <summary>
    /// Waits for the owned process to exit after the graceful quit, escalating to
    /// a forced kill of this fixture's own process when it does not.
    /// </summary>
    /// <remarks>
    /// The escalation exists because unreleased COM proxies in a test body keep
    /// the Application's reference count above zero, so <c>Quit()</c> returns
    /// without terminating the process. Left alone that process holds the packed
    /// XLL open and breaks the next run's publish step. The kill is targeted by
    /// the PID this fixture launched, never by process name, and it is counted in
    /// <see cref="ForcedKillCount"/> so a persistent leak stays visible instead of
    /// being masked by the force.
    /// </remarks>
    /// <param name="processId">The owned process ID.</param>
    /// <returns><see langword="true"/> when the process is gone by the end.</returns>
    private async Task<bool> EnsureOwnedProcessExitedAsync(int processId)
    {
        if (await PollForProcessExitAsync(processId, GracefulExitTimeout).ConfigureAwait(true))
        {
            return true;
        }

        if (KillOwnedProcess(processId))
        {
            ForcedKillCount++;
        }

        return await PollForProcessExitAsync(processId, ForcedExitTimeout).ConfigureAwait(true);
    }

    /// <summary>
    /// Forces the fixture's own Excel process to terminate.
    /// </summary>
    /// <remarks>
    /// Virtual so a contract test can substitute the kill and drive the
    /// escalation path without spawning Excel, mirroring the COM indexer seams
    /// elsewhere in this harness. The default kills by PID only; it never matches
    /// on the process name, so a user-owned Excel can never be terminated here.
    /// </remarks>
    /// <param name="processId">The owned process ID.</param>
    /// <returns>Whether the kill was issued successfully.</returns>
    internal virtual bool KillOwnedProcess(int processId)
    {
        if (processId == 0) return false;

        try
        {
            using var process = Process.GetProcessById(processId);
            if (process.HasExited) return false;

            process.Kill(entireProcessTree: true);
            return true;
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or NotSupportedException or System.ComponentModel.Win32Exception)
        {
            return false;
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
            var entry = new OwnedProcessIdEntry
            {
                ProcessId = processId,
                StartTimeUtcTicks = ReadStartTimeUtcTicks(processId),
            };
            var json = JsonSerializer.Serialize(entry);
            // Append, not overwrite: a run can launch more than one fixture
            // before a timeout, and each owned PID must survive. One complete
            // JSON record per line, so a reader can parse each line on its own
            // (a whole-file ConvertFrom-Json fails on concatenated objects).
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
    /// Reads a live process start time in UTC ticks for the manifest, or zero
    /// when it cannot be read.
    /// </summary>
    /// <remarks>
    /// The start time is the second half of the sweep's identity proof: a PID
    /// alone can be recycled by Windows between runs, so the shell only kills a
    /// manifest PID whose live process still reports the recorded start time. A
    /// zero here is deliberately recorded as unverifiable, and the shell's
    /// fail-safe direction is to skip such a PID rather than kill it.
    /// </remarks>
    /// <param name="processId">The process to sample.</param>
    /// <returns>The UTC start time in ticks, or zero when unreadable.</returns>
    internal static long ReadStartTimeUtcTicks(int processId)
    {
        try
        {
            using var process = Process.GetProcessById(processId);
            return process.StartTime.ToUniversalTime().Ticks;
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or NotSupportedException)
        {
            return 0;
        }
    }

    /// <summary>
    /// One owned-PID record as written to the manifest file: the process ID and
    /// the start time that identifies that specific process, so a recycled PID
    /// is never mistaken for the harness's own.
    /// </summary>
    private sealed class OwnedProcessIdEntry
    {
        public int ProcessId { get; set; }

        public long StartTimeUtcTicks { get; set; }
    }

    /// <summary>
    /// Test-only entry point for <see cref="RecordOwnedProcessId"/>, which is
    /// private. Kept internal (not public) so the production surface stays
    /// narrow; the xUnit test in the same assembly calls this.
    /// </summary>
    internal static void RecordOwnedProcessIdForTest(int processId) =>
        RecordOwnedProcessId(processId);
}
