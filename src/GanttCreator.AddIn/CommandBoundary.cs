using System.Globalization;
using GanttCreator.Core.Logging;

namespace GanttCreator.AddIn;

/// <summary>
/// The single command error boundary for Ribbon callbacks and future
/// application commands. A command that completes normally produces no log
/// record and no dialog. A command that throws produces exactly one
/// rolling-log record (operation ID, command name, stage, Excel version,
/// exception type, HRESULT, redacted detail) and exactly one user-safe
/// dialog containing the operation ID. The boundary never throws.
/// </summary>
/// <remarks>
/// <para>
/// docs/02-ARCHITECTURE.md "Error handling": unexpected exceptions cross one
/// command boundary where they are logged and translated to a concise message
/// containing an operation ID; message boxes are never scattered through lower
/// layers. The boundary is a singleton per Excel session
/// (<see cref="Instance"/>, <see cref="SetLog"/>, <see cref="Reset"/>),
/// mirroring the <see cref="DiagnosticsService"/> pattern; the log reference is
/// injected by <see cref="AddInHost"/> and the boundary does not own its
/// lifetime.
/// </para>
/// <para>
/// CA1031: every failure path in <see cref="Run(string, Action)"/> is individually guarded — a
/// log-write failure or a dialog-show failure degrades to "fewer
/// notifications" instead of propagating into Excel. A ribbon callback that
/// propagates an exception surfaces as a host error dialog, which is why the
/// boundary itself must be failure-proof.
/// </para>
/// </remarks>
public sealed class CommandBoundary
{
    private static CommandBoundary? _instance;
    private static readonly Lock _instanceGate = new();

    /// <summary>
    /// The singleton instance for the current Excel session. Creates the
    /// boundary with the system clock and the production dialog on first
    /// access; the injected log may be absent until <see cref="SetLog"/> is
    /// called.
    /// </summary>
    public static CommandBoundary Instance
    {
        get
        {
            lock (_instanceGate)
            {
                _instance ??= new CommandBoundary();
                return _instance;
            }
        }
    }

    private readonly TimeProvider _timeProvider;
    private readonly Action<string> _presenter;
    private readonly Func<string> _excelVersionSource;
    private int _sequence;
    private IRollingLog? _log;

    /// <summary>
    /// Creates the production boundary: system clock, TaskDialog presenter,
    /// live Excel-version probe. Internal so contract tests can construct a
    /// deterministic boundary (fixed clock, capturing presenter, injected
    /// log) without the singleton.
    /// </summary>
    /// <param name="timeProvider">Clock boundary for operation IDs; defaults to <see cref="TimeProvider.System"/>.</param>
    /// <param name="presenter">Receives the translated user message; defaults to the TaskDialog error dialog.</param>
    internal CommandBoundary(TimeProvider? timeProvider = null, Action<string>? presenter = null)
    {
        _timeProvider = timeProvider ?? TimeProvider.System;
        _presenter = presenter ?? CommandErrorDialog.Show;
        _excelVersionSource = ProbeExcelVersion;
    }

    /// <summary>
    /// Injects the rolling log reference. Called once per Excel session by
    /// <see cref="AddInHost"/> after the log is created. A boundary without a
    /// log still shows the user dialog; the record degrades to nothing.
    /// </summary>
    /// <param name="log">The rolling log.</param>
    public void SetLog(IRollingLog log) => _log = log ?? throw new ArgumentNullException(nameof(log));

    /// <summary>
    /// Resets the singleton and clears the log reference, so the next Excel
    /// session starts fresh. Called by <see cref="AddInHost.AutoClose"/>.
    /// </summary>
    internal static void Reset()
    {
        lock (_instanceGate)
        {
            _instance?._log = null;
            _instance = null;
        }
    }

    /// <summary>
    /// Runs one command across the boundary. On success the command runs and
    /// nothing is logged or shown. On a thrown exception exactly one log
    /// record and one user-safe dialog are produced, and the exception never
    /// escapes this method.
    /// </summary>
    /// <param name="commandName">The stable command identifier (the Ribbon control ID, e.g. <c>btnDiagnostics</c>).</param>
    /// <param name="command">The command delegate.</param>
    /// <exception cref="ArgumentNullException"><paramref name="commandName"/> or <paramref name="command"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException"><paramref name="commandName"/> is empty or whitespace.</exception>
    public void Run(string commandName, Action command)
    {
        ArgumentNullException.ThrowIfNull(commandName);
        if (string.IsNullOrWhiteSpace(commandName))
        {
            throw new ArgumentException("Command name must not be empty.", nameof(commandName));
        }

        ArgumentNullException.ThrowIfNull(command);
        Run(() => commandName, command, commandName);
    }

    /// <summary>
    /// Runs one command across the boundary with deferred command-name
    /// resolution. The resolver runs inside the boundary: a resolver failure
    /// degrades to <paramref name="fallbackCommandName"/> so the failure
    /// still produces one log record and one dialog.
    /// </summary>
    /// <param name="resolveCommandName">Resolves the stable command identifier (may probe COM state).</param>
    /// <param name="command">The command delegate.</param>
    /// <param name="fallbackCommandName">Used when the resolver throws or returns blank.</param>
    /// <exception cref="ArgumentNullException"><paramref name="resolveCommandName"/>, <paramref name="command"/>, or <paramref name="fallbackCommandName"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException"><paramref name="fallbackCommandName"/> is empty or whitespace.</exception>
    public void Run(Func<string> resolveCommandName, Action command, string fallbackCommandName)
    {
        ArgumentNullException.ThrowIfNull(resolveCommandName);
        ArgumentNullException.ThrowIfNull(command);
        ArgumentNullException.ThrowIfNull(fallbackCommandName);
        if (string.IsNullOrWhiteSpace(fallbackCommandName))
        {
            throw new ArgumentException("Command name must not be empty.", nameof(fallbackCommandName));
        }

        // CA1031: the command boundary is the one place where a generic catch
        // is the product contract — an unexpected exception crossing the
        // boundary must be logged and translated, never propagated into Excel
        // (docs/02-ARCHITECTURE.md "Error handling"). Specific categories are
        // handled inside the translator. The resolver runs inside the same
        // try so a COM control-ID probe failure is reported, not thrown.
#pragma warning disable CA1031
        var commandName = fallbackCommandName;
        try
        {
            commandName = SafeResolveCommandName(resolveCommandName, fallbackCommandName);
            // Debug-only force-failure hook (compiled out of Release). Lets the
            // required Office gate demonstrate a forced callback failure without
            // a throwaway spike. See docs/work-items/R1.4-command-error-boundary.md.
            DebugForceFailure.ThrowIfRequested(commandName);
            command();
        }
        catch (Exception ex)
        {
            ReportFailure(commandName, ex);
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Resolves the command name without throwing: a resolver failure or a
    /// blank result degrades to <paramref name="fallbackCommandName"/>.
    /// </summary>
    /// <param name="resolveCommandName">Resolves the stable command identifier.</param>
    /// <param name="fallbackCommandName">Used when the resolver throws or returns blank.</param>
    private static string SafeResolveCommandName(Func<string> resolveCommandName, string fallbackCommandName)
    {
        // CA1031: intentional degradation — the probe may touch COM state.
#pragma warning disable CA1031
        try
        {
            var resolved = resolveCommandName();
            return string.IsNullOrWhiteSpace(resolved) ? fallbackCommandName : resolved;
        }
        catch
        {
            return fallbackCommandName;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// The single failure path: exactly one log record, then exactly one
    /// user-safe dialog. Each step is individually guarded so a double
    /// failure still never propagates into Excel.
    /// </summary>
    /// <param name="commandName">The command identifier that failed.</param>
    /// <param name="ex">The caught exception.</param>
    private void ReportFailure(string commandName, Exception ex)
    {
        var operationId = NextOperationId();
        WriteErrorRecord(commandName, ex, operationId);
        ShowTranslatedDialog(ex, operationId);
    }

    /// <summary>
    /// Writes the one technical error record to the rolling log, if a log is
    /// injected. Never throws: a log-write failure degrades to no record,
    /// which is the safe direction inside an exception handler (the dialog
    /// still shows).
    /// </summary>
    /// <param name="commandName">The command identifier that failed.</param>
    /// <param name="ex">The caught exception.</param>
    /// <param name="operationId">The operation ID for this failure.</param>
    private void WriteErrorRecord(string commandName, Exception ex, string operationId)
    {
        // CA1031: this runs inside an exception handler; recording must never
        // introduce a new failure path.
#pragma warning disable CA1031
        try
        {
            var hresult = ex.HResult.ToString("X8", CultureInfo.InvariantCulture);
            _log?.Write(
                "CommandError: op={0} command={1} stage=ribbon-callback excel-version={2} exception={3} hresult=0x{4} detail={5}",
                operationId,
                commandName,
                _excelVersionSource(),
                ex.GetType().FullName,
                hresult,
                ex.ToString());
        }
        catch
        {
            // Intentionally empty: the record degrades to nothing; the dialog
            // still shows.
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Translates the exception to a user-safe message and shows exactly one
    /// dialog. Never throws: a dialog failure degrades to no dialog.
    /// </summary>
    /// <param name="ex">The caught exception.</param>
    /// <param name="operationId">The operation ID for this failure.</param>
    private void ShowTranslatedDialog(Exception ex, string operationId)
    {
        var message = CommandErrorTranslator.Translate(ex, operationId);
        // CA1031: the presenter (TaskDialog or test capture) must never make
        // the boundary throw.
#pragma warning disable CA1031
        try
        {
            _presenter(message);
        }
        catch
        {
            // Intentionally empty: a dialog failure degrades to no dialog.
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Produces the next deterministic operation ID:
    /// <c>GC-&lt;yyyyMMdd-HHmmss-fff&gt;-&lt;NN&gt;</c> in UTC with invariant
    /// culture; <c>NN</c> is a per-instance monotonic counter that makes
    /// multiple failures in the same millisecond distinguishable.
    /// </summary>
    /// <returns>The operation ID.</returns>
    private string NextOperationId()
    {
        DateTimeOffset now = _timeProvider.GetUtcNow();
        var sequence = Interlocked.Increment(ref _sequence).ToString("00", CultureInfo.InvariantCulture);
        return string.Create(
            CultureInfo.InvariantCulture,
            $"GC-{now:yyyyMMdd}-{now:HHmmss}-{now:fff}-{sequence}");
    }

    /// <summary>
    /// Probes the Excel version for the technical record. Called inside an
    /// exception handler, so any probe failure degrades to "unknown" and the
    /// record still reports the remaining fields.
    /// </summary>
    /// <returns>The Excel version in invariant culture, or "unknown".</returns>
    private static string ProbeExcelVersion()
    {
        // CA1031: ExcelDnaUtil reads host state that may be in an early-load
        // or teardown state during a failure path; a probe failure must not
        // propagate out of the record writer.
#pragma warning disable CA1031
        try
        {
            return ExcelDna.Integration.ExcelDnaUtil.ExcelVersion.ToString("0.0", CultureInfo.InvariantCulture);
        }
        catch
        {
            return "unknown";
        }
#pragma warning restore CA1031
    }
}
