using System.Globalization;
using ExcelDna.Integration;
using GanttCreator.Core;
using GanttCreator.Core.Logging;
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// The Excel-DNA entry point. <c>AutoOpen</c> writes one open record and
/// <c>AutoClose</c> one close record to the local rolling log; neither ever
/// propagates an exception into Excel, because an unhandled entry-point
/// exception aborts the add-in load with a host error dialog. Every
/// failure degrades to "no log record", which is the safe direction for a
/// logging-only operation.
/// </summary>
/// <param name="identitySource">Supplies the load-session identity.</param>
/// <param name="logSource">Supplies the rolling log.</param>
/// <param name="applicationAdapterSource">
/// Supplies the Excel application adapter during <see cref="AutoOpen"/>.
/// Defaults to the production <c>ExcelDnaUtil.Application</c> probe;
/// injectable so the <see cref="AutoClose"/> teardown order is
/// contract-testable outside Excel (work item R1.6).
/// </param>
/// <exception cref="ArgumentNullException">A source is <see langword="null"/>.</exception>
public sealed class AddInHost(
    Func<AddInIdentity> identitySource,
    Func<IRollingLog> logSource,
    Func<IExcelApplicationAdapter>? applicationAdapterSource = null) : IExcelAddIn
{
    private readonly Func<AddInIdentity> _identitySource =
        identitySource ?? throw new ArgumentNullException(nameof(identitySource));

    private readonly Func<IRollingLog> _logSource =
        logSource ?? throw new ArgumentNullException(nameof(logSource));

    // Defaults to the production ExcelDnaUtil.Application probe; injectable
    // so the AutoClose teardown order is contract-testable outside Excel
    // (work item R1.6).
    private readonly Func<IExcelApplicationAdapter> _applicationAdapterSource =
        applicationAdapterSource ?? DefaultApplicationAdapterSource;

    private IRollingLog? _log;
    private string? _sessionToken;

    /// <summary>
    /// Creates the add-in host with the production identity and log
    /// sources. Excel-DNA requires this public parameterless constructor.
    /// </summary>
    public AddInHost()
        : this(DefaultIdentitySource, AddInLogFactory.Create)
    {
    }

    /// <summary>
    /// Generates one per-load session token (work item R1.6 D5). The token
    /// correlates the <c>open</c>/<c>close</c> pair of a single
    /// <c>AutoOpen</c>/<c>AutoClose</c> cycle in the shared log. It is
    /// deliberately not a GUID or long hex string: Core's
    /// <see cref="Redactor"/> masks both to <c>[guid]</c>/<c>[token]</c> on
    /// disk, which would destroy correlation. The alphabet below excludes
    /// <c>a</c>–<c>f</c> so a 12-character token can never match the long-hex
    /// pattern; the timestamp prefix keeps separately generated tokens
    /// distinct.
    /// </summary>
    /// <returns>A redaction-safe correlation token.</returns>
    internal static string GenerateSessionToken()
    {
        Span<char> token = stackalloc char[12];
        // Draw one nibble at a time and map 10-15 to g-p (never a-f), so the
        // 12-character suffix can never match the Redactor's long-hex
        // pattern; the timestamp prefix keeps separately generated tokens
        // distinct even within one second boundary.
        var filled = 0;
        while (filled < token.Length)
        {
            var value = System.Security.Cryptography.RandomNumberGenerator.GetInt32(0, 16);
            token[filled++] = value < 10 ? (char)('0' + value) : (char)('g' + (value - 10));
        }
        return string.Create(
            CultureInfo.InvariantCulture,
            $"s{DateTimeOffset.UtcNow.ToUnixTimeSeconds():x}-{new string(token)}");
    }

    /// <summary>
    /// Excel-DNA entry point invoked when the XLL loads. Writes exactly one
    /// open record; failures degrade to no logging instead of propagating
    /// into Excel.
    /// </summary>
    public void AutoOpen()
    {
        // CA1031: The Excel-DNA entry point must never propagate an exception
        // into Excel — an unhandled AutoOpen exception aborts the add-in load
        // with a host error dialog. Every failure below (log creation,
        // identity probe, write) degrades to "no log record", which is the
        // safe direction for a logging-only operation. Core's RollingLog
        // already contains its own write/rotation failures; this guard
        // covers the AddIn-owned sources around it.
        try
        {
            IRollingLog log = _logSource();
            try
            {
                AddInIdentity identity = _identitySource();
                new AddInLifecycle(log).LogOpen(identity);
                _sessionToken = identity.SessionToken;
            }
            catch
            {
                log.Dispose();
                throw;
            }
            _log = log;

            // Wire the log into the diagnostics service and the command
            // boundary so the ribbon's Diagnostics button can display the
            // active log path and the boundary can write its error records.
            // SetLog only stores the reference (its only throw is a
            // null-argument guard, unreachable for the non-null local above);
            // any other unforeseen failure is handled by the catch below,
            // so AutoOpen stays never-throwing.
            DiagnosticsService.Instance.SetLog(log);
            CommandBoundary.Instance.SetLog(log);
        }
#pragma warning disable CA1031
        catch
#pragma warning restore CA1031
        {
            // Intentionally empty: no log record, no propagation. See the
            // justification comment above. The candidate log is disposed on
            // the failure path so AutoClose cannot emit an unmatched close
            // record; _log is only assigned after LogOpen succeeds.
        }

        // Arm the Ribbon state service in a separate guarded step so that
        // this runs even when log creation, LogOpen, or SetLog failed above.
        // ExcelDnaUtil.Application returns null outside a live Excel host, so
        // the adapter reports "not determinable" (which keeps the previous
        // value) instead of throwing; the state service never writes a log
        // record, so the one-record contract above holds. Any failure from
        // ExcelDnaUtil.Application or RibbonStateService is swallowed here so
        // it never escapes AutoOpen.
        try
        {
            RibbonStateService.Instance.SetApplicationAdapter(
                _applicationAdapterSource());
            RibbonStateService.Instance.SetLogAvailabilitySource(
                () => !string.IsNullOrWhiteSpace(DiagnosticsService.Instance.LogFilePath));
            RibbonStateService.Instance.Activate();
        }
#pragma warning disable CA1031
        catch
#pragma warning restore CA1031
        {
            // Intentionally empty: state-service arming degrades silently.
        }
    }

    /// <summary>
    /// Excel-DNA entry point invoked when the XLL unloads. Tears down in one
    /// deterministic order (work item R1.6): the owned COM event connection
    /// is detached first, then one close record is written, then every
    /// session singleton drops its log reference, and the log is disposed
    /// last. Failures degrade instead of propagating into Excel.
    /// </summary>
    public void AutoClose()
    {
        // CA1031: Teardown must never propagate into Excel — an exception
        // from AutoClose surfaces as a host error during unload. Every step
        // is individually guarded so one failure cannot block the remaining
        // steps; a failure degrades to "that step not done".
        try
        {
            // Step 1 (work item R1.6): detach the one owned COM resource —
            // the workbook-state event connection — before anything else, so
            // a workbook-state event can no longer fire into services
            // mid-teardown. Reset never throws: every detach failure is
            // guarded inside the subscription itself.
            RibbonStateService.Reset();

            // Step 2: exactly one close record carrying this load's session
            // token. A write failure degrades to a missing close record.
            try
            {
                if (_log is not null)
                {
                    new AddInLifecycle(_log).LogClose(_sessionToken);
                }
            }
#pragma warning disable CA1031
            catch
#pragma warning restore CA1031
            {
                // Intentionally empty: teardown degrades silently. See the
                // justification comment above.
            }

            // Step 3: drop every session singleton's log reference BEFORE the
            // log is disposed, so no service can hold (or write through) a
            // disposed log. Both resets only clear references under a lock.
            CommandBoundary.Reset();
            DiagnosticsService.Reset();
        }
#pragma warning disable CA1031
        catch
#pragma warning restore CA1031
        {
            // Intentionally empty: teardown must never propagate into Excel.
        }
        finally
        {
            // Step 4: the log is disposed last, when no other component holds
            // it. Disposal exceptions are suppressed so AutoClose never
            // propagates into Excel.
            try
            {
                _log?.Dispose();
            }
#pragma warning disable CA1031
            catch
#pragma warning restore CA1031
            {
                // Intentionally empty: see the disposal rationale above.
            }
            finally
            {
                _log = null;
                _sessionToken = null;
            }
        }
    }

    /// <summary>
    /// The production application-adapter source: the Excel application
    /// object supplied by the Excel-DNA host. Outside a live Excel host the
    /// probe returns null, so the adapter reports "not determinable" and
    /// subscribes no events.
    /// </summary>
    /// <returns>The live application adapter for this Excel session.</returns>
    private static IExcelApplicationAdapter DefaultApplicationAdapterSource() =>
        new ExcelApplicationAdapter(ExcelDnaUtil.Application);

    private static AddInIdentity DefaultIdentitySource()
    {
        // CA1031: Each identifier is probed while Excel may be in an
        // early-load or teardown state (ExcelDnaUtil reads host state).
        // A probe failure yields "unknown" for that field so the open
        // record still reports the remaining identifiers; it must never
        // propagate into AutoOpen.
        var excelVersion = "unknown";
        var xllFileName = "unknown";
#pragma warning disable CA1031
        try
        {
            excelVersion = ExcelDnaUtil.ExcelVersion.ToString("0.0", CultureInfo.InvariantCulture);
        }
        catch
        {
            // Intentionally empty: identifier degrades to "unknown".
        }

        try
        {
            xllFileName = Path.GetFileName(ExcelDnaUtil.XllPath);
        }
        catch
        {
            // Intentionally empty: identifier degrades to "unknown".
        }
#pragma warning restore CA1031

        return new AddInIdentity(
            VersionInfo.InformationalVersion,
            excelVersion,
            Environment.Is64BitProcess ? "x64" : "x86",
            string.IsNullOrWhiteSpace(xllFileName) ? "unknown" : xllFileName,
            GenerateSessionToken());
    }
}
