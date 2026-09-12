using System.Globalization;
using ExcelDna.Integration;
using GanttCreator.Core;
using GanttCreator.Core.Logging;

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
/// <exception cref="ArgumentNullException">A source is <see langword="null"/>.</exception>
public sealed class AddInHost(
    Func<AddInIdentity> identitySource,
    Func<IRollingLog> logSource) : IExcelAddIn
{
    private readonly Func<AddInIdentity> _identitySource =
        identitySource ?? throw new ArgumentNullException(nameof(identitySource));

    private readonly Func<IRollingLog> _logSource =
        logSource ?? throw new ArgumentNullException(nameof(logSource));

    private IRollingLog? _log;

    /// <summary>
    /// Creates the add-in host with the production identity and log
    /// sources. Excel-DNA requires this public parameterless constructor.
    /// </summary>
    public AddInHost()
        : this(DefaultIdentitySource, AddInLogFactory.Create)
    {
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
                new AddInLifecycle(log).LogOpen(_identitySource());
            }
            catch
            {
                log.Dispose();
                throw;
            }
            _log = log;
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
    }

    /// <summary>
    /// Excel-DNA entry point invoked when the XLL unloads. Writes exactly
    /// one close record and disposes the log; failures degrade instead of
    /// propagating into Excel.
    /// </summary>
    public void AutoClose()
    {
        // CA1031: Teardown must never propagate into Excel — an exception
        // from AutoClose surfaces as a host error during unload. The write
        // failure degrades to a missing close record.
        try
        {
            if (_log is not null)
            {
                new AddInLifecycle(_log).LogClose();
            }
        }
#pragma warning disable CA1031
        catch
#pragma warning restore CA1031
        {
            // Intentionally empty: teardown degrades silently. See the
            // justification comment above.
        }
        finally
        {
            try
            {
                _log?.Dispose();
            }
#pragma warning disable CA1031
            catch
#pragma warning restore CA1031
            {
                // Suppress disposal exceptions so AutoClose never
                // propagates into Excel.
            }
            finally
            {
                _log = null;
            }
        }
    }

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
            string.IsNullOrWhiteSpace(xllFileName) ? "unknown" : xllFileName);
    }
}
