using System.Globalization;
using ExcelDna.Integration.CustomUI;
using GanttCreator.Core;
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// The Ribbon state service: the single owner of the dynamic Ribbon getters'
/// input and of the invalidate mechanism.
/// </summary>
/// <remarks>
/// <para>
/// docs/02-ARCHITECTURE.md "Ribbon and commands": every callback is a thin error
/// boundary and dynamic Ribbon state getters must be fast and side-effect-free.
/// <see cref="GetEnabled"/> therefore reads one cached <see cref="RibbonState"/>
/// — it never probes Excel, never touches the ribbon handle, and never writes a
/// log record. The snapshot changes only inside <see cref="Refresh"/>, which is
/// driven by <c>onLoad</c>, by workbook-state events from the Office adapter,
/// and by each command the ribbon runs.
/// </para>
/// <para>
/// The invalidate mechanism is the handle Excel passes to the RibbonX
/// <c>onLoad</c> callback: Excel-DNA's own guidance is to keep that object and
/// call <c>Invalidate</c>/<c>InvalidateControl</c> later to drive dynamic
/// <c>getEnabled</c>/<c>getVisible</c> updates. The service is a singleton per
/// Excel session (<see cref="Instance"/>, <see cref="Reset"/>), mirroring the
/// <see cref="DiagnosticsService"/> and <see cref="CommandBoundary"/> pattern;
/// the ribbon handle and the application adapter are injected by
/// <see cref="GanttRibbon.OnLoad(IRibbonUI)"/> and <see cref="AddInHost"/> and are not
/// owned by the service.
/// </para>
/// <para>
/// CA1031: this type is reached from Ribbon callbacks and from Excel
/// application events, where a propagating exception surfaces as a host error
/// dialog. Every failure path is individually guarded and degrades to a safe
/// direction: the previous snapshot, a fail-open getter value, or no
/// invalidation.
/// </para>
/// </remarks>
internal class RibbonStateService
{
    private static RibbonStateService? _instance;
    private static readonly Lock _instanceGate = new();

    private IRibbonUI? _ribbon;
    private IExcelApplicationAdapter? _applicationAdapter;
    private Func<bool?>? _logAvailabilitySource;
    private IDisposable? _workbookStateSubscription;
    private IConfigCatalogueWriter? _catalogueWriter;
    private RibbonState _state = RibbonState.Initial;

    internal RibbonStateService()
    {
    }

    /// <summary>
    /// Injects the catalogue writer that persists plot-range settings to the
    /// workbook configuration sheet. A null writer leaves the persist step
    /// as a no-op, so the in-memory state is still correct and the next
    /// refresh re-reads the worksheet.
    /// </summary>
    /// <param name="writer">The catalogue writer, or null when unavailable.</param>
    internal void SetCatalogueWriter(IConfigCatalogueWriter? writer) => _catalogueWriter = writer;

    /// <summary>
    /// Gets the singleton instance for the current Excel session. Creates the
    /// service on first access; the ribbon handle, the application adapter, and
    /// the log source may all be absent until the host injects them.
    /// </summary>
    internal static RibbonStateService Instance
    {
        get
        {
            lock (_instanceGate)
            {
                _instance ??= new RibbonStateService();
                return _instance;
            }
        }
    }

    /// <summary>
    /// Publishes the RibbonUI handle Excel supplied to the <c>onLoad</c>
    /// callback, so later refreshes can invalidate the Ribbon. A null handle
    /// clears the reference; this never throws, because a throwing
    /// <c>onLoad</c> breaks the Ribbon.
    /// </summary>
    /// <param name="ribbon">The RibbonUI handle, or null when Excel supplied none.</param>
    internal void SetRibbon(IRibbonUI? ribbon) => _ribbon = ribbon;

    /// <summary>
    /// Returns the RibbonUI handle published by the most recent
    /// <c>onLoad</c> callback, for test verification only.
    /// </summary>
    internal IRibbonUI? GetRibbon() => _ribbon;

    /// <summary>
    /// Injects the Excel application adapter that supplies the workbook fact and
    /// workbook-state events. A null adapter leaves the fact at its previous
    /// value and subscribes no events.
    /// </summary>
    /// <param name="adapter">The application adapter, or null when unavailable.</param>
    internal void SetApplicationAdapter(IExcelApplicationAdapter? adapter) => _applicationAdapter = adapter;

    /// <summary>
    /// Injects the log-availability source that supplies the second state fact.
    /// A null source leaves the fact at its previous value; the delegate is
    /// invoked inside <see cref="Refresh"/> only, never by a getter.
    /// </summary>
    /// <param name="source">
    /// Returns true when the log exposes an active file path, or null when the
    /// fact is not determinable (which keeps the previous value).
    /// </param>
    internal void SetLogAvailabilitySource(Func<bool?>? source) => _logAvailabilitySource = source;

    /// <summary>
    /// Toggles the plot-start end between automatic (derived from the data)
    /// and explicit (the stored date). Checking AUTO clears the stored start
    /// date and persists DataRange immediately; unchecking only arms the edit
    /// box, because the explicit mode is persisted by
    /// <see cref="SetPlotStartDate"/> together with the committed date and an
    /// Explicit mode with no date would break the next Refresh. Never
    /// refreshes the chart.
    /// </summary>
    internal void TogglePlotStartAuto()
    {
        var nextAuto = !_state.PlotStartAuto;
        _state = _state with
        {
            PlotStartAuto = nextAuto,

            // Entering AUTO clears the date so the disabled edit box and the
            // next getText cannot show a stale explicit value; entering
            // explicit keeps whatever is stored (AUTO already cleared it).
            PlotStartDate = nextAuto ? string.Empty : _state.PlotStartDate,
        };

        if (!nextAuto)
        {
            // Nothing coherent to persist yet: the date commit owns the
            // Explicit write.
            return;
        }

#pragma warning disable CA1031
        try
        {
            PersistPlotSettings();
        }
        catch (Exception)
#pragma warning restore CA1031
        {
            // Degrade gracefully: the in-memory state is still correct,
            // and the next refresh re-reads the worksheet.
        }
    }

    /// <summary>
    /// Toggles the plot-finish end between automatic (derived from the data)
    /// and explicit (the stored date). Checking AUTO clears the stored finish
    /// date and persists DataRange immediately; unchecking only arms the edit
    /// box, because the explicit mode is persisted by
    /// <see cref="SetPlotFinishDate"/> together with the committed date and
    /// an Explicit mode with no date would break the next Refresh. Never
    /// refreshes the chart.
    /// </summary>
    internal void TogglePlotFinishAuto()
    {
        var nextAuto = !_state.PlotFinishAuto;
        _state = _state with
        {
            PlotFinishAuto = nextAuto,
            PlotFinishDate = nextAuto ? string.Empty : _state.PlotFinishDate,
        };

        if (!nextAuto)
        {
            // Nothing coherent to persist yet: the date commit owns the
            // Explicit write.
            return;
        }

#pragma warning disable CA1031
        try
        {
            PersistPlotSettings();
        }
        catch (Exception)
#pragma warning restore CA1031
        {
            // Degrade gracefully: the in-memory state is still correct,
            // and the next refresh re-reads the worksheet.
        }
    }

    /// <summary>
    /// Commits a user-typed plot-start date in explicit mode. Parses any
    /// reasonable date format, normalises it to dd/MM/yyyy for storage, and
    /// pivots the start end to Explicit; an unparsable value is a no-op (the
    /// edit box keeps showing the previously stored date on the next getter
    /// query, which is the revert). Persists to the workbook configuration
    /// sheet so it survives a close/reopen. Never refreshes the chart: the
    /// next Refresh chart applies the stored settings.
    /// </summary>
    /// <param name="dateText">The user-entered explicit plot start date text.</param>
    internal void SetPlotStartDate(string dateText)
    {
        if (!PlotRangeResolver.TryParsePlotDate(dateText, out DateOnly parsed))
        {
            return;
        }

        var normalised = parsed.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
        _state = _state with
        {
            PlotStartAuto = false,
            PlotStartDate = normalised,
        };

#pragma warning disable CA1031
        try
        {
            PersistPlotSettings();
        }
        catch (Exception)
#pragma warning restore CA1031
        {
            // Degrade gracefully: the in-memory state is still correct,
            // and the next refresh re-reads the worksheet.
        }
    }

    /// <summary>
    /// Commits a user-typed plot-finish date in explicit mode. Parses any
    /// reasonable date format, normalises it to dd/MM/yyyy for storage, and
    /// pivots the finish end to Explicit; an unparsable value is a no-op (the
    /// edit box keeps showing the previously stored date on the next getter
    /// query, which is the revert). Persists to the workbook configuration
    /// sheet so it survives a close/reopen. Never refreshes the chart: the
    /// next Refresh chart applies the stored settings.
    /// </summary>
    /// <param name="dateText">The user-entered explicit plot finish date text.</param>
    internal void SetPlotFinishDate(string dateText)
    {
        if (!PlotRangeResolver.TryParsePlotDate(dateText, out DateOnly parsed))
        {
            return;
        }

        var normalised = parsed.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
        _state = _state with
        {
            PlotFinishAuto = false,
            PlotFinishDate = normalised,
        };

#pragma warning disable CA1031
        try
        {
            PersistPlotSettings();
        }
        catch (Exception)
#pragma warning restore CA1031
        {
            // Degrade gracefully: the in-memory state is still correct,
            // and the next refresh re-reads the worksheet.
        }
    }

    /// <summary>
    /// Gets the stored explicit plot-start date for the ribbon
    /// <c>getText</c> getter: the last valid value the edit box displays and
    /// reverts to on an invalid entry. A pure snapshot read.
    /// </summary>
    /// <returns>The stored start date, or empty when none.</returns>
    internal virtual string GetPlotStartDate() => _state.PlotStartDate;

    /// <summary>
    /// Gets the stored explicit plot-finish date for the ribbon
    /// <c>getText</c> getter: the last valid value the edit box displays and
    /// reverts to on an invalid entry. A pure snapshot read.
    /// </summary>
    /// <returns>The stored finish date, or empty when none.</returns>
    internal virtual string GetPlotFinishDate() => _state.PlotFinishDate;

    /// <summary>
    /// Gets whether the plot-start end is automatic, for the ribbon
    /// <c>getChecked</c> getter. A pure snapshot read.
    /// </summary>
    /// <returns>True when the start derives from the data.</returns>
    internal virtual bool IsPlotStartAuto() => _state.PlotStartAuto;

    /// <summary>
    /// Gets whether the plot-finish end is automatic, for the ribbon
    /// <c>getChecked</c> getter. A pure snapshot read.
    /// </summary>
    /// <returns>True when the finish derives from the data.</returns>
    internal virtual bool IsPlotFinishAuto() => _state.PlotFinishAuto;

    /// <summary>
    /// Writes the four plot-range keys from the current snapshot through the
    /// injected catalogue writer. Centralises the per-end persist path so the
    /// four setters cannot construct divergent payloads.
    /// </summary>
    private void PersistPlotSettings()
    {
        var startMode = _state.PlotStartAuto ? nameof(PlotRangeMode.DataRange) : nameof(PlotRangeMode.Explicit);
        var finishMode = _state.PlotFinishAuto ? nameof(PlotRangeMode.DataRange) : nameof(PlotRangeMode.Explicit);
        var startDate = _state.PlotStartAuto ? string.Empty : _state.PlotStartDate;
        var finishDate = _state.PlotFinishAuto ? string.Empty : _state.PlotFinishDate;
        _ = _catalogueWriter?.WriteSettings(
            new Dictionary<string, string>
            {
                ["PlotStartMode"] = startMode,
                ["PlotFinishMode"] = finishMode,
                ["PlotStartDate"] = startDate,
                ["PlotFinishDate"] = finishDate,
            });
    }

    /// <summary>
    /// Arms the service for this session: subscribes the workbook-state events
    /// and performs the first refresh. Called once by <see cref="AddInHost"/>
    /// after the host sources are wired. Never throws.
    /// </summary>
    internal void Activate()
    {
        SubscribeToWorkbookStateChanges();
        Refresh();
    }

    /// <summary>
    /// Captures a fresh snapshot from the injected sources and invalidates the
    /// Ribbon. A source that reports "not determinable" (or throws) leaves the
    /// previous value in place, so a transient COM failure cannot grey a control.
    /// Never throws.
    /// </summary>
    internal void Refresh()
    {
        _state = CaptureSnapshot(_state);
        Invalidate();
    }

    /// <summary>
    /// Gets whether the control with <paramref name="controlId"/> is enabled.
    /// A pure read of the cached snapshot: no Excel probe, no ribbon call, no
    /// log write, and no exception for any input. Never throws.
    /// </summary>
    /// <param name="controlId">The Ribbon control ID, or null when unprobeable.</param>
    /// <returns>The control's enabled state (fail-open for unknown IDs).</returns>
    internal virtual bool GetEnabled(string? controlId) => _state.IsEnabled(controlId);

    /// <summary>
    /// Invalidates the whole Ribbon so Excel re-queries every dynamic getter
    /// against the current snapshot. A no-op without a handle; a failing
    /// <c>Invalidate</c> degrades to no invalidation. Never throws.
    /// </summary>
    internal void Invalidate() => InvokeRibbon(ribbon => ribbon.Invalidate());

    /// <summary>
    /// Invalidates a single Ribbon control so Excel re-queries its dynamic
    /// getters. A null or whitespace control ID is a no-op (nothing can be
    /// addressed), and a failing call degrades to no invalidation. Never
    /// throws.
    /// </summary>
    /// <param name="controlId">The Ribbon control ID to invalidate.</param>
    internal void InvalidateControl(string? controlId)
    {
        if (string.IsNullOrWhiteSpace(controlId))
        {
            return;
        }

        InvokeRibbon(ribbon => ribbon.InvalidateControl(controlId));
    }

    /// <summary>
    /// Resets the singleton and releases everything it holds: the event
    /// subscription (detached exactly once), the ribbon handle, the injected
    /// sources, and the snapshot. Called by <see cref="AddInHost.AutoClose"/> so
    /// the next Excel session starts fresh and unload leaves no COM connection
    /// behind (docs/03-ROADMAP.md R1.6).
    /// </summary>
    internal static void Reset()
    {
        lock (_instanceGate)
        {
            _instance?.Teardown();
            _instance = null;
        }
    }

    /// <summary>
    /// Builds the next snapshot from the injected sources. A source that is
    /// absent or reports "not determinable" leaves that fact at its previous
    /// value (snapshot stickiness). A source that throws keeps the entire
    /// previous snapshot.
    /// </summary>
    /// <param name="previous">The snapshot currently published.</param>
    /// <returns>The snapshot to publish.</returns>
    private RibbonState CaptureSnapshot(RibbonState previous)
    {
        // CA1031: the capture runs inside Ribbon callbacks and Excel events; a
        // source failure keeps the last known snapshot instead of propagating
        // into Excel.
#pragma warning disable CA1031
        try
        {
            var hasActiveWorkbook = _applicationAdapter?.HasActiveWorkbook() ?? null;
            var logAvailable = _logAvailabilitySource?.Invoke() ?? null;

            return DebugForceRibbonState.Apply(new RibbonState(
                hasActiveWorkbook ?? previous.HasActiveWorkbook,
                logAvailable ?? previous.LogAvailable,
                previous.PlotStartAuto,
                previous.PlotFinishAuto,
                previous.PlotStartDate,
                previous.PlotFinishDate));
        }
        catch
        {
            return previous;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Attaches the workbook-state change handler, replacing any previous
    /// subscription so a re-activation cannot double-subscribe.
    /// </summary>
    private void SubscribeToWorkbookStateChanges()
    {
        DisposeWorkbookStateSubscription();

        IExcelApplicationAdapter? adapter = _applicationAdapter;
        if (adapter is null)
        {
            return;
        }

        // CA1031: activation runs inside the never-throwing AutoOpen; a failure
        // to subscribe degrades to command-driven invalidation only.
#pragma warning disable CA1031
        try
        {
            _workbookStateSubscription = adapter.SubscribeWorkbookStateChanged(OnWorkbookStateChanged);
        }
        catch
        {
            _workbookStateSubscription = null;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Handles one workbook-state change: a single refresh (snapshot capture
    /// plus one ribbon invalidation). Never throws.
    /// </summary>
    private void OnWorkbookStateChanged() => Refresh();

    /// <summary>
    /// Detaches the workbook-state subscription exactly once. Never throws.
    /// </summary>
    private void DisposeWorkbookStateSubscription()
    {
        IDisposable? subscription = _workbookStateSubscription;
        _workbookStateSubscription = null;
        if (subscription is null)
        {
            return;
        }

        // CA1031: teardown must never throw into Excel; a failed detach
        // degrades to a retained event connection.
#pragma warning disable CA1031
        try
        {
            subscription.Dispose();
        }
        catch
        {
            // Intentionally empty: see the teardown rationale above.
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Applies <paramref name="action"/> to the ribbon handle when one is
    /// available. Never throws.
    /// </summary>
    /// <param name="action">The ribbon operation to perform.</param>
    private void InvokeRibbon(Action<IRibbonUI> action)
    {
        IRibbonUI? ribbon = _ribbon;
        if (ribbon is null)
        {
            return;
        }

        // CA1031: a failing ribbon call must never propagate into Excel — the
        // only consequence is a stale control until the next invalidation.
#pragma warning disable CA1031
        try
        {
            action(ribbon);
        }
        catch
        {
            // Intentionally empty: see the invalidation rationale above.
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Releases session state on <see cref="Reset"/>: detaches the subscription
    /// first, then drops every injected reference and the snapshot.
    /// </summary>
    private void Teardown()
    {
        DisposeWorkbookStateSubscription();
        _ribbon = null;
        _applicationAdapter = null;
        _logAvailabilitySource = null;
        _catalogueWriter = null;
        _state = RibbonState.Initial;
    }
}
