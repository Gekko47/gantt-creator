using System.Globalization;
using ExcelDna.Integration.CustomUI;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
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
    private IConfigCatalogueReader? _catalogueReader;
    private IGanttTableReader? _tableReader;
    private IPanelGridMeasurementPort? _panelGridMeasurementPort;
    private RibbonState _state = RibbonState.Initial;

    /// <summary>
    /// The metric token naming the frame's outer right padding, in points.
    /// Named once here rather than inlined so the display's chrome split and
    /// the Refresh pipeline's cannot drift (R4.8A D5: a literal here is a
    /// second authority).
    /// </summary>
    private const string _outerPaddingToken = "ChartOuterPaddingPt";

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
    /// Returns the injected catalogue writer, or <see langword="null"/> when
    /// the service was not armed. The Initialise-sheet command injects this
    /// instance into the workbook initialiser so the host-rejection evidence
    /// log (wired by <see cref="AddInHost"/>) reaches the writer the command
    /// actually runs — the initialiser's default writer has no log, so
    /// without this injection the <c>CatalogueWriteRejected</c> record never
    /// fires.
    /// </summary>
    internal IConfigCatalogueWriter? GetCatalogueWriter() => _catalogueWriter;

    /// <summary>
    /// Injects the catalogue reader that reads plot-range settings from the
    /// workbook configuration sheet. A null reader leaves the read step
    /// as a no-op, so the in-memory state is still correct.
    /// </summary>
    /// <param name="reader">The catalogue reader, or null when unavailable.</param>
    internal void SetCatalogueReader(IConfigCatalogueReader? reader) => _catalogueReader = reader;

    /// <summary>
    /// Injects the panel-grid measurement port that supplies the live text-panel
    /// width the plot-width display reads (R5.2). A null port leaves the width
    /// and height displays showing nothing rather than a guess, because there
    /// is no page to measure against.
    /// </summary>
    /// <param name="port">The measurement port, or null when unavailable.</param>
    internal void SetPanelGridMeasurementPort(IPanelGridMeasurementPort? port) => _panelGridMeasurementPort = port;

    /// <summary>
    /// Injects the table reader that derives the effective AUTO plot dates
    /// from the sheet's event data (fix plan ruling 1). A null reader leaves
    /// the read step as a no-op, so AUTO ends display empty until armed.
    /// </summary>
    /// <param name="reader">The table reader, or null when unavailable.</param>
    internal void SetTableReader(IGanttTableReader? reader) => _tableReader = reader;

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
    /// Sets the plot-start end to the state Excel reports (fix plan ruling 2:
    /// the checkbox honours Excel's pressed state rather than blindly
    /// toggling). Entering AUTO keeps the last explicit date in memory so it
    /// is restored when the user returns to explicit, and persists DataRange
    /// immediately. Unchecking AUTO seeds the edit box from the CURRENT plot
    /// bound (the month-snapped chart bound the AUTO box was displaying, per
    /// the owner ruling 2026-10-08) — never from the stale stored date — and
    /// persists Explicit immediately, so the next Refresh renders exactly
    /// what the box shows and no volatile "armed but unpersisted" window
    /// remains. Only when no current bound is known does the uncheck arm the
    /// edit box without persisting, because an Explicit mode with no date
    /// would break the next Refresh. A stored explicit date is read only
    /// after a user edit: a valid edit replaces it, an invalid edit reverts
    /// to it. Never refreshes the chart.
    /// </summary>
    /// <param name="pressed">True when Excel reports the AUTO box checked.</param>
    internal void SetPlotStartAuto(bool pressed)
    {
        var nextAuto = pressed;
        if (nextAuto)
        {
            _state = _state with
            {
                PlotStartAuto = true,

                // Entering AUTO preserves the typed date in memory (fix plan
                // ruling 1): it is no longer cleared, so the disabled edit box
                // keeps showing it until the effective-date derivation replaces
                // it, and returning to explicit restores it.
                PlotStartDate = _state.PlotStartDate,
                EffectivePlotStartDate = _state.EffectivePlotStartDate,
            };
        }
        else
        {
            // Leaving AUTO: seed BOTH the stored and the display date from
            // the current plot bound when one is known. The stored date is
            // deliberately NOT reused: it may predate the latest data, and
            // the box was showing the bound, so Explicit must start from
            // what the user saw. A parseable bound (always dd/MM/yyyy from
            // the derivation, but parsed defensively) persists at once.
            if (PlotRangeResolver.TryParsePlotDate(_state.EffectivePlotStartDate, out DateOnly bound))
            {
                var seeded = GanttDateFormatting.FormatDdMMMyy(bound);
                _state = _state with
                {
                    PlotStartAuto = false,
                    PlotStartDate = seeded,
                    EffectivePlotStartDate = seeded,
                };
            }
            else
            {
                _state = _state with
                {
                    PlotStartAuto = false,
                    PlotStartDate = _state.PlotStartDate,
                    EffectivePlotStartDate = _state.PlotStartDate,
                };

                if (string.IsNullOrWhiteSpace(_state.PlotStartDate))
                {
                    // Nothing coherent to persist yet: the date commit owns the
                    // Explicit write. The ribbon is still invalidated so the
                    // unchecked box and the enabled edit box repaint; without this
                    // the click's visual state never lands because ribbon getters
                    // only re-query on invalidate.
                    Invalidate();
                    return;
                }
            }
        }

#pragma warning disable CA1031
        try
        {
            CommandBoundary.Instance.Run("SetPlotStartAuto", PersistPlotSettings);
        }
        catch (Exception)
#pragma warning restore CA1031
        {
            // Degrade gracefully: the in-memory state is still correct,
            // and the next refresh re-reads the worksheet.
        }

        Invalidate();
    }

    /// <summary>
    /// Toggles the plot-start end between automatic (derived from the data)
    /// and explicit (the stored date). Retained for contract tests; the
    /// Excel-facing path is <see cref="SetPlotStartAuto(bool)"/>.
    /// </summary>
    internal void TogglePlotStartAuto() => SetPlotStartAuto(!_state.PlotStartAuto);

    /// <summary>
    /// Sets the plot-finish end to the state Excel reports (fix plan ruling 2:
    /// the checkbox honours Excel's pressed state rather than blindly
    /// toggling). Entering AUTO keeps the last explicit date in memory so it
    /// is restored when the user returns to explicit, and persists DataRange
    /// immediately. Unchecking AUTO seeds the edit box from the CURRENT plot
    /// bound (the month-snapped chart bound the AUTO box was displaying, per
    /// the owner ruling 2026-10-08) — never from the stale stored date — and
    /// persists Explicit immediately, so the next Refresh renders exactly
    /// what the box shows and no volatile "armed but unpersisted" window
    /// remains. Only when no current bound is known does the uncheck arm the
    /// edit box without persisting, because an Explicit mode with no date
    /// would break the next Refresh. A stored explicit date is read only
    /// after a user edit: a valid edit replaces it, an invalid edit reverts
    /// to it. Never refreshes the chart.
    /// </summary>
    /// <param name="pressed">True when Excel reports the AUTO box checked.</param>
    internal void SetPlotFinishAuto(bool pressed)
    {
        var nextAuto = pressed;
        if (nextAuto)
        {
            _state = _state with
            {
                PlotFinishAuto = true,
                PlotFinishDate = _state.PlotFinishDate,
                EffectivePlotFinishDate = _state.EffectivePlotFinishDate,
            };
        }
        else
        {
            // Leaving AUTO: seed BOTH the stored and the display date from
            // the current plot bound when one is known (see the start-end
            // note above for why the stored date is not reused).
            if (PlotRangeResolver.TryParsePlotDate(_state.EffectivePlotFinishDate, out DateOnly bound))
            {
                var seeded = GanttDateFormatting.FormatDdMMMyy(bound);
                _state = _state with
                {
                    PlotFinishAuto = false,
                    PlotFinishDate = seeded,
                    EffectivePlotFinishDate = seeded,
                };
            }
            else
            {
                _state = _state with
                {
                    PlotFinishAuto = false,
                    PlotFinishDate = _state.PlotFinishDate,
                    EffectivePlotFinishDate = _state.PlotFinishDate,
                };

                if (string.IsNullOrWhiteSpace(_state.PlotFinishDate))
                {
                    // Nothing coherent to persist yet: the date commit owns the
                    // Explicit write. The ribbon is still invalidated so the
                    // unchecked box and the enabled edit box repaint; without this
                    // the click's visual state never lands because ribbon getters
                    // only re-query on invalidate.
                    Invalidate();
                    return;
                }
            }
        }

#pragma warning disable CA1031
        try
        {
            CommandBoundary.Instance.Run("SetPlotFinishAuto", PersistPlotSettings);
        }
        catch (Exception)
#pragma warning restore CA1031
        {
            // Degrade gracefully: the in-memory state is still correct,
            // and the next refresh re-reads the worksheet.
        }

        Invalidate();
    }

    /// <summary>
    /// Toggles the plot-finish end between automatic (derived from the data)
    /// and explicit (the stored date). Retained for contract tests; the
    /// Excel-facing path is <see cref="SetPlotFinishAuto(bool)"/>.
    /// </summary>
    internal void TogglePlotFinishAuto() => SetPlotFinishAuto(!_state.PlotFinishAuto);

    /// <summary>
    /// Commits a user-typed plot-start date in explicit mode. Parses the exact
    /// plot-date input (dd-MMM-yy, dd/MM/yyyy legacy, plus the bare Excel serial
    /// the edit box echoes back), normalises it to dd-MMM-yy for storage, and
    /// pivots the start end to Explicit; an unparsable value is a no-op (the
    /// edit box keeps showing the previously stored date on the next getter
    /// query, which is the revert). Persists to the workbook configuration
    /// sheet so it survives a close/reopen, and updates the display value
    /// immediately since the commit happens outside a workbook load. Never
    /// refreshes the chart: the next Refresh chart applies the stored settings.
    /// </summary>
    /// <param name="dateText">The user-entered explicit plot start date text.</param>
    internal void SetPlotStartDate(string dateText)
    {
        if (!PlotRangeResolver.TryParsePlotDate(dateText, out DateOnly parsed))
        {
            return;
        }

        var normalised = GanttDateFormatting.FormatDdMMMyy(parsed);
        _state = _state with
        {
            PlotStartAuto = false,
            PlotStartDate = normalised,
            EffectivePlotStartDate = normalised,
        };

#pragma warning disable CA1031
        try
        {
            CommandBoundary.Instance.Run("SetPlotStartDate", PersistPlotSettings);
        }
        catch (Exception)
#pragma warning restore CA1031
        {
            // Degrade gracefully: the in-memory state is still correct,
            // and the next refresh re-reads the worksheet.
        }

        Invalidate();
    }

    /// <summary>
    /// Commits a user-typed plot-finish date in explicit mode. Parses the exact
    /// plot-date input (dd-MMM-yy, dd/MM/yyyy legacy, plus the bare Excel serial
    /// the edit box echoes back), normalises it to dd-MMM-yy for storage, and
    /// pivots the finish end to Explicit; an unparsable value is a no-op (the
    /// edit box keeps showing the previously stored date on the next getter
    /// query, which is the revert). Persists to the workbook configuration
    /// sheet so it survives a close/reopen, and updates the display value
    /// immediately since the commit happens outside a workbook load. Never
    /// refreshes the chart: the next Refresh chart applies the stored settings.
    /// </summary>
    /// <param name="dateText">The user-entered explicit plot finish date text.</param>
    internal void SetPlotFinishDate(string dateText)
    {
        if (!PlotRangeResolver.TryParsePlotDate(dateText, out DateOnly parsed))
        {
            return;
        }

        var normalised = GanttDateFormatting.FormatDdMMMyy(parsed);
        _state = _state with
        {
            PlotFinishAuto = false,
            PlotFinishDate = normalised,
            EffectivePlotFinishDate = normalised,
        };

#pragma warning disable CA1031
        try
        {
            CommandBoundary.Instance.Run("SetPlotFinishDate", PersistPlotSettings);
        }
        catch (Exception)
#pragma warning restore CA1031
        {
            // Degrade gracefully: the in-memory state is still correct,
            // and the next refresh re-reads the worksheet.
        }

        Invalidate();
    }

    /// <summary>
    /// Commits a user-selected plot time scale (R5.2). The value is validated
    /// against the closed <see cref="GanttTimeScale"/> set, so an unknown scale
    /// is refused rather than defaulted: a silently wrong scale would redraw the
    /// whole period band with the wrong calendar unit. Persists through the
    /// injected catalogue writer so it survives a close/reopen, and updates the
    /// in-memory state immediately since the commit happens outside a workbook
    /// load. Never refreshes the chart: the next Refresh chart applies the
    /// stored scale.
    /// </summary>
    /// <param name="timeScale">The selected scale, or null when the dropdown reported none.</param>
    internal void SetTimeScale(GanttTimeScale? timeScale)
    {
        if (timeScale is not GanttTimeScale scale)
        {
            return;
        }

        _state = _state with { TimeScale = scale };

#pragma warning disable CA1031
        try
        {
            CommandBoundary.Instance.Run("SetTimeScale", PersistPlotSettings);
        }
        catch (Exception)
#pragma warning restore CA1031
        {
            // Degrade gracefully: the in-memory state is still correct,
            // and the next refresh re-reads the worksheet.
        }

        Invalidate();
    }

    /// <summary>
    /// Commits a user-selected plot margin preset (R5.2). The value is validated
    /// against the closed <see cref="GanttPlotMargin"/> set, so an unknown preset
    /// is refused rather than defaulted. When the preset is
    /// <see cref="GanttPlotMargin.Custom"/>, the currently stored custom
    /// centimetre value is carried through; a preset ignores it. Persists
    /// through the injected catalogue writer. Never refreshes the chart.
    /// </summary>
    /// <param name="margin">The selected margin preset, or null when the combobox reported none.</param>
    internal void SetMargin(GanttPlotMargin? margin)
    {
        if (margin is not GanttPlotMargin selected)
        {
            return;
        }

        var customCm = selected == GanttPlotMargin.Custom ? _state.MarginCm : GanttPlotMargins.DefaultCustomCm;
        _state = _state with { Margin = selected, MarginCm = customCm };

#pragma warning disable CA1031
        try
        {
            CommandBoundary.Instance.Run("SetMargin", PersistPlotSettings);
        }
        catch (Exception)
#pragma warning restore CA1031
        {
            // Degrade gracefully: the in-memory state is still correct,
            // and the next refresh re-reads the worksheet.
        }

        Invalidate();
    }

    /// <summary>
    /// Commits a user-entered custom plot margin, in centimetres (R5.2). Used
    /// only when the stored preset is <see cref="GanttPlotMargin.Custom"/>; an
    /// out-of-range value is clamped to the permitted band rather than refused,
    /// so a stale stored value cannot break a Refresh. Persists through the
    /// injected catalogue writer. Never refreshes the chart.
    /// </summary>
    /// <param name="customCm">The custom margin, in centimetres.</param>
    internal void SetMarginCm(double customCm)
    {
        var clamped = Math.Clamp(customCm, GanttPlotMargins.MinimumCustomCm, GanttPlotMargins.MaximumCustomCm);
        if (!double.IsFinite(clamped))
        {
            clamped = GanttPlotMargins.DefaultCustomCm;
        }

        _state = _state with { MarginCm = clamped };

#pragma warning disable CA1031
        try
        {
            CommandBoundary.Instance.Run("SetMarginCm", PersistPlotSettings);
        }
        catch (Exception)
#pragma warning restore CA1031
        {
            // Degrade gracefully: the in-memory state is still correct,
            // and the next refresh re-reads the worksheet.
        }

        Invalidate();
    }

    /// <summary>
    /// Commits a user-selected output size preset (R5.2). The value is validated
    /// against the closed <see cref="SizePresets"/> catalogue, so an unknown key
    /// is refused rather than defaulted: a silently wrong preset would render
    /// the chart on the wrong paper. Persists through the injected catalogue
    /// writer and recomputes the width/height displays from the new preset.
    /// Never refreshes the chart.
    /// </summary>
    /// <param name="preset">The selected preset, or null when none was supplied.</param>
    internal void SetPreset(SizePreset? preset)
    {
        if (preset is null)
        {
            return;
        }

        // The preset is committed first: C# evaluates every `with` initializer
        // against the OLD snapshot, so measuring inside the same expression
        // would read the previous (possibly null) preset and leave the width
        // display one click behind. The height comes straight from the preset.
        _state = _state with { Preset = preset };
        _state = _state with { PlotWidthPt = MeasurePlotWidth(), PlotHeightPt = preset.HeightPt };

#pragma warning disable CA1031
        try
        {
            CommandBoundary.Instance.Run("SetPreset", PersistPlotSettings);
        }
        catch (Exception)
#pragma warning restore CA1031
        {
            // Degrade gracefully: the in-memory state is still correct,
            // and the next refresh re-reads the worksheet.
        }

        Invalidate();
    }

    /// <summary>
    /// Gets the date the start edit box displays for the ribbon
    /// <c>getText</c> getter: the explicit date in explicit mode, the derived
    /// data range start in automatic mode (fix plan ruling 1), or empty when
    /// neither is known. A pure snapshot read.
    /// </summary>
    /// <returns>The display start date, or empty when none.</returns>
    internal virtual string GetPlotStartDate() => _state.EffectivePlotStartDate;

    /// <summary>
    /// Gets the date the finish edit box displays for the ribbon
    /// <c>getText</c> getter: the explicit date in explicit mode, the derived
    /// data range finish in automatic mode (fix plan ruling 1), or empty when
    /// neither is known. A pure snapshot read.
    /// </summary>
    /// <returns>The display finish date, or empty when none.</returns>
    internal virtual string GetPlotFinishDate() => _state.EffectivePlotFinishDate;

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
    /// Gets the stored plot time scale for the ribbon dropdown's
    /// <c>getSelectedItemIndex</c> getter. A pure snapshot read.
    /// </summary>
    /// <returns>The stored scale.</returns>
    internal virtual GanttTimeScale GetTimeScale() => _state.TimeScale;

    /// <summary>
    /// Gets the stored plot margin preset for the ribbon combobox's
    /// <c>getSelectedItemIndex</c> getter. A pure snapshot read.
    /// </summary>
    /// <returns>The stored margin preset.</returns>
    internal virtual GanttPlotMargin GetMargin() => _state.Margin;

    /// <summary>
    /// Gets the stored custom margin, in centimetres, for the ribbon's custom
    /// margin display. A pure snapshot read.
    /// </summary>
    /// <returns>The stored custom margin.</returns>
    internal virtual double GetMarginCm() => _state.MarginCm;

    /// <summary>
    /// Gets the stored output size preset for the ribbon preset buttons'
    /// pressed-state getter. A pure snapshot read.
    /// </summary>
    /// <returns>The stored preset, or null when the workbook names none.</returns>
    internal virtual SizePreset? GetPreset() => _state.Preset;

    /// <summary>
    /// Gets the plot width the resolver would derive, for the ribbon's
    /// read-only Width display (R5.2). A pure snapshot read.
    /// </summary>
    /// <returns>The plot width in points, or empty when no preset is known.</returns>
    internal virtual string GetPlotWidth() => _state.PlotWidthPt?.ToString("F2", CultureInfo.InvariantCulture) ?? string.Empty;

    /// <summary>
    /// Gets the plot height the resolver would derive, for the ribbon's
    /// read-only Height display (R5.2). A pure snapshot read.
    /// </summary>
    /// <returns>The plot height in points, or empty when no preset is known.</returns>
    internal virtual string GetPlotHeight() => _state.PlotHeightPt?.ToString("F2", CultureInfo.InvariantCulture) ?? string.Empty;

    /// <summary>
    /// Recomputes the plot width the resolver would derive from the current
    /// preset and a live panel measurement, so a preset change is reflected in
    /// the read-only display without a Refresh. Uses the SAME subtraction
    /// <see cref="PlotGeometryResolver.MeasurePlotWidth"/> performs, read once
    /// through that authority so the displayed figure and the rendered plot
    /// cannot drift. Returns null when no preset is known or the measurement
    /// refuses, which the display renders as empty rather than a guess.
    /// </summary>
    /// <returns>The plot width in points, or null.</returns>
    private double? MeasurePlotWidth()
    {
        SizePreset? preset = _state.Preset;
        if (preset is null)
        {
            return null;
        }

        IPanelGridMeasurementPort? port = _panelGridMeasurementPort;
        if (port is null)
        {
            return null;
        }

#pragma warning disable CA1031
        try
        {
            PanelGridOutcome outcome = port.Measure(MeasuredColumns());
            if (!outcome.Succeeded || outcome.Grid is null)
            {
                return null;
            }

            PanelCellGrid grid = outcome.Grid;
            var marginPt = GanttPlotMargins.MarginPt(_state.Margin, _state.MarginCm);
            // Same chrome split the Refresh pipeline uses (ADR-0031 D2): the
            // plot is flush-left, so the left chrome is zero and the R5.2 margin
            // (twice) plus the frame's outer padding rides on the right.
            var chrome = GanttCatalogues.MetricDefault(_outerPaddingToken);
            return PlotGeometryResolver.MeasurePlotWidth(
                preset,
                grid.TotalWidthPt,
                leftChromePt: 0,
                rightChromePt: (marginPt * 2) + chrome);
        }
        catch
        {
            return null;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// The display names of the visible table columns, in panel order — the
    /// same list the live measurement port measures. Derived from the schema
    /// rather than hand-written so a hidden column is never measured (R4.8A
    /// D5: a hidden column reports width 0 and the grid refuses it).
    /// </summary>
    private static List<string> MeasuredColumns() =>
        [.. GanttTableSchema.Default.Columns
            .Where(column => !column.IsHidden)
            .Select(column => column.Name)];

    /// <summary>
    /// Writes the ten plot-layout keys from the current snapshot through the
    /// injected catalogue writer. Centralises the per-setting persist path so
    /// the six setters cannot construct divergent payloads. The four
    /// plot-range keys behave as before (fix plan ruling 1: the stored
    /// explicit dates persist even while their end is AUTO, so they survive
    /// the round trip and are restored on return to explicit; the Refresh
    /// pipeline still treats an AUTO end as data-derived and ignores its date
    /// text, so persisting it cannot blend the modes). The six R5.2 layout
    /// keys are written alongside them so a single commit keeps the scale,
    /// the margin pair, and the preset consistent with the dates.
    /// Reports ConfigWriteOutcome refusals and write exceptions to the user
    /// via CommandBoundary.
    /// </summary>
    private void PersistPlotSettings()
    {
        var startMode = _state.PlotStartAuto ? nameof(PlotRangeMode.DataRange) : nameof(PlotRangeMode.Explicit);
        var finishMode = _state.PlotFinishAuto ? nameof(PlotRangeMode.DataRange) : nameof(PlotRangeMode.Explicit);

        var settings = new Dictionary<string, string>
        {
            ["PlotStartMode"] = startMode,
            ["PlotFinishMode"] = finishMode,
            ["PlotStartDate"] = _state.PlotStartDate,
            ["PlotFinishDate"] = _state.PlotFinishDate,
            ["TimeScale"] = _state.TimeScale.ToString(),
            // The period label format is derived from the scale through the
            // compatibility table (GanttChartSettings.IsCompatible), so the
            // pair is always valid. Each scale has exactly one canonical
            // format: MMM for months, Qn for quarters, Wnn for weeks.
            ["PeriodLabelFormat"] = _state.TimeScale switch
            {
                GanttTimeScale.Month => nameof(GanttPeriodLabelFormat.MMM),
                GanttTimeScale.Quarter => nameof(GanttPeriodLabelFormat.Quarter),
                GanttTimeScale.Week => nameof(GanttPeriodLabelFormat.Week),
                _ => nameof(GanttPeriodLabelFormat.MMM),
            },
            ["Margin"] = _state.Margin.ToString(),
            ["MarginCm"] = _state.MarginCm.ToString("R", CultureInfo.InvariantCulture),
            ["SizePreset"] = _state.Preset?.Key ?? SizePresets.Default.Key,
        };

        ConfigWriteOutcome? outcome = _catalogueWriter?.WriteSettings(settings);
        if (outcome is not null && !outcome.Succeeded)
        {
            var message = $"Failed to persist plot settings: {outcome.Refusal}";
            CommandBoundary.Instance.Run("PersistPlotSettings", () => throw new InvalidOperationException(message));
        }
    }

    /// <summary>
    /// Loads the four plot-range settings from the active workbook into the
    /// current state snapshot. Uses the injected IConfigCatalogueReader to read
    /// the settings table. A refused or failed read keeps the in-memory state
    /// untouched (fix for the stuck checkbox: the post-click refresh must not
    /// overwrite a just-set mode with stale snapshot data when the workbook
    /// cannot be read mid-click). A successful read adopts the workbook's
    /// modes and explicit dates, normalising any parseable explicit date
    /// (bare Excel serials from pre-fix workbooks, unpadded input) to
    /// dd-MMM-yy so the edit box never shows a raw number; an unparsable
    /// stored date is kept verbatim so the failure stays visible instead of
    /// being silently replaced. Then derives the effective display dates so
    /// disabled AUTO edit boxes show the current data range date.
    /// Mode text parses through the single PlotRangeModes authority, so the
    /// display and the Refresh pipeline cannot disagree on what AUTO means.
    /// Never throws.
    /// </summary>
    private void LoadPlotSettingsFromWorkbook()
    {
        if (_catalogueReader is null)
        {
            return;
        }

        // CA1031: the capture runs inside Ribbon callbacks and Excel events; a
        // source failure keeps the last known snapshot instead of propagating
        // into Excel.
#pragma warning disable CA1031
        try
        {
            ConfigReadOutcome outcome = _catalogueReader.Read();
            if (outcome.Succeeded && outcome.Settings is not null)
            {
                IReadOnlyDictionary<string, string> settings = outcome.Settings;
                var startMode = settings.TryGetValue("PlotStartMode", out var startModeVal) ? startModeVal : string.Empty;
                var finishMode = settings.TryGetValue("PlotFinishMode", out var finishModeVal) ? finishModeVal : string.Empty;
                var startDate = NormaliseStoredPlotDate(
                    settings.TryGetValue("PlotStartDate", out var startDateVal) ? startDateVal : string.Empty);
                var finishDate = NormaliseStoredPlotDate(
                    settings.TryGetValue("PlotFinishDate", out var finishDateVal) ? finishDateVal : string.Empty);

                var startAuto = PlotRangeModes.TryParse(startMode, out PlotRangeMode parsedStart)
                    && parsedStart == PlotRangeMode.DataRange;
                var finishAuto = PlotRangeModes.TryParse(finishMode, out PlotRangeMode parsedFinish)
                    && parsedFinish == PlotRangeMode.DataRange;

                // R5.2 layout settings. Each is parsed through its closed Core
                // authority and falls back to the catalogue default on an
                // absent or unknown stored value, so a workbook that never set
                // a scale renders identically to a freshly initialised one.
                GanttTimeScale timeScale = ReadTimeScale(settings);
                GanttPlotMargin margin = ReadMargin(settings);
                var marginCm = ReadMarginCm(settings);
                SizePreset? preset = ReadPreset(settings);

                _state = _state with
                {
                    PlotStartAuto = startAuto,
                    PlotFinishAuto = finishAuto,
                    PlotStartDate = startDate,
                    PlotFinishDate = finishDate,
                    TimeScale = timeScale,
                    Margin = margin,
                    MarginCm = marginCm,
                    Preset = preset,
                    PlotWidthPt = MeasurePlotWidth(),
                    PlotHeightPt = preset?.HeightPt,
                };

                RefreshEffectivePlotDates(settings);
            }
        }
        catch
        {
            // Degrade gracefully: the in-memory state is still correct,
            // and the next refresh re-reads the worksheet.
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Reads the stored plot time scale, falling back to the catalogue default
    /// when the key is absent or names an unknown scale (R5.2).
    /// </summary>
    private static GanttTimeScale ReadTimeScale(IReadOnlyDictionary<string, string> settings) =>
        settings.TryGetValue("TimeScale", out var text)
            && GanttChartSettings.TryParseTimeScale(text, out GanttTimeScale scale)
            ? scale
            : GanttTimeScale.Month;

    /// <summary>
    /// Reads the stored plot margin preset, falling back to the Windows
    /// "Normal" default when the key is absent or names an unknown preset
    /// (R5.2).
    /// </summary>
    private static GanttPlotMargin ReadMargin(IReadOnlyDictionary<string, string> settings) =>
        settings.TryGetValue("Margin", out var text)
            && GanttPlotMargins.TryParse(text, out GanttPlotMargin margin)
            ? margin
            : GanttPlotMargins.Default;

    /// <summary>
    /// Reads the stored custom margin, in centimetres (R5.2). Used only when
    /// the preset is <see cref="GanttPlotMargin.Custom"/>; an absent or
    /// unparsable value falls back to the catalogue default, and
    /// <see cref="GanttPlotMargins.MarginPt"/> clamps an out-of-range value
    /// rather than producing a negative or unbounded margin.
    /// </summary>
    private static double ReadMarginCm(IReadOnlyDictionary<string, string> settings) =>
        settings.TryGetValue("MarginCm", out var raw)
            && double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed)
            && double.IsFinite(parsed)
            ? parsed
            : GanttPlotMargins.DefaultCustomCm;

    /// <summary>
    /// Reads the stored output size preset, returning null when the key is
    /// absent or names an unknown preset (R5.2). A null preset means the
    /// width and height displays show nothing rather than a guess.
    /// </summary>
    private static SizePreset? ReadPreset(IReadOnlyDictionary<string, string> settings) =>
        settings.TryGetValue("SizePreset", out var text) ? SizePresets.ByKey(text) : null;

    /// <summary>
    /// Derives the effective display dates for both edit boxes (fix plan
    /// ruling 1): the explicit date in explicit mode, the TRUE chart plot
    /// bound in automatic mode (the month-snapped extent Refresh renders, not
    /// the raw padded data extent), or empty when neither is known. Uses the
    /// same ingredients Refresh consumes (table rows, row validation, the
    /// plot resolver, then the per-end padding-then-snap), so the disabled
    /// AUTO box shows exactly what Refresh would render. A failed derivation
    /// keeps the last known display value rather than blanking the box.
    /// Never throws.
    /// </summary>
    /// <param name="settings">The effective settings map just loaded from the workbook.</param>
    private void RefreshEffectivePlotDates(IReadOnlyDictionary<string, string> settings)
    {
        // CA1031: runs inside Ribbon callbacks and Excel events; any failure
        // keeps the last known display values.
#pragma warning disable CA1031
        try
        {
            var explicitStart = _state.PlotStartAuto ? string.Empty : _state.PlotStartDate;
            var explicitFinish = _state.PlotFinishAuto ? string.Empty : _state.PlotFinishDate;

            var effectiveStart = explicitStart;
            var effectiveFinish = explicitFinish;

            if (_state.PlotStartAuto || _state.PlotFinishAuto)
            {
                GanttTableReadOutcome? table = _tableReader?.Read();
                if (table is { Succeeded: true } && table.Rows is not null)
                {
                    GanttValidationOutcome validation = GanttRowValidator.Validate(table.Rows);
                    var padding = ReadRangePaddingDays(settings);
                    PlotRangeOutcome range = PlotRangeResolver.TryResolve(
                        validation.Events,
                        _state.PlotStartAuto ? nameof(PlotRangeMode.DataRange) : nameof(PlotRangeMode.Explicit),
                        _state.PlotFinishAuto ? nameof(PlotRangeMode.DataRange) : nameof(PlotRangeMode.Explicit),
                        _state.PlotStartDate,
                        _state.PlotFinishDate,
                        padding);

                    if (range.Succeeded && range.Start is { } resolvedStart && range.Finish is { } resolvedFinish)
                    {
                        // The TRUE chart bounds: per-end padding-then-snap
                        // through the single PlotMonthBounds authority Refresh
                        // renders (owner ruling 2026-10-02). DataRange ends are
                        // month-snapped; explicit ends pass through verbatim.
                        // Mirrors TryResolveDateRange in ExcelSceneBuildRequestFactory.
                        PlotRangeMode startMode = _state.PlotStartAuto ? PlotRangeMode.DataRange : PlotRangeMode.Explicit;
                        PlotRangeMode finishMode = _state.PlotFinishAuto ? PlotRangeMode.DataRange : PlotRangeMode.Explicit;

                        DateOnly start = resolvedStart;
                        DateOnly finish = resolvedFinish;

                        if (startMode == PlotRangeMode.DataRange)
                        {
                            start = PlotMonthBounds.SnapToMonthStart(start);
                        }

                        if (finishMode == PlotRangeMode.DataRange)
                        {
                            finish = PlotMonthBounds.SnapToMonthEnd(finish);
                        }

                        // Degenerate guard: if snap moved finish on or before
                        // start, keep last display values.
                        if (finish.DayNumber <= start.DayNumber)
                        {
                            // Degenerate; last display values stand.
                        }
                        else
                        {
                            if (_state.PlotStartAuto)
                            {
                                effectiveStart = GanttDateFormatting.FormatDdMMMyy(start);
                            }

                            if (_state.PlotFinishAuto)
                            {
                                effectiveFinish = GanttDateFormatting.FormatDdMMMyy(finish);
                            }
                        }
                    }
                }
            }

            _state = _state with
            {
                EffectivePlotStartDate = effectiveStart,
                EffectivePlotFinishDate = effectiveFinish,
            };
        }
        catch
        {
            // Degrade gracefully: the last known display values stand.
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Normalises one stored explicit plot date for display and state: a
    /// parseable value (dd-MMM-yy, dd/MM/yyyy, or the bare Excel serial the
    /// edit box echoes) becomes canonical dd-MMM-yy; blank stays blank; an
    /// unparsable value is kept verbatim so the failure stays visible.
    /// Pure and never throws.
    /// </summary>
    /// <param name="stored">The stored date text.</param>
    /// <returns>The display-ready date text.</returns>
    private static string NormaliseStoredPlotDate(string stored)
    {
        return string.IsNullOrWhiteSpace(stored)
            ? string.Empty
            : PlotRangeResolver.TryParsePlotDate(stored, out DateOnly parsed)
                ? GanttDateFormatting.FormatDdMMMyy(parsed)
                : stored;
    }

    /// <summary>
    /// Reads the automatic-mode range padding from the settings map, mirroring
    /// the Refresh pipeline's default: a stored value outside the valid range
    /// falls back to the same default Refresh uses.
    /// </summary>
    /// <param name="settings">The effective settings map.</param>
    /// <returns>The non-negative padding in days per side.</returns>
    private static int ReadRangePaddingDays(IReadOnlyDictionary<string, string> settings)
    {
        return settings.TryGetValue("RangePaddingDays", out var text)
            && int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed)
            && parsed >= 0
            ? parsed
            : 3;
    }

    /// <summary>
    /// Arms the service for this session: subscribes the workbook-state events,
    /// loads plot settings from the active workbook, and performs the first
    /// refresh. Called once by <see cref="AddInHost"/> after the host sources
    /// are wired. Never throws.
    /// </summary>
    internal void Activate()
    {
        SubscribeToWorkbookStateChanges();
        LoadPlotSettingsFromWorkbook();
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
        LoadPlotSettingsFromWorkbook();

        // Probe whether the active workbook has an initialised Gantt sheet.
        // TableMissing means the sheet hasn't been initialised yet; other
        // refusals mean it exists but is corrupted (still counts as initialised
        // for gating purposes — the user can still repair). NoActiveWorkbook
        // means no workbook at all.
        var sheetInitialised = false;
        if (_tableReader is not null)
        {
            GanttTableReadOutcome tableOutcome = _tableReader.Read();
            sheetInitialised = tableOutcome.Refusal != GanttTableReadRefusalReason.TableMissing;
        }

        _state = CaptureSnapshot(_state, sheetInitialised);
        Invalidate();
    }

    /// <summary>
    /// Refreshes the ribbon after a plot control commit WITHOUT re-reading
    /// the workbook. Retained for contract tests; the setters now invalidate
    /// directly, so the click path no longer calls this. Never throws.
    /// </summary>
    internal void RefreshAfterPlotClick()
    {
        // CA1031: runs inside Ribbon callbacks; any failure degrades to the
        // last known snapshot rather than propagating into Excel.
#pragma warning disable CA1031
        try
        {
            _state = CaptureSnapshot(_state, _state.SheetInitialised);
            Invalidate();
        }
        catch
        {
            // Degrade gracefully: the in-memory state is still correct.
        }
#pragma warning restore CA1031
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
    /// <param name="sheetInitialised">Whether the active workbook has an initialised Gantt sheet.</param>
    /// <returns>The snapshot to publish.</returns>
    private RibbonState CaptureSnapshot(RibbonState previous, bool sheetInitialised)
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
                sheetInitialised,
                previous.PlotStartAuto,
                previous.PlotFinishAuto,
                previous.PlotStartDate,
                previous.PlotFinishDate,
                previous.EffectivePlotStartDate,
                previous.EffectivePlotFinishDate,
                previous.TimeScale,
                previous.Margin,
                previous.MarginCm,
                previous.Preset,
                previous.PlotWidthPt,
                previous.PlotHeightPt));
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
        _catalogueReader = null;
        _tableReader = null;
        _panelGridMeasurementPort = null;
        _state = RibbonState.Initial;
    }
}
