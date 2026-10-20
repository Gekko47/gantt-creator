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

    /// <summary>The standalone selection identifier (R5.5 ruling 6), or null when not wired.</summary>
    private ISelectionIdentifier? _selectionIdentifier;

    /// <summary>The selection-change subscription, detached exactly once.</summary>
    private IDisposable? _selectionSubscription;

    private IConfigCatalogueWriter? _catalogueWriter;
    private IConfigCatalogueReader? _catalogueReader;
    private IGanttTableReader? _tableReader;
    private IPanelGridMeasurementPort? _panelGridMeasurementPort;
    private RibbonState _state = RibbonState.Initial;

    /// <summary>
    /// The description-label dropdown's per-selection context: whether the selection
    /// resolves to exactly one table body row, that row's stored position, and
    /// the positions the row's type (or Custom Activity style) permits.
    /// </summary>
    /// <param name="Available">True when exactly one body row is selected.</param>
    /// <param name="Selected">The row's stored position, or null when blank/unknown.</param>
    /// <param name="Options">The row's allowed positions in display order.</param>
    private readonly record struct DescriptionLabelContext(
        bool Available,
        GanttLabelPosition? Selected,
        IReadOnlyList<GanttLabelPosition>? Options);

    /// <summary>
    /// The date-label dropdown's per-selection context: whether the
    /// selection resolves to exactly one table body row, that row's
    /// stored choice, and the choices the row's type permits (R5.14).
    /// The capability is type-only (D4): a Custom Activity style carries
    /// no date mode, so no style registry is consulted.
    /// </summary>
    /// <param name="Available">True when exactly one body row is selected.</param>
    /// <param name="Selected">The row's stored choice, or null when blank/unknown.</param>
    /// <param name="Options">The row's allowed choices in display order.</param>
    private readonly record struct DateLabelContext(
        bool Available,
        GanttDateLabelPosition? Selected,
        IReadOnlyList<GanttDateLabelPosition>? Options);

    /// <summary>
    /// The date-label dropdown's per-selection context in the
    /// positionable vocabulary (owner ruling 2026-10-14): whether
    /// the selection resolves to exactly one table body row whose
    /// date label is positionable (the delineator), that row's
    /// stored corner, and the corners the row permits.
    /// </summary>
    /// <param name="Available">True when exactly one body row is selected.</param>
    /// <param name="Selected">The row's stored corner, or null when blank (Auto).</param>
    /// <param name="Options">The row's allowed corners in display order.</param>
    private readonly record struct DateLabelPlacementContext(
        bool Available,
        GanttLabelPosition? Selected,
        IReadOnlyList<GanttLabelPosition>? Options);

    /// <summary>
    /// The custom-style dropdown's per-selection context (R5.6a): whether the
    /// selection resolves to exactly one style-key-type body row, that row's
    /// stored style key (only when it names a user-created style), and the
    /// user-created (non-built-in) style keys the configuration registry
    /// offers, in catalogue order.
    /// </summary>
    /// <param name="Available">True when exactly one style-key-type body row is selected.</param>
    /// <param name="SelectedKey">The row's stored custom style key, or null when blank/built-in.</param>
    /// <param name="Options">The user-created style keys in catalogue order.</param>
    private readonly record struct CustomStyleContext(
        bool Available,
        string? SelectedKey,
        IReadOnlyList<string>? Options);

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
    /// The presenter for a user-facing failure. It is a seam rather than a
    /// direct call to <see cref="CommandErrorDialog.Show"/> because that
    /// dialog is modal: a test that drives a refused write would block on a
    /// message box with nobody to dismiss it, so the test could never reach
    /// its own assertion. Every other command in this assembly takes the same
    /// seam (see <see cref="RefreshSheetCommand.Run"/>).
    /// </summary>
    private Action<string> _errorPresenter = CommandErrorDialog.Show;

    /// <summary>
    /// Injects the presenter used for a user-facing failure.
    /// </summary>
    /// <param name="presenter">The presenter; a null presenter restores the default dialog.</param>
    internal void SetErrorPresenter(Action<string>? presenter) =>
        _errorPresenter = presenter ?? CommandErrorDialog.Show;

    /// <summary>
    /// Injects the catalogue writer that persists plot-range settings to the
    /// workbook configuration sheet. A null writer leaves the persist step
    /// as a no-op, so the in-memory state is still correct and the next
    /// refresh re-reads the worksheet.
    /// </summary>
    /// <param name="writer">The catalogue writer, or null when unavailable.</param>
    internal void SetCatalogueWriter(IConfigCatalogueWriter? writer) => _catalogueWriter = writer;

    /// <summary>
    /// Injects the standalone selection identifier (R5.5 ruling 6) that feeds
    /// the description-label dropdown's per-row gating and item list. A null
    /// identifier leaves the description-label context unavailable, so the dropdown stays
    /// disabled until one is wired. Subsequent calls replace any previous
    /// subscription via <see cref="Activate"/>.
    /// </summary>
    /// <param name="identifier">The selection identifier, or null when unavailable.</param>
    internal void SetSelectionIdentifier(ISelectionIdentifier? identifier) =>
        _selectionIdentifier = identifier;

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
        bool nextAuto = pressed;
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
                string seeded = GanttDateFormatting.FormatDdMMMyy(bound);
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
        bool nextAuto = pressed;
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
                string seeded = GanttDateFormatting.FormatDdMMMyy(bound);
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

        string normalised = GanttDateFormatting.FormatDdMMMyy(parsed);
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

        string normalised = GanttDateFormatting.FormatDdMMMyy(parsed);
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
    /// whole period band with the wrong calendar unit.
    /// <para>
    /// The month band carries two label forms, so the dropdown's two Month rows
    /// name one explicitly through <paramref name="periodLabelFormat"/>. A
    /// suffix-less selection (<c>Quarter</c>, <c>Week</c>) passes null and takes
    /// the scale's canonical form; a named form the scale cannot carry is
    /// refused rather than coerced, and an incompatible stored form is
    /// normalised, so the stored pair the catalogue reader validates is always
    /// one the dropdown could have produced.
    /// </para>
    /// Persists through the injected catalogue writer so it survives a
    /// close/reopen, and updates the in-memory state immediately since the
    /// commit happens outside a workbook load. Never refreshes the chart: the
    /// next Refresh chart applies the stored scale.
    /// </summary>
    /// <param name="timeScale">The selected scale, or null when the dropdown reported none.</param>
    /// <param name="periodLabelFormat">
    /// The explicitly selected label format, or null when the selection names
    /// only the scale.
    /// </param>
    internal void SetTimeScale(GanttTimeScale? timeScale, GanttPeriodLabelFormat? periodLabelFormat = null)
    {
        if (timeScale is not GanttTimeScale scale)
        {
            return;
        }

        // A named form the scale cannot carry is refused, not coerced: storing
        // it would write a pair the catalogue reader rejects.
        if (periodLabelFormat is GanttPeriodLabelFormat requested
            && !GanttChartSettings.IsCompatible(scale, requested))
        {
            return;
        }

        // The pair is normalised here, not at persist time, so the dropdown and
        // the stored setting agree the moment the selection changes.
        GanttPeriodLabelFormat format = periodLabelFormat is GanttPeriodLabelFormat selected
            ? selected
            : GanttChartSettings.IsCompatible(scale, _state.PeriodLabelFormat)
                ? _state.PeriodLabelFormat
                : CanonicalPeriodLabelFormat(scale);

        _state = _state with { TimeScale = scale, PeriodLabelFormat = format };

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

        double customCm = selected == GanttPlotMargin.Custom ? _state.MarginCm : GanttPlotMargins.DefaultCustomCm;
        // Commit first, then re-measure: C# evaluates every `with`
        // initializer against the OLD snapshot, so measuring inside
        // the same expression would read the previous margin and
        // leave the width display one click behind (same two-step
        // shape as SetPreset, so the displayed figure and the
        // rendered plot cannot drift — R5.2 D4).
        _state = _state with { Margin = selected, MarginCm = customCm };
        _state = _state with { PlotWidthPt = MeasurePlotWidth() };

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
        double clamped = Math.Clamp(customCm, GanttPlotMargins.MinimumCustomCm, GanttPlotMargins.MaximumCustomCm);
        if (!double.IsFinite(clamped))
        {
            clamped = GanttPlotMargins.DefaultCustomCm;
        }

        _state = _state with { MarginCm = clamped };
        // Re-measure after the commit (see SetMargin) so the width
        // display tracks the new margin rather than the previous one.
        _state = _state with { PlotWidthPt = MeasurePlotWidth() };

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
    /// Gets the stored period label format. Read alongside
    /// <see cref="GetTimeScale"/> by the ribbon's <c>getSelectedItemIndex</c>
    /// getter: the Month scale owns two dropdown rows (one per label form), so
    /// the selected index is derived from the PAIR, not from the scale alone.
    /// A pure snapshot read.
    /// </summary>
    /// <returns>The stored format.</returns>
    internal virtual GanttPeriodLabelFormat GetPeriodLabelFormat() => _state.PeriodLabelFormat;

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
            double marginPt = GanttPlotMargins.MarginPt(_state.Margin, _state.MarginCm);
            // Same chrome split the Refresh pipeline uses (ADR-0031 D2): the
            // plot is flush-left, so the left chrome is zero and the R5.2 margin
            // (twice) plus the frame's outer padding rides on the right.
            double chrome = GanttCatalogues.MetricDefault(_outerPaddingToken);
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
        string startMode = _state.PlotStartAuto ? nameof(PlotRangeMode.DataRange) : nameof(PlotRangeMode.Explicit);
        string finishMode = _state.PlotFinishAuto ? nameof(PlotRangeMode.DataRange) : nameof(PlotRangeMode.Explicit);

        // Validate before persisting so invalid settings never reach the workbook.
        // R5.1: ValidateSettings is wired here to reject bad modes, missing dates,
        // unparsable dates, and start-after-finish before any workbook write.
        PlotRangeResolver.ValidateSettings(new PlotRangeSettings
        {
            StartMode = startMode,
            FinishMode = finishMode,
            StartDate = _state.PlotStartDate,
            FinishDate = _state.PlotFinishDate,
        });

        Dictionary<string, string> values = new Dictionary<string, string>
        {
            ["PlotStartMode"] = startMode,
            ["PlotFinishMode"] = finishMode,
            ["PlotStartDate"] = _state.PlotStartDate,
            ["PlotFinishDate"] = _state.PlotFinishDate,
            ["TimeScale"] = _state.TimeScale.ToString(),
            // The month band carries two label forms (MM and MMM), so the
            // format is the user's stored selection rather than a derivation
            // of the scale. The pair is kept compatible by the setter and by
            // the load path (both normalise through
            // GanttChartSettings.IsCompatible), so
            // persisting the selection cannot write a pair the catalogue
            // reader refuses.
            ["PeriodLabelFormat"] = _state.PeriodLabelFormat.ToString(),
            ["Margin"] = _state.Margin.ToString(),
            ["MarginCm"] = _state.MarginCm.ToString("R", CultureInfo.InvariantCulture),
            ["SizePreset"] = _state.Preset?.Key ?? SizePresets.Default.Key,
        };

        ConfigWriteOutcome? outcome = _catalogueWriter?.WriteSettings(values);
        if (outcome is not null && !outcome.Succeeded)
        {
            string message = $"Failed to persist plot settings: {outcome.Refusal}";
            CommandBoundary.Instance.Run("PersistPlotSettings", () => throw new InvalidOperationException(message));
        }
    }

    /// <summary>
    /// Opens the theme settings dialog, applies edits through the catalogue
    /// writer, and returns whether themed edits were written. Never throws.
    /// </summary>
    internal bool ShowThemeSettingsDialog()
    {
        if (_catalogueReader is null || _catalogueWriter is null)
        {
            return false;
        }

        // CA1031: the dialog runs on the Excel main STA thread (the Ribbon
        // onAction thread). Any failure must not propagate into Excel.
#pragma warning disable CA1031
        try
        {
            // Read current themed values from the workbook
            ConfigReadOutcome outcome = _catalogueReader.Read();
            if (!outcome.Succeeded)
            {
                return false;
            }

            // Build current theme request from the workbook's themed styles.
            // Read the edited colours from outcome.Styles (the validated
            // GanttStyleRegistry) so the dialog shows the user's current theme,
            // not the code-owned defaults. Fall back to the preset values when
            // a preset is not present in the registry.
            List<ThemeStyleEdit> currentStyles = [];
            Dictionary<string, double> currentMetrics = [];
            List<CustomStyleDef> currentCustomDefs = [];

            HashSet<string> builtInKeys = new HashSet<string>(StringComparer.Ordinal);
            foreach (GanttStylePreset preset in GanttCatalogues.StylePresets)
            {
                _ = builtInKeys.Add(preset.StyleKey);

                string fill = preset.FillColour;
                string stroke = preset.StrokeColour;
                string text = preset.TextColour;
                GanttStrokeDash? dash = null;

                if (outcome.Styles.TryGet(preset.StyleKey, out GanttStyleDefinition? definition)
                    && definition is { HasFormatting: true })
                {
                    fill = definition.FillColour ?? preset.FillColour;
                    stroke = definition.StrokeColour ?? preset.StrokeColour;
                    text = definition.TextColour ?? preset.TextColour;

                    // The delineator's line type is the one built-in dash the
                    // user may change, so the reader projects it from the row
                    // rather than exact-matching it. The dialog must open on
                    // the line type the chart actually draws.
                    if (GanttStyleDashPolicy.IsUserEditable(preset.StyleKey))
                    {
                        dash = definition.StrokeDash ?? preset.StrokeDash;
                    }
                }

                currentStyles.Add(new ThemeStyleEdit(preset.StyleKey, fill, stroke, text, dash));
            }

            // Custom styles are registry entries that are not built-in presets.
            foreach (string key in outcome.Styles.StyleKeys)
            {
                if (builtInKeys.Contains(key))
                {
                    continue;
                }

                if (outcome.Styles.TryGet(key, out GanttStyleDefinition? definition)
                    && definition is { HasFormatting: true })
                {
                    // The user row's display name is projected by the reader;
                    // the key is the fallback when the row carries no name.
                    currentCustomDefs.Add(new CustomStyleDef(
                        key,
                        definition.DisplayName is { Length: > 0 } name ? name : key,
                        definition.FillColour ?? "",
                        definition.StrokeColour ?? "",
                        definition.TextColour ?? "",
                        definition.HatchPattern ?? GanttHatchPattern.None,
                        definition.StandardOutlinePt,
                        definition.ActivityHeightPt,
                        definition.MilestoneSizePt,
                        definition.DefaultLabelPosition ?? GanttLabelPosition.Auto,
                        definition.AllowedLabelPositions,
                        definition.StrokeDash ?? GanttStrokeDash.Solid));
                }
            }

            foreach (GanttMetricToken token in GanttCatalogues.Metrics)
            {
                if (ThemeEditRequest.EditableMetrics.Contains(token.Name)
                    && outcome.Metrics.TryGetValue(token.Name, out double parsed)
                    && double.IsFinite(parsed))
                {
                    currentMetrics[token.Name] = parsed;
                }
                else if (ThemeEditRequest.EditableMetrics.Contains(token.Name))
                {
                    currentMetrics[token.Name] = token.DefaultValue;
                }
            }

            // The chart-chrome colours: the workbook's current values, so the
            // dialog opens showing what the chart actually renders.
            Dictionary<string, string> currentChromeColours = [];
            foreach (GanttChartChromeRole role in GanttChartChrome.Roles)
            {
                string key = GanttChartChrome.SettingKeyFor(role);
                currentChromeColours[key] = outcome.Settings.TryGetValue(key, out string? stored)
                    && GanttColourToken.IsValidHex(stored)
                    ? stored
                    : GanttChartChrome.DefaultHexFor(role);
            }

            ThemeEditRequest request = new ThemeEditRequest(
                currentStyles,
                currentMetrics,
                null,
                currentCustomDefs,
                currentChromeColours);

            // Show the WinForms dialog. Virtual so tests can substitute a
            // non-blocking dialog result instead of running a message loop.
            using ThemeSettingsForm dialog = new ThemeSettingsForm(request);
            if (ShowThemeSettingsDialog(dialog) == DialogResult.OK)
            {
                ThemeEditRequest? validated = dialog.GetValidatedRequest();
                if (validated != null)
                {
                    // Write the themed edits
                    ConfigWriteOutcome? writeOutcome = _catalogueWriter.WriteTheme(validated);
                    if (writeOutcome is not null && writeOutcome.Succeeded)
                    {
                        return true;
                    }
                    else if (writeOutcome is not null)
                    {
                        _errorPresenter(
                            $"Failed to write theme settings: {writeOutcome.Refusal}");
                    }
                }
            }
        }
        catch
        {
            // Intentionally empty: a dialog failure must not escape into Excel.
        }
#pragma warning restore CA1031

        return false;
    }

    /// <summary>
    /// Shows the theme settings dialog modally and returns the user's result.
    /// Virtual so tests can substitute a non-blocking dialog result instead of
    /// running a WinForms message loop, which would block forever without a
    /// user to click OK or Cancel.
    /// </summary>
    protected virtual DialogResult ShowThemeSettingsDialog(ThemeSettingsForm dialog) => dialog.ShowDialog();

    /// <summary>
    /// Opens the chart-title dialog, applies edits through the catalogue
    /// writer, and returns whether title edits were written. Never throws.
    /// </summary>
    internal bool ShowChartTitleDialog()
    {
        if (_catalogueReader is null || _catalogueWriter is null)
        {
            return false;
        }

        // CA1031: the dialog runs on the Excel main STA thread (the Ribbon
        // onAction thread). Any failure must not propagate into Excel.
#pragma warning disable CA1031
        try
        {
            // Read current values from the workbook so the dialog shows the
            // user's actual title, not the code-owned default.
            ConfigReadOutcome outcome = _catalogueReader.Read();
            if (!outcome.Succeeded)
            {
                return false;
            }

            string currentTitle = outcome.Settings.TryGetValue("ChartTitle", out string? title)
                ? title
                : GanttCatalogues.SettingDefault("ChartTitle") ?? GanttWorkbookContract.DefaultChartTitle;
            bool currentShow = !outcome.Settings.TryGetValue("ShowTitle", out string? showText)
                || !string.Equals(showText, "FALSE", StringComparison.Ordinal);

            // Show the WinForms dialog
            using ChartTitleForm dialog = new ChartTitleForm(currentTitle, currentShow);
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                ChartTitleEditRequest? validated = dialog.GetValidatedRequest();
                if (validated != null)
                {
                    // Write exactly the two intended settings rows
                    ConfigWriteOutcome? writeOutcome = _catalogueWriter.WriteSettings(
                        ChartTitleEditRequest.ToSettings(validated));
                    if (writeOutcome is not null && writeOutcome.Succeeded)
                    {
                        return true;
                    }
                    else if (writeOutcome is not null)
                    {
                        _errorPresenter(
                            $"Failed to write chart title: {writeOutcome.Refusal}");
                    }
                }
            }
        }
        catch
        {
            // Intentionally empty: a dialog failure must not escape into Excel.
        }
#pragma warning restore CA1031

        return false;
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
                string startMode = settings.TryGetValue("PlotStartMode", out string? startModeVal) ? startModeVal : string.Empty;
                string finishMode = settings.TryGetValue("PlotFinishMode", out string? finishModeVal) ? finishModeVal : string.Empty;
                string startDate = NormaliseStoredPlotDate(
                    settings.TryGetValue("PlotStartDate", out string? startDateVal) ? startDateVal : string.Empty);
                string finishDate = NormaliseStoredPlotDate(
                    settings.TryGetValue("PlotFinishDate", out string? finishDateVal) ? finishDateVal : string.Empty);

                bool startAuto = PlotRangeModes.TryParse(startMode, out PlotRangeMode parsedStart)
                    && parsedStart == PlotRangeMode.DataRange;
                bool finishAuto = PlotRangeModes.TryParse(finishMode, out PlotRangeMode parsedFinish)
                    && parsedFinish == PlotRangeMode.DataRange;

                // R5.2 layout settings. Each is parsed through its closed Core
                // authority and falls back to the catalogue default on an
                // absent or unknown stored value, so a workbook that never set
                // a scale renders identically to a freshly initialised one.
                GanttTimeScale timeScale = ReadTimeScale(settings);
                GanttPlotMargin margin = ReadMargin(settings);
                double marginCm = ReadMarginCm(settings);
                SizePreset? preset = ReadPreset(settings);
                GanttPeriodLabelFormat periodLabelFormat = ReadPeriodLabelFormat(settings, timeScale);

                _state = _state with
                {
                    PlotStartAuto = startAuto,
                    PlotFinishAuto = finishAuto,
                    PlotStartDate = startDate,
                    PlotFinishDate = finishDate,
                    TimeScale = timeScale,
                    PeriodLabelFormat = periodLabelFormat,
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
    /// The one canonical period format for a scale: MMM for months, Qn for
    /// quarters, Wnn for weeks. Used when a stored format is absent, names an
    /// unknown value, or is incompatible with the scale it is stored against.
    /// </summary>
    private static GanttPeriodLabelFormat CanonicalPeriodLabelFormat(GanttTimeScale scale) => scale switch
    {
        GanttTimeScale.Month => GanttPeriodLabelFormat.MMM,
        GanttTimeScale.Quarter => GanttPeriodLabelFormat.Quarter,
        GanttTimeScale.Week => GanttPeriodLabelFormat.Week,
        _ => GanttPeriodLabelFormat.MMM,
    };

    /// <summary>
    /// Reads the stored period label format, falling back to the scale's
    /// canonical form when the key is absent, names an unknown format, or is
    /// incompatible with the scale it is stored against (R5.12). A workbook
    /// last saved on the quarter scale therefore renders the quarter's
    /// canonical form rather than an unreadable month label.
    /// </summary>
    private static GanttPeriodLabelFormat ReadPeriodLabelFormat(
        IReadOnlyDictionary<string, string> settings,
        GanttTimeScale scale) =>
        settings.TryGetValue("PeriodLabelFormat", out string? text)
        && GanttChartSettings.TryParsePeriodLabelFormat(text, out GanttPeriodLabelFormat parsed)
        && GanttChartSettings.IsCompatible(scale, parsed)
            ? parsed
            : CanonicalPeriodLabelFormat(scale);

    /// <summary>
    /// Reads the stored plot time scale, falling back to the catalogue default
    /// when the key is absent or names an unknown scale (R5.2).
    /// </summary>
    private static GanttTimeScale ReadTimeScale(IReadOnlyDictionary<string, string> settings) =>
        settings.TryGetValue("TimeScale", out string? text)
            && GanttChartSettings.TryParseTimeScale(text, out GanttTimeScale scale)
            ? scale
            : GanttTimeScale.Month;

    /// <summary>
    /// Reads the stored plot margin preset, falling back to the Windows
    /// "Normal" default when the key is absent or names an unknown preset
    /// (R5.2).
    /// </summary>
    private static GanttPlotMargin ReadMargin(IReadOnlyDictionary<string, string> settings) =>
        settings.TryGetValue("Margin", out string? text)
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
        settings.TryGetValue("MarginCm", out string? raw)
            && double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            && double.IsFinite(parsed)
            ? parsed
            : GanttPlotMargins.DefaultCustomCm;

    /// <summary>
    /// Reads the stored output size preset, returning null when the key is
    /// absent or names an unknown preset (R5.2). A null preset means the
    /// width and height displays show nothing rather than a guess.
    /// </summary>
    private static SizePreset? ReadPreset(IReadOnlyDictionary<string, string> settings) =>
        settings.TryGetValue("SizePreset", out string? text) ? SizePresets.ByKey(text) : null;

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
            string explicitStart = _state.PlotStartAuto ? string.Empty : _state.PlotStartDate;
            string explicitFinish = _state.PlotFinishAuto ? string.Empty : _state.PlotFinishDate;

            string effectiveStart = explicitStart;
            string effectiveFinish = explicitFinish;

            if (_state.PlotStartAuto || _state.PlotFinishAuto)
            {
                GanttTableReadOutcome? table = _tableReader?.Read();
                if (table is { Succeeded: true } && table.Rows is not null)
                {
                    GanttValidationOutcome validation = GanttRowValidator.Validate(table.Rows);
                    int padding = ReadRangePaddingDays(settings);
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
        return settings.TryGetValue("RangePaddingDays", out string? text)
            && int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)
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
        SubscribeToSelectionChanges();
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
        // Start from the last known state so a transient probe failure
        // cannot grey the Refresh gate: a fresh `false` would. The
        // probe is guarded for the same reason as the source probes —
        // this method never throws, and a COM failure here must not
        // escape that contract.
        bool sheetInitialised = _state.SheetInitialised;
        if (_tableReader is not null)
        {
#pragma warning disable CA1031
            try
            {
                GanttTableReadOutcome tableOutcome = _tableReader.Read();
                sheetInitialised = tableOutcome.Refusal != GanttTableReadRefusalReason.TableMissing;
            }
            catch
            {
                // Keep the previous value: a transient COM failure must
                // not grey the controls (same policy as the source probes).
            }
#pragma warning restore CA1031
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
    /// Returns how many items the description-label dropdown offers: the row's
    /// projected positions plus the leading <c>Default</c> item that clears the
    /// cell (R5.5 ruling 5 — blank resolves to the style default / Auto
    /// cascade, §22). A pure read of the cached snapshot. Never throws.
    /// </summary>
    internal int GetDescriptionLabelCount() =>
        _state.DescriptionLabelOptions is { Count: > 0 } options ? options.Count + 1 : 0;

    /// <summary>
    /// Returns the item ID at <paramref name="index"/>: <c>Default</c> for the
    /// leading clear item, otherwise the position's enum name. Index 0 is
    /// always Default because the clear item is always offered. A pure read of
    /// the cached snapshot; an out-of-range index yields an empty string, which
    /// Excel treats as an empty item. Never throws.
    /// </summary>
    /// <param name="index">The zero-based item index.</param>
    internal string GetDescriptionLabelItemID(int index)
    {
        return index == 0
            ? EntityTypeCatalog.DefaultLabelSelectionId
            : _state.DescriptionLabelOptions is { Count: > 0 } options
            && index >= 1
            && index <= options.Count
                ? options[index - 1].ToString()
                : string.Empty;
    }

    /// <summary>
    /// Returns the user-facing label at <paramref name="index"/> (R5.5 ruling
    /// 3: the enum name is never shown verbatim — <c>Inside</c> reads
    /// <c>Centre</c>, <c>None</c> reads <c>None (hide)</c>, and the splitter
    /// positions read as their meaning). A pure read of the cached snapshot; an
    /// out-of-range index yields an empty string. Never throws.
    /// </summary>
    /// <param name="index">The zero-based item index.</param>
    internal string GetDescriptionLabelItemLabel(int index) =>
        index == 0
            ? "Default"
            : _state.DescriptionLabelOptions is { Count: > 0 } options
            && index >= 1
            && index <= options.Count
                ? LabelPositionDisplayText(options[index - 1])
                : string.Empty;

    /// <summary>
    /// Returns the selected item index: 0 (<c>Default</c>) when the row's cell
    /// is blank, otherwise one past the position's place in the projected
    /// options. A position outside the projected set resolves to
    /// <c>Default</c> rather than an unmatched index, because such a cell could
    /// only exist before this control existed. A pure read of the cached
    /// snapshot. Never throws.
    /// </summary>
    internal int GetDescriptionLabelSelectedIndex()
    {
        if (_state.DescriptionLabelOptions is not { Count: > 0 } options)
        {
            return 0;
        }

        if (_state.SelectedDescriptionLabelPosition is not GanttLabelPosition selected)
        {
            return 0;
        }

        for (int index = 0; index < options.Count; index++)
        {
            if (options[index] == selected)
            {
                return index + 1;
            }
        }

        return 0;
    }

    /// <summary>
    /// Returns the display text for a label position. Isolated so the wording
    /// table (R5.5 ruling 3) has exactly one authority and the getters above
    /// stay pure reads.
    /// </summary>
    /// <param name="position">The position being displayed.</param>
    private static string LabelPositionDisplayText(GanttLabelPosition position) =>
        position switch
        {
            GanttLabelPosition.Auto => "Auto",
            GanttLabelPosition.Left => "Left",
            GanttLabelPosition.Right => "Right",
            GanttLabelPosition.Inside => "Centre",
            GanttLabelPosition.TopLeft => "Top left",
            GanttLabelPosition.TopRight => "Top right",
            GanttLabelPosition.BottomLeft => "Bottom left",
            GanttLabelPosition.BottomRight => "Bottom right",
            GanttLabelPosition.DataPanelLeft => "Data panel",
            GanttLabelPosition.PlotCentre => "Plot centre",
            GanttLabelPosition.Both => "Both",
            GanttLabelPosition.None => "None (hide)",
            _ => position.ToString(),
        };

    /// <summary>
    /// Applies one description-label selection to the selected row: the item ID is
    /// parsed through the catalogue (<c>Default</c> clears the cell, an unknown
    /// ID is refused — never defaulted), the command runs, and the ribbon
    /// refreshes so the dropdown reflects the written cell (R1.5 D3). Never
    /// throws.
    /// </summary>
    /// <param name="selectionId">The selected item ID, or null in tests.</param>
    internal void SetDescriptionLabelPosition(string? selectionId)
    {
        if (!EntityTypeCatalog.TryParseLabelSelection(selectionId, out GanttLabelPosition? position))
        {
            return;
        }

        SetDescriptionLabelCommand.RunForExcel(position);
        Refresh();
    }

    /// <summary>
    /// Constants shared between the state service and the ribbon handlers
    /// for the custom-style dropdown (R5.6a).
    /// </summary>
    private static class Constants
    {
        /// <summary>The leading clear item's ID.</summary>
        internal const string NoneItem = "None";

        /// <summary>The leading clear item's user-facing label.</summary>
        internal const string NoneLabel = "None";
    }

    /// <summary>
    /// Returns how many items the custom-style dropdown offers: a leading
    /// <c>None</c> clear item plus the user-created (non-built-in) style keys
    /// of the configuration registry in catalogue order (R5.6a). Always at
    /// least 1 so Excel never queries an out-of-range index. A pure read of
    /// the cached snapshot. Never throws.
    /// </summary>
    internal int GetCustomStyleCount() =>
        (_state.CustomStyleOptions is { Count: > 0 } options ? options.Count : 0) + 1;

    /// <summary>
    /// Returns the item ID at <paramref name="index"/>: <c>None</c> for the
    /// leading clear item (blank key clears the row's StyleKey cell),
    /// otherwise the style key itself. A pure read of the cached snapshot.
    /// Never throws.
    /// </summary>
    /// <param name="index">The zero-based item index.</param>
    internal string GetCustomStyleItemID(int index) =>
        index == 0 ? Constants.NoneItem :
        _state.CustomStyleOptions is { Count: > 0 } options
        && index >= 1
        && index < options.Count + 1
            ? options[index - 1]
            : string.Empty;

    /// <summary>
    /// Returns the user-facing label at <paramref name="index"/>: <c>None</c>
    /// for the clear item, otherwise the style key shown verbatim. A pure
    /// read of the cached snapshot. Never throws.
    /// </summary>
    /// <param name="index">The zero-based item index.</param>
    internal string GetCustomStyleItemLabel(int index) =>
        index == 0 ? Constants.NoneLabel : GetCustomStyleItemID(index);

    /// <summary>
    /// Returns the selected item index: 0 for <c>None</c> when the cell is
    /// blank, names a built-in preset, or no single style-key row is
    /// selected; otherwise the row's stored custom key's place in the
    /// projected options plus 1. Excel's dropDown requires a valid
    /// zero-based index, so -1 is never returned (matching the date-label
    /// and description-label dropdowns). A pure read of the cached
    /// snapshot. Never throws.
    /// </summary>
    internal int GetCustomStyleSelectedIndex()
    {
        if (_state.CustomStyleOptions is not { Count: > 0 } options
            || _state.SelectedCustomStyleKey is not string selected)
        {
            return 0;
        }

        for (int index = 0; index < options.Count; index++)
        {
            if (string.Equals(options[index], selected, StringComparison.Ordinal))
            {
                return index + 1;
            }
        }

        return 0;
    }

    /// <summary>
    /// Applies one custom-style selection to the selected row: <c>None</c>
    /// clears the row's StyleKey cell, otherwise the item ID IS the style
    /// key. The command runs and the ribbon refreshes so the dropdown
    /// reflects the written cell (R1.5 D3). Never throws.
    /// </summary>
    /// <param name="styleKey">The selected item ID (<c>None</c> clears), or null in tests.</param>
    internal void SetCustomStyleKey(string? styleKey)
    {
        if (string.IsNullOrWhiteSpace(styleKey) || styleKey == Constants.NoneItem)
        {
            styleKey = string.Empty;
        }

        SetStyleKeyCommand.RunForExcel(styleKey.Trim());
        Refresh();
    }

    /// <summary>
    /// Returns how many items the date-label dropdown offers: the
    /// selected row's allowed date-label choices, projected from its
    /// <c>Type</c> (R5.14 D1/D4 — the projected set already includes
    /// <c>Auto</c>, so there is no separate leading clear item). A
    /// pure read of the cached snapshot. Never throws.
    /// </summary>
    internal int GetDateLabelCount() =>
        _state.DateLabelPlacementOptions is { Count: > 0 } placementOptions
            ? placementOptions.Count
            : _state.DateLabelOptions is { Count: > 0 } options ? options.Count : 0;

    /// <summary>
    /// Returns the item ID at <paramref name="index"/>: the leading
    /// <c>Auto</c> choice carries the <c>Default</c> clear-ID (selecting
    /// it clears the cell to blank = Auto per R5.14 D3), and every other
    /// choice carries its enum name. A pure read of the cached snapshot;
    /// an out-of-range index yields an empty string, which Excel treats
    /// as an empty item. Never throws.
    /// </summary>
    /// <param name="index">The zero-based item index.</param>
    internal string GetDateLabelItemID(int index)
    {
        // Owner ruling 2026-10-14: a delineator row carries the
        // positionable vocabulary, whose items are the description
        // label's corners. Auto clears the cell, so its item id is
        // the clear marker, not the enum name
        // (TryParseDateLabelPlacementSelection maps Default to null).
        if (_state.DateLabelPlacementOptions is { Count: > 0 } placementOptions)
        {
            return index < 0 || index >= placementOptions.Count
                ? string.Empty
                : placementOptions[index] == GanttLabelPosition.Auto
                    ? EntityTypeCatalog.DefaultDateLabelSelectionId
                    : placementOptions[index].ToString();
        }

        if (_state.DateLabelOptions is not { Count: > 0 } options
            || index < 0 || index >= options.Count)
        {
            return string.Empty;
        }

        // Auto clears the cell, so its item id is the clear marker, not
        // the enum name (TryParseDateLabelSelection maps Default to null).
        return options[index] == GanttDateLabelPosition.Auto
            ? EntityTypeCatalog.DefaultDateLabelSelectionId
            : options[index].ToString();
    }

    /// <summary>
    /// Returns the user-facing label at <paramref name="index"/> (the
    /// enum name is never shown verbatim; <see cref="DateLabelDisplayText"/>
    /// is the single wording authority). A pure read of the cached
    /// snapshot; an out-of-range index yields an empty string. Never throws.
    /// </summary>
    /// <param name="index">The zero-based item index.</param>
    internal string GetDateLabelItemLabel(int index) =>
        _state.DateLabelPlacementOptions is { Count: > 0 } placementOptions
        && index >= 0
        && index < placementOptions.Count
            ? LabelPositionDisplayText(placementOptions[index])
            : _state.DateLabelOptions is { Count: > 0 } options
            && index >= 0
            && index < options.Count
                ? DateLabelDisplayText(options[index])
                : string.Empty;

    /// <summary>
    /// Returns the selected item index: <c>Auto</c> (index 0) when the
    /// row's cell is blank, otherwise the choice's place in the projected
    /// options. A blank cell reads as Auto (R5.14 D3), and an out-of-set
    /// position — only possible for a stale cell — also reads as Auto. A
    /// pure read of the cached snapshot. Never throws.
    /// </summary>
    internal int GetDateLabelSelectedIndex()
    {
        // Owner ruling 2026-10-14: the positionable vocabulary's
        // leading choice is Auto, exactly as the which-date vocabulary's
        // is, so a blank cell reads as index 0 in either.
        if (_state.DateLabelPlacementOptions is { Count: > 0 } placementOptions)
        {
            // Blank (null) resolves to Auto, and Auto is the leading choice.
            GanttLabelPosition selectedPlacement = _state.SelectedDateLabelPlacement ?? GanttLabelPosition.Auto;
            for (int index = 0; index < placementOptions.Count; index++)
            {
                if (placementOptions[index] == selectedPlacement)
                {
                    return index;
                }
            }

            return 0;
        }

        if (_state.DateLabelOptions is not { Count: > 0 } options)
        {
            return 0;
        }

        // Blank (null) resolves to Auto, and Auto is the leading choice.
        GanttDateLabelPosition selected = _state.SelectedDateLabelPosition ?? GanttDateLabelPosition.Auto;
        for (int index = 0; index < options.Count; index++)
        {
            if (options[index] == selected)
            {
                return index;
            }
        }

        return 0;
    }

    /// <summary>
    /// Returns the display text for a date-label choice. Isolated so the
    /// wording table has exactly one authority and the getters above stay
    /// pure reads.
    /// </summary>
    /// <param name="position">The choice being displayed.</param>
    private static string DateLabelDisplayText(GanttDateLabelPosition position) =>
        position switch
        {
            GanttDateLabelPosition.Auto => "Auto",
            GanttDateLabelPosition.Start => "Start",
            GanttDateLabelPosition.Finish => "Finish",
            GanttDateLabelPosition.None => "None",
            _ => position.ToString(),
        };

    /// <summary>
    /// Applies one date-label selection to the selected row. The
    /// dropdown carries one of two vocabularies (owner ruling
    /// 2026-10-14): for a delineator row the item is a position,
    /// parsed through the catalogue (<c>Default</c> clears the cell
    /// to blank = Auto, an unknown ID is refused — never defaulted)
    /// and the placement command runs; for every other dated type
    /// the item is a which-date choice and the date-label command
    /// runs. The ribbon then refreshes so the dropdown reflects the
    /// written cell (R1.5 D3). Never throws.
    /// </summary>
    /// <param name="selectionId">The selected item ID, or null in tests.</param>
    internal void SetDateLabelPosition(string? selectionId)
    {
        if (_state.DateLabelPlacementOptions is { Count: > 0 }
            && EntityTypeCatalog.TryParseDateLabelPlacementSelection(
                selectionId,
                out GanttLabelPosition? placement))
        {
            SetDateLabelPlacementCommand.RunForExcel(placement);
            Refresh();
            return;
        }

        if (!EntityTypeCatalog.TryParseDateLabelSelection(selectionId, out GanttDateLabelPosition? position))
        {
            return;
        }

        SetDateLabelCommand.RunForExcel(position);
        Refresh();
    }

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
            bool? hasActiveWorkbook = _applicationAdapter?.HasActiveWorkbook() ?? null;
            bool? logAvailable = _logAvailabilitySource?.Invoke() ?? null;

            // R5.5 ruling 6: the description-label dropdown is the first selection-gated
            // control, so its context is captured here rather than left at its
            // initial value. The capture is sticky like every other fact — a
            // transient failure keeps the previous context rather than greying
            // the dropdown mid-selection.
            DescriptionLabelContext descriptionLabelContext = CaptureDescriptionLabelContext(previous);
            DateLabelContext dateLabelContext = CaptureDateLabelContext(previous);
            DateLabelPlacementContext dateLabelPlacementContext = CaptureDateLabelPlacementContext(previous);
            CustomStyleContext customStyleContext = CaptureCustomStyleContext(previous);

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
                previous.PlotHeightPt,
                previous.PeriodLabelFormat,
                descriptionLabelContext.Available,
                descriptionLabelContext.Selected,
                descriptionLabelContext.Options,
                dateLabelContext.Available,
                dateLabelContext.Selected,
                dateLabelContext.Options,
                dateLabelPlacementContext.Available,
                dateLabelPlacementContext.Selected,
                dateLabelPlacementContext.Options,
                customStyleContext.Available,
                customStyleContext.SelectedKey,
                customStyleContext.Options));
        }
        catch
        {
            return previous;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Captures the description-label dropdown's context from the selection identifier, the
    /// table reader, and the catalogue reader (R5.5 ruling 6 + D1 projection).
    /// Never throws: an indeterminable selection keeps the previous context
    /// (snapshot stickiness), so a transient COM failure cannot grey the
    /// dropdown mid-selection.
    /// </summary>
    /// <param name="previous">The snapshot currently published.</param>
    /// <returns>The captured context, or the previous one when unavailable.</returns>
    private DescriptionLabelContext CaptureDescriptionLabelContext(RibbonState previous)
    {
        ISelectionIdentifier? identifier = _selectionIdentifier;
        IGanttTableReader? reader = _tableReader;
        if (identifier is null || reader is null)
        {
            return UnavailableDescriptionLabelContext(previous);
        }

        // CA1031: the capture runs inside Ribbon callbacks and Excel events;
        // a failure keeps the last known context instead of propagating.
#pragma warning disable CA1031
        try
        {
            SelectionFacts facts = identifier.Identify();
            if (!facts.IsSingleBodyRow || facts.BodyIndex is not int bodyIndex)
            {
                return UnavailableDescriptionLabelContext(previous);
            }

            GanttTableReadOutcome read = reader.Read();
            if (!read.Succeeded || bodyIndex < 1 || bodyIndex > read.Rows.Count)
            {
                return UnavailableDescriptionLabelContext(previous);
            }

            GanttRowDto row = read.Rows[bodyIndex - 1];
            if (!EntityTypeCatalog.TryParse(row.TypeText, out GanttEntityType type))
            {
                return UnavailableDescriptionLabelContext(previous);
            }

            // Custom Activity / Custom Milestone resolve their set from the
            // row's style (R2.7b/c; R5.6a: the catalogue's RequiresStyleKey
            // flag is the single authority); every other type reads its
            // catalogue definition. A failed catalogue read yields an empty
            // registry, which the projection turns into an empty option set —
            // the dropdown then offers only Default, and the command refuses
            // with StyleUnknown at execution.
            GanttStyleRegistry styles = GanttStyleRegistry.Empty;
            if (EntityTypeCatalog.GetDefinition(type) is { RequiresStyleKey: true }
                && _catalogueReader is not null)
            {
                ConfigReadOutcome catalogues = _catalogueReader.Read();
                if (catalogues.Succeeded)
                {
                    styles = catalogues.Styles;
                }
            }

            IReadOnlyList<GanttLabelPosition> options =
                EntityTypeCatalog.GetAllowedLabelPositions(type, row.StyleKey, styles);
            if (options.Count == 0)
            {
                return UnavailableDescriptionLabelContext(previous);
            }

            // The cell stores the position by name; blank resolves to the
            // style default / Auto cascade (§22), so it reads as null.
            GanttLabelPosition? selected = EntityTypeCatalog.TryParseLabelSelection(
                row.LabelPositionText,
                out GanttLabelPosition? parsed)
                ? parsed
                : null;

            return new DescriptionLabelContext(true, selected, options);
        }
        catch
        {
            return UnavailableDescriptionLabelContext(previous);
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// The context published when the description-label dropdown cannot be resolved: not
    /// available, carrying the previously published selection and options so a
    /// transient failure cannot blank a resolved dropdown.
    /// </summary>
    /// <param name="previous">The snapshot currently published.</param>
    /// <returns>The unavailable context.</returns>
    private static DescriptionLabelContext UnavailableDescriptionLabelContext(RibbonState previous) =>
        new(false, previous.SelectedDescriptionLabelPosition, previous.DescriptionLabelOptions);

    /// <summary>
    /// Captures the custom-style dropdown's context from the selection
    /// identifier, the table reader, and the catalogue reader (R5.6a ruling 6
    /// + the non-built-in projection). Never throws: an indeterminable
    /// selection keeps the previous context (snapshot stickiness), so a
    /// transient COM failure cannot grey the dropdown mid-selection.
    /// </summary>
    /// <param name="previous">The snapshot currently published.</param>
    /// <returns>The captured context, or the previous one when unavailable.</returns>
    private CustomStyleContext CaptureCustomStyleContext(RibbonState previous)
    {
        ISelectionIdentifier? identifier = _selectionIdentifier;
        IGanttTableReader? reader = _tableReader;
        if (identifier is null || reader is null || _catalogueReader is null)
        {
            return UnavailableCustomStyleContext(previous);
        }

        // CA1031: the capture runs inside Ribbon callbacks and Excel events;
        // a failure keeps the last known context instead of propagating.
#pragma warning disable CA1031
        try
        {
            SelectionFacts facts = identifier.Identify();
            if (!facts.IsSingleBodyRow || facts.BodyIndex is not int bodyIndex)
            {
                return UnavailableCustomStyleContext(previous);
            }

            GanttTableReadOutcome read = reader.Read();
            if (!read.Succeeded || bodyIndex < 1 || bodyIndex > read.Rows.Count)
            {
                return UnavailableCustomStyleContext(previous);
            }

            GanttRowDto row = read.Rows[bodyIndex - 1];
            if (!EntityTypeCatalog.TryParse(row.TypeText, out GanttEntityType type)
                || EntityTypeCatalog.GetDefinition(type) is not { RequiresStyleKey: true })
            {
                return UnavailableCustomStyleContext(previous);
            }

            // The dropdown offers user-created styles only: the registry's
            // keys minus the code-owned built-in presets (the same
            // StylePresets authority the Theme settings dialog subtracts).
            // A failed catalogue read yields no options — the dropdown
            // disables, and the command still refuses unknown keys at
            // execution.
            ConfigReadOutcome catalogues = _catalogueReader.Read();
            if (!catalogues.Succeeded || catalogues.Styles is null)
            {
                return UnavailableCustomStyleContext(previous);
            }

            HashSet<string> builtInKeys = new HashSet<string>(StringComparer.Ordinal);
            foreach (GanttStylePreset preset in GanttCatalogues.StylePresets)
            {
                _ = builtInKeys.Add(preset.StyleKey);
            }

            List<string> options = [];
            foreach (string key in catalogues.Styles.StyleKeys)
            {
                if (!builtInKeys.Contains(key))
                {
                    options.Add(key);
                }
            }

            // A built-in assignment reads as no match (null): the dropdown
            // never lists built-ins, so it must not select one.
            string? selectedKey = options.Contains(row.StyleKey ?? string.Empty, StringComparer.Ordinal)
                ? row.StyleKey
                : null;

            return new CustomStyleContext(true, selectedKey, options);
        }
        catch
        {
            return UnavailableCustomStyleContext(previous);
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// The context published when the custom-style dropdown cannot be
    /// resolved: not available, carrying the previously published selection
    /// and options so a transient failure cannot blank a resolved dropdown.
    /// </summary>
    /// <param name="previous">The snapshot currently published.</param>
    /// <returns>The unavailable context.</returns>
    private static CustomStyleContext UnavailableCustomStyleContext(RibbonState previous) =>
        new(false, previous.SelectedCustomStyleKey, previous.CustomStyleOptions);

    /// <summary>
    /// Captures the date-label dropdown's context from the selection
    /// identifier and the table reader (R5.14 ruling 6 + D1/D4
    /// projection). The capability is type-only (D4): a Custom
    /// Activity style carries no date mode, so — unlike the
    /// description-label capture — no catalogue or style registry is
    /// consulted. Owner ruling 2026-10-14: the delineator's date
    /// label is positionable, so its which-date set is empty and this
    /// context reads as unavailable for a delineator row — the
    /// positionable context carries that row instead. Never throws:
    /// an indeterminable selection keeps the previous context
    /// (snapshot stickiness), so a transient COM failure cannot grey
    /// the dropdown mid-selection.
    /// </summary>
    /// <param name="previous">The snapshot currently published.</param>
    /// <returns>The captured context, or the previous one when unavailable.</returns>
    private DateLabelContext CaptureDateLabelContext(RibbonState previous)
    {
        ISelectionIdentifier? identifier = _selectionIdentifier;
        IGanttTableReader? reader = _tableReader;
        if (identifier is null || reader is null)
        {
            return UnavailableDateLabelContext(previous);
        }

        // CA1031: the capture runs inside Ribbon callbacks and Excel events;
        // a failure keeps the last known context instead of propagating.
#pragma warning disable CA1031
        try
        {
            SelectionFacts facts = identifier.Identify();
            if (!facts.IsSingleBodyRow || facts.BodyIndex is not int bodyIndex)
            {
                return UnavailableDateLabelContext(previous);
            }

            GanttTableReadOutcome read = reader.Read();
            if (!read.Succeeded || bodyIndex < 1 || bodyIndex > read.Rows.Count)
            {
                return UnavailableDateLabelContext(previous);
            }

            GanttRowDto row = read.Rows[bodyIndex - 1];
            if (!EntityTypeCatalog.TryParse(row.TypeText, out GanttEntityType type))
            {
                return UnavailableDateLabelContext(previous);
            }

            IReadOnlyList<GanttDateLabelPosition> options =
                EntityTypeCatalog.GetAllowedDateLabelPositions(type);
            if (options.Count == 0)
            {
                return UnavailableDateLabelContext(previous);
            }

            // The cell stores the choice by name; blank resolves to Auto
            // (R5.14 D3), so it reads as null.
            GanttDateLabelPosition? selected = EntityTypeCatalog.TryParseDateLabelSelection(
                row.DateLabelPositionText,
                out GanttDateLabelPosition? parsed)
                ? parsed
                : null;

            return new DateLabelContext(true, selected, options);
        }
        catch
        {
            return UnavailableDateLabelContext(previous);
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// The context published when the date-label dropdown cannot be
    /// resolved: not available, carrying the previously published choice
    /// and options so a transient failure cannot blank a resolved
    /// dropdown.
    /// </summary>
    /// <param name="previous">The snapshot currently published.</param>
    /// <returns>The unavailable context.</returns>
    private static DateLabelContext UnavailableDateLabelContext(RibbonState previous) =>
        new(false, previous.SelectedDateLabelPosition, previous.DateLabelOptions);

    /// <summary>
    /// Captures the date-label dropdown's context in the
    /// positionable vocabulary (owner ruling 2026-10-14): the
    /// selection identifier and the table reader, like the which-date
    /// capture, but the option set is the description label's corner
    /// set and it is available only for the delineator, the one type
    /// whose date label is positionable. Never throws: an
    /// indeterminable selection keeps the previous context (snapshot
    /// stickiness), so a transient COM failure cannot grey the
    /// dropdown mid-selection.
    /// </summary>
    /// <param name="previous">The snapshot currently published.</param>
    /// <returns>The captured context, or the previous one when unavailable.</returns>
    private DateLabelPlacementContext CaptureDateLabelPlacementContext(RibbonState previous)
    {
        ISelectionIdentifier? identifier = _selectionIdentifier;
        IGanttTableReader? reader = _tableReader;
        if (identifier is null || reader is null)
        {
            return UnavailableDateLabelPlacementContext(previous);
        }

        // CA1031: the capture runs inside Ribbon callbacks and Excel events;
        // a failure keeps the last known context instead of propagating.
#pragma warning disable CA1031
        try
        {
            SelectionFacts facts = identifier.Identify();
            if (!facts.IsSingleBodyRow || facts.BodyIndex is not int bodyIndex)
            {
                return UnavailableDateLabelPlacementContext(previous);
            }

            GanttTableReadOutcome read = reader.Read();
            if (!read.Succeeded || bodyIndex < 1 || bodyIndex > read.Rows.Count)
            {
                return UnavailableDateLabelPlacementContext(previous);
            }

            GanttRowDto row = read.Rows[bodyIndex - 1];
            if (!EntityTypeCatalog.TryParse(row.TypeText, out GanttEntityType type)
                || type != GanttEntityType.Delineator)
            {
                return UnavailableDateLabelPlacementContext(previous);
            }

            IReadOnlyList<GanttLabelPosition> options =
                EntityTypeCatalog.GetAllowedDateLabelPlacements(type);
            if (options.Count == 0)
            {
                return UnavailableDateLabelPlacementContext(previous);
            }

            // The cell stores the corner by name; blank resolves to Auto
            // (the which-date precedent), so it reads as null.
            GanttLabelPosition? selected = EntityTypeCatalog.TryParseDateLabelPlacementSelection(
                row.DateLabelPositionText,
                out GanttLabelPosition? parsed)
                ? parsed
                : null;

            return new DateLabelPlacementContext(true, selected, options);
        }
        catch
        {
            return UnavailableDateLabelPlacementContext(previous);
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// The context published when the date-label dropdown's
    /// positionable vocabulary cannot be resolved: not available,
    /// carrying the previously published corner and options so a
    /// transient failure cannot blank a resolved dropdown.
    /// </summary>
    /// <param name="previous">The snapshot currently published.</param>
    /// <returns>The unavailable context.</returns>
    private static DateLabelPlacementContext UnavailableDateLabelPlacementContext(RibbonState previous) =>
        new(false, previous.SelectedDateLabelPlacement, previous.DateLabelPlacementOptions);

    /// <summary>
    /// Attaches the selection-change handler, replacing any previous
    /// subscription so a re-activation cannot double-subscribe.
    /// </summary>
    private void SubscribeToSelectionChanges()
    {
        DisposeSelectionSubscription();

        ISelectionIdentifier? identifier = _selectionIdentifier;
        if (identifier is null)
        {
            return;
        }

        // CA1031: activation runs inside the never-throwing AutoOpen; a failure
        // to subscribe degrades to command-driven invalidation only.
#pragma warning disable CA1031
        try
        {
            _selectionSubscription = identifier.SubscribeSelectionChanged(OnSelectionChanged);
        }
        catch
        {
            _selectionSubscription = null;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Handles one selection change: re-captures the snapshot and invalidates
    /// the description-label, date-label, and custom-style dropdowns. Only
    /// those three controls are invalidated, so moving the selection does not
    /// re-query every getter on the ribbon. Never throws.
    /// </summary>
    private void OnSelectionChanged()
    {
        // CA1031: runs inside an Excel event dispatch; a failure must not
        // escape into Excel.
#pragma warning disable CA1031
        try
        {
            _state = CaptureSnapshot(_state, _state.SheetInitialised);
            InvalidateControl(RibbonControlIds.DescriptionLabel);
            InvalidateControl(RibbonControlIds.DateLabel);
            InvalidateControl(RibbonControlIds.CustomStyle);
        }
        catch
        {
            // Intentionally empty: the snapshot is unchanged and the next
            // refresh re-reads the selection.
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Detaches the selection-change subscription exactly once. Never throws.
    /// </summary>
    private void DisposeSelectionSubscription()
    {
        IDisposable? subscription = _selectionSubscription;
        _selectionSubscription = null;
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
        DisposeSelectionSubscription();
        DisposeWorkbookStateSubscription();
        _ribbon = null;
        _applicationAdapter = null;
        _logAvailabilitySource = null;
        _selectionIdentifier = null;
        _catalogueWriter = null;
        _catalogueReader = null;
        _tableReader = null;
        _panelGridMeasurementPort = null;
        _state = RibbonState.Initial;
    }
}
