using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office;

/// <summary>
/// The live <see cref="ISceneBuildRequestFactory"/>: it turns validated events, the
/// effective settings map, the style registry, and the measured live panel grid into
/// the <see cref="SceneBuildRequest"/> the scene builder consumes (R4.8A D5).
/// </summary>
/// <remarks>
/// <para>
/// <b>Every scene input is decided here.</b> The size preset, the plot bounds, the
/// plot date range, the time scale, the row-height tokens, the composition profile,
/// and the frame/band and panel themes are all resolved in this one type. That is the
/// point of D5: a second place computing any of them would be a second authority,
/// and the two would disagree at some margin that presents as a layout bug.
/// </para>
/// <para>
/// <b>Plot bounds come from <see cref="PlotGeometryResolver"/> and nowhere else.</b>
/// This type does not perform the subtraction itself. That is enforced by
/// <c>PlotGeometryAuthorityTests</c>, which fails if the D2 subtraction appears in
/// any other production file — and this is the one production caller it permits.
/// </para>
/// <para>
/// <b>The composition profile is fixed to <see cref="SceneCompositionProfile.LiveExcel"/>
/// here, in code, not by a caller remembering.</b> A live refresh's destination is
/// the worksheet, whose own cells are the data panel; supplying one would draw a
/// replica over the user's rows. <c>SceneBuilder</c> refuses a live request carrying
/// a panel, so this is enforced rather than merely intended.
/// </para>
/// <para>
/// Pure over its inputs and free of Excel: it reads no COM proxy, so a test supplies
/// four values and asserts the exact request produced.
/// </para>
/// </remarks>
public sealed class ExcelSceneBuildRequestFactory(ITextMetrics? metrics = null) : ISceneBuildRequestFactory
{
    /// <summary>
    /// The columns the live panel measures, in panel order.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The live sheet's real cells <em>are</em> the data panel, so the measured grid
    /// exists to position the chart against them, not to reproduce them.
    /// </para>
    /// <para>
    /// <b>Derived from the schema's own visibility classification, never a
    /// hand-written list.</b> This list was the literal <c>["Id", "Type",
    /// "Description"]</c>, and <c>Id</c> is an <c>EngineHidden</c> column. Step 4 of
    /// the refresh hides it before this measurement runs, Excel reports a hidden
    /// column's <c>Range.Width</c> as <c>0</c>, <c>PanelCellGrid.TryCreate</c> refused
    /// that as a non-positive width, and every live refresh died at
    /// <c>MeasurementRefused</c> with "The worksheet columns could not be measured."
    /// The refusal was correct; the input was not. The
    /// <c>PanelGridMeasurementIntegrationTests</c> suite had already recorded the trap
    /// in a comment and worked around it by measuring
    /// <c>Columns.First(c =&gt; !c.IsHidden)</c>, which is why CI was green and the
    /// product was not.
    /// </para>
    /// <para>
    /// <b>Why every visible column, not a chosen few.</b> Entity guide section 3
    /// requires the panel's right edge to touch the plot's left edge without overlap
    /// or gap, and <see cref="PlotGeometryResolver"/> places the plot at
    /// <c>textPanelWidthPt + chrome</c>. Measuring only the label columns would
    /// therefore draw the plot on top of the still-visible <c>Start</c>,
    /// <c>Finish</c>, and <c>Duration</c> columns. The measured width is the table's
    /// real visible width, so it is derived from
    /// <see cref="GanttTableSchema.Default"/> rather than restated here.
    /// </para>
    /// <para>
    /// <b>This also absorbs a later change to which columns are hidden.</b> Because
    /// the set is computed from <see cref="GanttTableColumn.IsHidden"/> at type
    /// initialisation, reclassifying a column moves it in or out of the measured set
    /// with no edit to this file. Schema order is preserved, because
    /// <see cref="PanelCellGrid"/> treats column order as significant.
    /// </para>
    /// </remarks>
    private static IReadOnlyList<string> MeasuredColumnsCore =>
        [.. GanttTableSchema.Default.Columns.Where(column => !column.IsHidden).Select(column => column.Name)];

    /// <summary>The setting key naming the size preset.</summary>
    private const string _sizePresetKey = "SizePreset";

    /// <summary>
    /// The approved event-date display format setting key (ADR-0016).
    /// </summary>
    private const string _dateFormatKey = "DateDisplayFormat";

    /// <summary>The setting key naming the plot range padding, in days.</summary>
    private const string _rangePaddingKey = "RangePaddingDays";

    /// <summary>
    /// The default plot-range padding, in days, on each side.
    /// </summary>
    /// <remarks>
    /// It was <c>7</c>. The plot extent is now snapped to whole months (owner ruling,
    /// 2026-10-02), and a 7-day pad is wider than the gap the snap is meant to
    /// resolve: it would push a 10 Jan start back to 27 Dec, contradicting the stated
    /// requirement that a 10 Jan earliest date renders from 1 Jan. Three days is the
    /// owner's figure and is what makes 10 Jan land on 1 Jan while 3 Jan escapes to
    /// 1 Dec.
    /// </remarks>
    private const int _defaultRangePaddingDays = 3;

    // The metric tokens this factory resolves. They are named here once and
    // resolved through GanttCatalogues.MetricDefault rather than being written as
    // numeric literals, because a literal here is a second authority: the previous
    // version of this file carried eight of them and six disagreed with the
    // catalogue, so the rendered chart did not match the tokens the workbook
    // publishes. A token name that is not in the catalogue now throws at
    // construction rather than silently falling back (R4.8A D5).
    private const string _outerPaddingToken = "ChartOuterPaddingPt";
    private const string _titleBandToken = "TitleBandHeightPt";
    private const string _yearBandToken = "YearBandHeightPt";
    private const string _periodBandToken = "PeriodBandHeightPt";
    private const string _minHeaderLabelWidthToken = "MinimumHeaderLabelWidthPt";
    private const string _gridLineToken = "GridLinePt";
    private const string _majorBoundaryToken = "MajorBoundaryPt";
    private const string _delineatorLineToken = "DelineatorLinePt";
    private const string _stackGapToken = "StackGapPt";
    private const string _lanePaddingTopToken = "LanePaddingTopPt";
    private const string _lanePaddingBottomToken = "LanePaddingBottomPt";
    private const string _splitterHeightToken = "SplitterHeightPt";
    private const string _spacerHeightToken = "SpacerHeightPt";
    private const string _milestoneSizeToken = "MilestoneSizePt";
    private const string _labelGapToken = "LabelGapPt";
    private const string _rowHeightToken = "GanttRowHeightPt";

    /// <inheritdoc />
    public SceneBuildRequestOutcome Create(
        IReadOnlyList<GanttEvent> events,
        IReadOnlyDictionary<string, string> settings,
        GanttStyleRegistry registry,
        PanelCellGrid grid)
    {
        ArgumentNullException.ThrowIfNull(events);
        ArgumentNullException.ThrowIfNull(settings);
        ArgumentNullException.ThrowIfNull(registry);
        ArgumentNullException.ThrowIfNull(grid);

        // The range is derived from the events themselves, not from a stored date
        // range, so a refreshed chart always covers what it actually draws. A stored
        // range would be a second source of truth that survives the data changing.
        if (!TryResolveDateRange(events, settings, out DateOnly plotStart, out DateOnly plotFinish, out SceneBuildRequestRefusal? rangeRefusal, out var rangeMessage))
        {
            return SceneBuildRequestOutcome.Refused(rangeRefusal!.Value, rangeMessage!);
        }

        if (!TryResolvePreset(settings, out SizePreset? preset, out SceneBuildRequestRefusal? presetRefusal, out var presetMessage))
        {
            return SceneBuildRequestOutcome.Refused(presetRefusal!.Value, presetMessage!);
        }

        if (!TryParseScale(settings, out GanttTimeScale scale, out GanttPeriodLabelFormat periodFormat)
            || !TryParseDateFormat(settings, out GanttDateDisplayFormat dateFormat))
        {
            return SceneBuildRequestOutcome.Refused(
                SceneBuildRequestRefusal.InvalidSetting,
                "A configured chart setting could not be parsed into its typed form.");
        }

        PanelCellGrid measuredGrid = grid;

        // The grid's own total, not a re-sum of its columns. Re-deriving it here
        // would be a second computation of a figure the grid already publishes, and
        // the two could disagree on what "the panel width" means.
        var textPanelWidthPt = measuredGrid.TotalWidthPt;
        var chrome = GanttCatalogues.MetricDefault(_outerPaddingToken);

        // ADR-0031 D1/D2: the chart's frame margin is per side. Top and bottom are
        // the measured padding rows, so the chart's outer edge lands on a row
        // boundary the user can see; the left is ZERO so the plot sits flush against
        // the data table instead of one margin-width away from it.
        //
        // The left margin is not a tuning value. It is the gap between the table's
        // right edge and the plot's left edge, and a user reading the sheet sees a
        // visible channel of nothing between two things that are one object. The
        // right margin is untouched: the plot's right edge abuts the sheet edge and
        // the frame's right padding is what stops the last period label touching it.
        var padding = new ChartPaddingPt(
            LeftPt: 0,
            TopPt: measuredGrid.TopPaddingHeightPt,
            RightPt: chrome,
            BottomPt: measuredGrid.BottomPaddingHeightPt);

        // The header band heights are NOT read here any more. They used to be summed into
        // the plot's top offset; that sum was the page-coordinate origin ADR-0030
        // removes. The bands are now positioned by the scene from the reserved and
        // header rows the worksheet already carries (slice 1), so the factory has no
        // band arithmetic left to get wrong.

        // ADR-0030 D1/D2: the plot's vertical extent comes from the WORKSHEET, not from
        // the paper. `topPt` is the first body row's measured top, so lane 0 begins
        // exactly where that row does; `heightPt` is the measured total body height,
        // so the chart ends with the last row. Both were previously the preset's page
        // coordinates - a fixed 58pt offset and the remaining page height - which is
        // why a live bar sat ~2.4 rows below its own row and a 7-row chart ran ~400pt
        // past the table.
        //
        // The page no longer bounds the vertical extent (D7): a live sheet is as tall
        // as the user made it, and page-bounding it would refuse to render a 40-row
        // schedule because A4 landscape is 29 rows tall. The WIDTH budget is
        // untouched, so the plot still refuses a panel too wide to leave a readable
        // plot (R4.7H D4).
        var topPt = measuredGrid.OriginTopPt;
        var heightPt = measuredGrid.TotalRowHeightPt;

        // The single production call to the plot authority. PlotGeometryResolver owns
        // the subtraction and the bounds; this type supplies the measurements and
        // consumes the result.
        PlotGeometryOutcome geometry = PlotGeometryResolver.TryResolve(
            preset,
            textPanelWidthPt,

            // ADR-0031 D2: no left chrome. The plot begins exactly where the data
            // table ends, so the two read as one object. The RIGHT chrome is
            // unchanged and still earns its place: the plot's right edge is the
            // sheet's own edge, and without that margin the final period label would
            // sit flush against it.
            leftChromePt: 0,
            rightChromePt: chrome,
            topPt: topPt,
            heightPt: heightPt,
            boundVerticallyToPage: false);

        // Lane metrics are resolved from the MEASURED grid rather than from a
        // constant, because a lane's height is the row height the user can see
        // (ADR-0026 D2). Hardcoding one would build a chart that silently disagrees
        // with the worksheet beside it — the exact defect row-height normalisation
        // exists to prevent.
        LaneMetricsResolution lanes = LaneMetricsResolver.Resolve(
            measuredGrid,
            lanePaddingTopPt: GanttCatalogues.MetricDefault(_lanePaddingTopToken),
            lanePaddingBottomPt: GanttCatalogues.MetricDefault(_lanePaddingBottomToken),
            stackGapPt: GanttCatalogues.MetricDefault(_stackGapToken),
            splitterHeightPt: GanttCatalogues.MetricDefault(_splitterHeightToken),
            spacerHeightPt: GanttCatalogues.MetricDefault(_spacerHeightToken));
        if (lanes.Metrics is not { } laneMetrics)
        {
            return SceneBuildRequestOutcome.Refused(
                SceneBuildRequestRefusal.MeasurementRefused,
                "The measured row heights could not be resolved into lane geometry.");
        }

        if (metrics is null)
        {
            return SceneBuildRequestOutcome.Refused(
                SceneBuildRequestRefusal.MeasurementRefused,
                "No text-measurement service is available, so the chart cannot be composed.");
        }

        if (geometry.Succeeded)
        {
            SceneBuildRequest request = BuildRequest(
                events,
                registry,
                preset!,
                measuredGrid,
                geometry,
                plotStart,
                plotFinish,
                scale,
                periodFormat,
                dateFormat,
                settings,
                padding,
                metrics,
                laneMetrics);

            return SceneBuildRequestOutcome.Ok(request);
        }

        return SceneBuildRequestOutcome.Refused(
            SceneBuildRequestRefusal.PlotGeometryRefused,
            DescribeGeometryRefusal(geometry));
    }

    /// <summary>
    /// Assembles the request once every decision has been resolved.
    /// </summary>
    private static SceneBuildRequest BuildRequest(
        IReadOnlyList<GanttEvent> events,
        GanttStyleRegistry registry,
        SizePreset preset,
        PanelCellGrid measuredGrid,
        PlotGeometryOutcome geometry,
        DateOnly plotStart,
        DateOnly plotFinish,
        GanttTimeScale scale,
        GanttPeriodLabelFormat periodFormat,
        GanttDateDisplayFormat dateFormat,
        IReadOnlyDictionary<string, string> settings,
        ChartPaddingPt padding,
        ITextMetrics metrics,
        LaneLayoutMetrics laneMetrics) =>
        new()
        {
            Events = events,
            Registry = registry,
            Metrics = metrics,
            LaneMetrics = laneMetrics,
            FrameTheme = FrameTheme(),
            Preset = preset,
            Grid = measuredGrid,
            PlotBounds = geometry.Geometry!.PlotBounds,

            // LiveExcel, chosen in code (D4). The panel stays null because the
            // worksheet's own cells are the panel; SceneBuilder refuses a live
            // request that supplies one, so this cannot drift into drawing a replica
            // over the user's rows.
            Profile = SceneCompositionProfile.LiveExcel,
            Panel = null,

            PlotStart = plotStart,
            PlotFinish = plotFinish,
            Scale = scale,
            PeriodLabelFormat = periodFormat,
            DateFormat = dateFormat,
            Title = ReadString(settings, "ChartTitle"),
            AlternateBanding = ReadBool(settings, "AlternateBanding", fallback: true),
            ShowMinorGrid = ReadBool(settings, "ShowMinorGrid", fallback: true),
            ShowMajorGrid = ReadBool(settings, "ShowMajorGrid", fallback: true),
            YearBandHeightPt = GanttCatalogues.MetricDefault(_yearBandToken),
            PeriodBandHeightPt = GanttCatalogues.MetricDefault(_periodBandToken),
            TitleBandHeightPt = GanttCatalogues.MetricDefault(_titleBandToken),
            ChartPadding = padding,

            // ADR-0032 D2: the date-label text style. `SceneBuilder` skips the WHOLE
            // date-label pass when this is null, and it did skip it: the factory never
            // assigned it, so every activity bar rendered with no start or finish
            // date while a delineator label - which is built on a different path with
            // its own metrics - still appeared. That asymmetry is the exact shape of
            // the report ("only delineator labels generate"), and the skip is a bare
            // `return`, not a warning, so nothing in the scene recorded it.
            //
            // The style is the code-owned DefaultText token, exactly as the export
            // composition resolves its outside-label colour: the token table stays
            // the single authority for a colour rather than a literal here.
            //
            // The token is the label's TEXT colour, never its fill. It was written
            // as `fillColour: ColourHex.Parse("#000000")`, and `OfficeStyleMapper`
            // reports a TextBox as carrying a fill, so every description and date
            // label reached the host with an opaque black rectangle and default black
            // text on top of it. The comment above already said "DefaultText token",
            // so the intent was the text colour all along; only the named argument was
            // wrong. A label has no fill and no stroke - owner ruling - so both are
            // left null rather than defaulted to white.
            LabelStyle = new SceneStyle("DefaultText", textColour: ResolveDefaultTextColour()),
            MinimumHeaderLabelWidthPt = GanttCatalogues.MetricDefault(_minHeaderLabelWidthToken),
            GridLinePt = GanttCatalogues.MetricDefault(_gridLineToken),
            MajorBoundaryPt = GanttCatalogues.MetricDefault(_majorBoundaryToken),
            MilestoneSizePt = GanttCatalogues.MetricDefault(_milestoneSizeToken),
            DelineatorLinePt = GanttCatalogues.MetricDefault(_delineatorLineToken),
            DelineatorStackGapPt = GanttCatalogues.MetricDefault(_stackGapToken),
            LabelGapPt = GanttCatalogues.MetricDefault(_labelGapToken),
            RowHeightPt = GanttCatalogues.MetricDefault(_rowHeightToken),
        };

    /// <summary>
    /// Resolves the code-owned <c>DefaultText</c> colour token, or
    /// <see langword="null"/> when the catalogue does not publish it.
    /// </summary>
    /// <returns>The parsed token colour, or <see langword="null"/>.</returns>
    /// <remarks>
    /// <para>
    /// The token table is the single authority for the colour; this type must not
    /// restate <c>#000000</c> as a literal. A missing or unparseable token resolves
    /// to <see langword="null"/>, which reaches the host as "leave the font alone"
    /// rather than as a substituted colour.
    /// </para>
    /// <para>
    /// This mirrors the outside-label lookup <c>SceneBuilder</c> performs for
    /// section 17, so both paths resolve the same token the same way and cannot
    /// disagree about what "default text" means.
    /// </para>
    /// </remarks>
    private static ColourHex? ResolveDefaultTextColour() =>
        GanttCatalogues.Colours.FirstOrDefault(token => token.Name == "DefaultText")
            is { } defaultTextToken
            && ColourHex.TryParse(defaultTextToken.HexValue, out ColourHex? parsed)
                ? parsed
                : null;

    /// <summary>
    /// The frame and band styles, which the scene requires to be non-null on every
    /// member.
    /// </summary>
    /// <remarks>
    /// A single, neutral theme. The style-token resolution that will make these
    /// user-configurable is R5.6a's work; what matters here is that the scene builder
    /// refuses a null member rather than defaulting one, so a missing theme is a
    /// typed refusal at composition time instead of an unstyled chart.
    /// </remarks>
    /// <returns>The frame theme.</returns>
    private static FrameBandsTheme FrameTheme() =>
        new(
            Background: new SceneStyle("Background", fillColour: ColourHex.Parse("#FFFFFF")),
            AlternateBand: new SceneStyle("AlternateBand", fillColour: ColourHex.Parse("#F5F5F5")),
            MinorGrid: new SceneStyle("MinorGrid", strokeColour: ColourHex.Parse("#E0E0E0"), outlineWidthPt: 0.5),
            MajorGrid: new SceneStyle("MajorGrid", strokeColour: ColourHex.Parse("#BFBFBF"), outlineWidthPt: 1),
            YearHeader: new SceneStyle("YearHeader", fillColour: ColourHex.Parse("#D9D9D9"), textColour: ColourHex.Parse("#000000"), fontSizePt: 10, bold: true),
            PeriodHeader: new SceneStyle("PeriodHeader", fillColour: ColourHex.Parse("#E6E6E6"), textColour: ColourHex.Parse("#000000"), fontSizePt: 9),
            Title: new SceneStyle("Title", textColour: ColourHex.Parse("#000000"), fontSizePt: 12, bold: true));

    /// <summary>
    /// The columns this factory measures, for the caller to pass to
    /// <see cref="IPanelGridMeasurementPort"/>.
    /// </summary>
    public static IReadOnlyList<string> MeasuredColumns => MeasuredColumnsCore;

    /// <summary>
    /// Resolves the inclusive plot range from the events themselves, padded by the
    /// configured number of days.
    /// </summary>
    private static bool TryResolveDateRange(
        IReadOnlyList<GanttEvent> events,
        IReadOnlyDictionary<string, string> settings,
        out DateOnly plotStart,
        out DateOnly plotFinish,
        out SceneBuildRequestRefusal? refusal,
        out string? message)
    {
        plotStart = default;
        plotFinish = default;
        refusal = null;
        message = null;

        // Every entity contributes its Start, and ONLY a span-dated one contributes its
        // Finish. A milestone, a delineator, a Splitter and a Spacer read Start as
        // their single date per the entity guide's date-mode classification; letting
        // their Finish widen the range would be a claim the guide does not make. In
        // practice the validator clears Finish for those types, but this method
        // consumes GanttEvent values and must not depend on that having happened.
        var starts = new List<DateOnly>();
        var finishes = new List<DateOnly>();
        foreach (GanttEvent @event in events)
        {
            if (@event.Start is { } start)
            {
                starts.Add(start);
            }

            if (EntityTypeCatalog.GetDefinition(@event.Type)?.DateMode == EntityDateMode.StartFinish
                && @event.Finish is { } finish)
            {
                finishes.Add(finish);
            }
        }

        if (starts.Count == 0)
        {
            refusal = SceneBuildRequestRefusal.NoPlotRange;
            message = "No row carries a start date, so the chart has no date range to draw.";
            return false;
        }

        var padding = (int)ReadDouble(settings, _rangePaddingKey, 0, fallback: _defaultRangePaddingDays);
        if (padding < 0)
        {
            padding = 0;
        }

        DateOnly earliest = starts.Min();

        // The later of the two maxima, NOT the finish maximum alone. A single-date
        // event therefore always lands inside the range: taking finishes.Max()
        // whenever any finish existed could place the plot's right edge before a
        // milestone's own date, clipping the very event that set the range.
        DateOnly latest = starts.Max();
        if (finishes.Count > 0 && finishes.Max() > latest)
        {
            latest = finishes.Max();
        }

        // Month-snapped plot extent (owner ruling, 2026-10-02).
        //
        // The order is PADDING FIRST, THEN SNAP outward to the containing month, and
        // that order is load-bearing. Snapping first and then padding would place the
        // 10 Jan edge at 1 Jan minus the pad; padding first gives 10 Jan - 3 = 7 Jan,
        // which is still inside January, so it snaps back to 1 Jan. The 3-day pad is
        // therefore a tie-breaker for dates near a month edge, not a visible margin:
        // 10 Jan renders from 1 Jan, while 3 Jan pads to 31 Dec and snaps a whole
        // month further out to 1 Dec. Both were stated by the owner and only this
        // order satisfies both.
        plotStart = MonthStart(earliest.AddDays(-padding));
        plotFinish = MonthEnd(latest.AddDays(padding));

        // A one-day chart is degenerate: the scale builder has no interval to divide
        // and would either refuse or emit a single unreadable column. Widening to two
        // days is not a silent fudge — the alternative is a refusal on data that is
        // perfectly valid, and a user cannot act on "your chart is too narrow".
        if (plotFinish.DayNumber <= plotStart.DayNumber)
        {
            plotFinish = plotStart.AddDays(1);
        }

        return true;
    }

    /// <summary>
    /// Returns the first day of the month containing <paramref name="date"/>.
    /// </summary>
    /// <param name="date">The date whose month is wanted.</param>
    /// <returns>The first day of that month.</returns>
    /// <remarks>
    /// Constructed arithmetically rather than with <c>DateOnly.AddMonths</c> plus a
    /// day subtraction, which would overflow the <c>DateOnly.MinValue</c> range. The
    /// year and month are rebuilt from the date's own parts instead.
    /// </remarks>
    private static DateOnly MonthStart(DateOnly date) => new(date.Year, date.Month, 1);

    /// <summary>
    /// Returns the last day of the month containing <paramref name="date"/>.
    /// </summary>
    /// <param name="date">The date whose month is wanted.</param>
    /// <returns>The last day of that month.</returns>
    /// <remarks>
    /// Day zero of the FOLLOWING month is the last day of this one, which avoids
    /// both a hard-coded 28/30/31 table and the December overflow that
    /// <c>AddMonths(1)</c> would need a range check for.
    /// </remarks>
    private static DateOnly MonthEnd(DateOnly date) => MonthStart(date).AddMonths(1).AddDays(-1);

    /// <summary>
    /// Resolves the size preset, refusing an unknown key rather than defaulting.
    /// </summary>
    private static bool TryResolvePreset(
        IReadOnlyDictionary<string, string> settings,
        out SizePreset? preset,
        out SceneBuildRequestRefusal? refusal,
        out string? message)
    {
        preset = null;
        refusal = null;
        message = null;

        var key = ReadString(settings, _sizePresetKey) ?? SizePresets.Default.Key;
        preset = SizePresets.ByKey(key);
        if (preset is not null)
        {
            return true;
        }

        // NOT defaulted to A4. Substituting a preset the user did not choose would
        // render a chart at a size they never asked for, and the mismatch would only
        // show up when they printed it.
        refusal = SceneBuildRequestRefusal.UnknownSizePreset;
        message =
            "The configured size preset '"
            + key
            + "' is not one of: "
            + string.Join(", ", SizePresets.All.Select(static p => p.Key))
            + ".";
        return false;
    }

    private static bool TryParseScale(
        IReadOnlyDictionary<string, string> settings,
        out GanttTimeScale scale,
        out GanttPeriodLabelFormat periodFormat)
    {
        scale = GanttTimeScale.Month;
        periodFormat = GanttPeriodLabelFormat.MMM;

        var scaleText = ReadString(settings, "TimeScale");

        // IDE0046 would fold these two into one condition, and it is right that it
        // can: `A is not null && !B` is `A is null || B`. It is left unfolded because
        // the short-circuit is the point — each parse is skipped when the setting is
        // absent, and folding them hides which one is being talked about.
#pragma warning disable IDE0046
        if (scaleText is not null && !GanttChartSettings.TryParseTimeScale(scaleText, out scale))
        {
            return false;
        }

        var formatText = ReadString(settings, "PeriodLabelFormat");
        if (formatText is not null && !GanttChartSettings.TryParsePeriodLabelFormat(formatText, out periodFormat))
        {
            return false;
        }
#pragma warning restore IDE0046

        return GanttChartSettings.IsCompatible(scale, periodFormat);
    }

    private static bool TryParseDateFormat(
        IReadOnlyDictionary<string, string> settings,
        out GanttDateDisplayFormat dateFormat)
    {
        dateFormat = GanttDateDisplayFormat.DdMMyyyy;
        // The catalogue key is "DateDisplayFormat" (ADR-0016). This previously read
        // "DateFormat", which is not a key the settings table can ever contain, so the
        // lookup always missed and every chart fell back to DdMMyyyy — the approved
        // setting was unreachable in production while its test injected the key
        // directly and passed.
        var text = ReadString(settings, _dateFormatKey);

        // Same reason as TryParseScale: the short-circuit is the point.
#pragma warning disable IDE0046
        return text is null || GanttChartSettings.TryParseDateDisplayFormat(text, out dateFormat);
#pragma warning restore IDE0046
    }

    /// <summary>
    /// Explains a plot-geometry refusal in the user's terms, naming the shortfall
    /// where the resolver measured one.
    /// </summary>
    private static string DescribeGeometryRefusal(PlotGeometryOutcome outcome) =>
        outcome.Refusal == PlotGeometryRefusal.InsufficientPlotWidth
            ? "The data panel is "
                + Math.Round(outcome.Geometry!.TextPanelWidthPt)
                + "pt wide, which leaves no room for a readable plot on the chosen "
                + "paper size — it is "
                + Math.Round(outcome.Geometry.ShortfallPt)
                + "pt too wide. Narrow the label columns or choose a larger preset."
            : "The plot area could not be resolved: " + outcome.Refusal + ".";

    private static string? ReadString(IReadOnlyDictionary<string, string> settings, string key) =>
        settings.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value) ? value : null;

    private static double ReadDouble(
        IReadOnlyDictionary<string, string> settings,
        string key,
        double minimum,
        double fallback) =>
        settings.TryGetValue(key, out var raw)
        && double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed)
        && double.IsFinite(parsed)
        && parsed >= minimum
            ? parsed
            : fallback;

    private static bool ReadBool(IReadOnlyDictionary<string, string> settings, string key, bool fallback) =>
        settings.TryGetValue(key, out var raw)
        && bool.TryParse(raw, out var parsed)
            ? parsed
            : fallback;
}
