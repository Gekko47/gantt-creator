using System.Globalization;

namespace GanttCreator.Core.Scene;

/// <summary>The reason a scene could not be built.</summary>
public enum SceneBuilderRefusal
{
    /// <summary>The build request was null.</summary>
    NullRequest = 0,

    /// <summary>The request carried no renderable validated events.</summary>
    EmptyEvents = 1,

    /// <summary>The injected text-metrics seam was null.</summary>
    NullMetrics = 2,

    /// <summary>The caller-supplied measured panel grid was null.</summary>
    NullPanelGrid = 3,

    /// <summary>The caller-supplied panel and plot bounds were null.</summary>
    NullBounds = 4,

    /// <summary>The plot range could not be turned into a time scale.</summary>
    InvalidPlotRange = 5,

    /// <summary>The lane metrics or the resolved frame settings were refused.</summary>
    InvalidLayoutSettings = 6,

    /// <summary>One event's named style could not be resolved to a renderable style.</summary>
    UnresolvableStyle = 7,
}

/// <summary>
/// Every input <see cref="SceneBuilder"/> needs that Core cannot measure itself.
/// </summary>
/// <remarks>
/// Core is Office-free and offline, so measured geometry is a dependency rather
/// than a computation: a Phase 4/5 adapter reads the live Excel column widths and
/// supplies <see cref="Grid"/> and the bounds. <see cref="SceneBuilder"/> consumes
/// them and never re-measures or silently defaults a missing measurement.
/// </remarks>
public sealed record SceneBuildRequest
{
    /// <summary>Gets the validated events to render, in any order.</summary>
    public IReadOnlyList<GanttEvent> Events { get; init; } = [];

    /// <summary>Gets the named-style registry used to resolve each event's style.</summary>
    public GanttStyleRegistry Registry { get; init; } = GanttStyleRegistry.Empty;

    /// <summary>Gets the caller-supplied measured panel cell grid.</summary>
    public PanelCellGrid? Grid { get; init; }

    /// <summary>Gets the caller-supplied measured panel bounds.</summary>
    public RectD? PanelBounds { get; init; }

    /// <summary>Gets the caller-supplied measured plot bounds.</summary>
    public RectD? PlotBounds { get; init; }

    /// <summary>Gets the single injected text-metrics seam.</summary>
    public ITextMetrics? Metrics { get; init; }

    /// <summary>Gets the resolved lane layout metrics.</summary>
    public LaneLayoutMetrics? LaneMetrics { get; init; }

    /// <summary>Gets the resolved frame/band styles.</summary>
    public FrameBandsTheme? FrameTheme { get; init; }

    /// <summary>Gets the resolved data-panel styles, or null to omit the panel.</summary>
    public PanelTheme? Panel { get; init; }

    /// <summary>Gets the inclusive plot start date.</summary>
    public DateOnly PlotStart { get; init; }

    /// <summary>Gets the inclusive plot finish date.</summary>
    public DateOnly PlotFinish { get; init; }

    /// <summary>Gets the selected calendar scale.</summary>
    public GanttTimeScale Scale { get; init; } = GanttTimeScale.Month;

    /// <summary>Gets the selected period label format.</summary>
    public GanttPeriodLabelFormat PeriodLabelFormat { get; init; } = GanttPeriodLabelFormat.MMM;

    /// <summary>Gets the selected event-date display format.</summary>
    public GanttDateDisplayFormat DateFormat { get; init; } = GanttDateDisplayFormat.DdMMyyyy;

    /// <summary>Gets the chart title, or null when untitled.</summary>
    public string? Title { get; init; }

    /// <summary>Gets whether alternating plot bands are emitted.</summary>
    public bool AlternateBanding { get; init; } = true;

    /// <summary>Gets whether minor period grid lines are emitted.</summary>
    public bool ShowMinorGrid { get; init; } = true;

    /// <summary>Gets whether major year/plot grid lines are emitted.</summary>
    public bool ShowMajorGrid { get; init; } = true;

    /// <summary>Gets the year-header band height.</summary>
    public double YearBandHeightPt { get; init; }

    /// <summary>Gets the period-header band height.</summary>
    public double PeriodBandHeightPt { get; init; }

    /// <summary>Gets the title strip height.</summary>
    public double TitleBandHeightPt { get; init; }

    /// <summary>Gets the padding applied once around the panel/plot union.</summary>
    public double ChartOuterPaddingPt { get; init; }

    /// <summary>Gets the minimum visible period-label width.</summary>
    public double MinimumHeaderLabelWidthPt { get; init; }

    /// <summary>Gets the minor grid line width.</summary>
    public double GridLinePt { get; init; }

    /// <summary>Gets the major boundary/frame line width.</summary>
    public double MajorBoundaryPt { get; init; }

    /// <summary>Gets the milestone diamond tip-to-tip size.</summary>
    public double MilestoneSizePt { get; init; }

    /// <summary>Gets the critical-interval overlay line width.</summary>
    public double CriticalLinePt { get; init; }

    /// <summary>Gets the delineator line width, from the <c>DelineatorLinePt</c> token.</summary>
    public double DelineatorLinePt { get; init; }

    /// <summary>Gets the gap between a shape bound and an external label.</summary>
    public double LabelGapPt { get; init; }

    /// <summary>Gets the one-line label box height.</summary>
    public double LabelHeightPt { get; init; }

    /// <summary>
    /// Gets the maximum width of an external label, transcribed from the
    /// <c>MaximumExternalLabelWidthPt</c> metric token.
    /// </summary>
    /// <remarks>
    /// Entity guide section 22 makes this token the maximum external width, and
    /// <see cref="LabelPlanner"/> already honours whatever it is given. Passing the
    /// plot width instead made the cap depend on how wide the caller happened to
    /// draw the time axis rather than on the approved token, so a wide plot
    /// produced a label wider than the 36-360pt the catalogue allows.
    /// </remarks>
    /// <remarks>
    /// The default is the code-owned catalogue value, not zero. The planner accepts
    /// a zero maximum as valid, so a request that omitted the property would have
    /// silently suppressed every external label rather than failing; seeding the
    /// approved default keeps an incomplete request behaving like the catalogue.
    /// </remarks>
    public double MaximumExternalLabelWidthPt { get; init; } =
        GanttCatalogues.Metrics.First(token => token.Name == "MaximumExternalLabelWidthPt").DefaultValue;

    /// <summary>Gets the vertical gap between stacked same-date delineator labels.</summary>
    public double DelineatorStackGapPt { get; init; }

    /// <summary>Gets the resolved description-label text style, or null to omit them.</summary>
    public SceneStyle? LabelStyle { get; init; }

    /// <summary>Gets whether section 23 start and finish date labels are emitted.</summary>
    public bool ShowDateLabels { get; init; } = true;
}


/// <summary>The built scene plus the scale every date-to-X mapping came from.</summary>
/// <param name="Scene">The deterministically ordered scene.</param>
/// <param name="TimeScale">The time scale the geometry was derived from.</param>
public sealed record SceneBuildResult(GanttScene Scene, TimeScale TimeScale);

/// <summary>The typed result of attempting to build a scene.</summary>
/// <param name="Result">The successful result, or <see langword="null"/>.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/>.</param>
public sealed record SceneBuildOutcome(SceneBuildResult? Result, SceneBuilderRefusal? Refusal)
{
    /// <summary>Gets whether the build succeeded.</summary>
    public bool Succeeded => Result is not null;
}


/// <summary>
/// Composes the landed R3.4-R3.11 builders into the single entry point the
/// renderers consume. It derives no geometry of its own: every coordinate comes
/// from a builder that already owns its contract.
/// </summary>
/// <remarks>
/// <para>
/// The order is fixed by dependency, not preference. The time scale needs the
/// caller-supplied plot bounds; lane layout needs the events and their resolved
/// heights; <see cref="LaneEventLayout"/> binds each event to the slot centre lane
/// layout derived; the span builder produces each parent's post-plot-clip visible
/// bounds, which the critical overlay clips to and which must therefore exist
/// before overlays run.
/// </para>
/// <para>
/// Every builder refusal is mapped to a typed <see cref="SceneBuilderRefusal"/> and
/// no exception crosses this boundary, because a broken dependency must not read
/// as an ordinary suppression.
/// </para>
/// </remarks>
public static class SceneBuilder
{
    /// <summary>Attempts to build a complete scene from validated events and measured geometry.</summary>
    /// <param name="request">The typed build request.</param>
    /// <returns>A typed result or refusal.</returns>
    public static SceneBuildOutcome TryBuild(SceneBuildRequest? request)
    {
        if (request is null)
        {
            return Refused(SceneBuilderRefusal.NullRequest);
        }

        if (request.Events is null || request.Events.Count == 0 || request.Events.Any(@event => @event is null))
        {
            return Refused(SceneBuilderRefusal.EmptyEvents);
        }

        if (request.Metrics is null)
        {
            return Refused(SceneBuilderRefusal.NullMetrics);
        }

        if (request.Grid is null)
        {
            return Refused(SceneBuilderRefusal.NullPanelGrid);
        }

        if (request.PanelBounds is not { } panelBounds
            || request.PlotBounds is not { } plotBounds)
        {
            return Refused(SceneBuilderRefusal.NullBounds);
        }

        // There is deliberately no chart-bounds input. Entity guide section 2
        // defines ChartBounds as the union of the title, data panel, time headers,
        // and plot plus ChartOuterPaddingPt, so the frame builder derives it and the
        // scene is created with that derived value. A caller-supplied chart bounds
        // was a second, contradictory source: the scene could declare bounds the
        // chart:background primitive it contains did not have.

        TimeScaleCreationOutcome scale = TimeScale.TryCreate(
            request.PlotStart,
            request.PlotFinish,
            plotBounds.Left,
            plotBounds.Right);
        if (scale.Scale is not { } timeScale)
        {
            return Refused(SceneBuilderRefusal.InvalidPlotRange);
        }

        if (request.LaneMetrics is not { } laneMetrics)
        {
            return Refused(SceneBuilderRefusal.InvalidLayoutSettings);
        }

        // A hidden row validates but must not render, and a splitter or spacer is
        // lane geometry with no entity primitive. A delineator is excluded too: §24
        // makes it a full-height plot line, not a lane-bound entity, so giving it a
        // lane would reserve vertical space for a row it must not occupy. It is
        // built separately by the delineator pass.
        List<GanttEvent> renderable =
        [
            .. request.Events.Where(@event =>
                @event.Visible && @event.Type is not (GanttEntityType.Splitter
                    or GanttEntityType.Spacer
                    or GanttEntityType.Delineator)),
        ];
        if (renderable.Count == 0)
        {
            return Refused(SceneBuilderRefusal.EmptyEvents);
        }

        List<LaneEventInput> laneInputs = [];
        Dictionary<GanttRowId, ResolvedEventStyle> styles = [];

        // A delineator takes no lane, so its resolved line style is collected
        // separately and the grouping pass reads this map.
        List<GanttEvent> delineators =
        [
            .. request.Events.Where(@event =>
                @event.Visible && @event.Type == GanttEntityType.Delineator),
        ];
        Dictionary<GanttRowId, SceneStyle> delineatorStyles = [];
        foreach (GanttEvent @event in delineators)
        {
            // §24 draws the line from the resolved line style, but a Delineator has
            // no named-style default and no built-in preset, so an unresolvable one
            // is NOT a broken workbook: it falls back to the chart's own delineator
            // token style rather than refusing the whole scene. Only Types that
            // *have* a default are refused when it cannot resolve.
            delineatorStyles[@event.Id] = TryResolveStyle(request.Registry, @event, out ResolvedEventStyle? resolved) && resolved is not null
                ? resolved.Style
                : new SceneStyle("DefaultDelineator", strokeColour: ColourHex.Parse("#404040"));
        }

        foreach (GanttEvent @event in renderable)
        {
            // A Splitter, Spacer, or Delineator has no named-style default, so
            // requiring a resolvable style for one would refuse a valid workbook.
            // A Type that does have a default is still refused when it cannot
            // resolve, which is the R2.7c Custom Activity rule.
            ResolvedEventStyle? resolved = null;
            var requiresStyle = @event.Type is not (GanttEntityType.Splitter
                or GanttEntityType.Spacer
                or GanttEntityType.Delineator);
            if (requiresStyle)
            {
                if (!TryResolveStyle(request.Registry, @event, out resolved) || resolved is null)
                {
                    return Refused(SceneBuilderRefusal.UnresolvableStyle);
                }
            }

            // A style-less Type still occupies a lane, so it carries an explicit
            // zero-height marker style rather than being dropped from the map that
            // the milestone and overlay passes read.
            ResolvedEventStyle style = resolved ?? new ResolvedEventStyle(new SceneStyle("None"), 0);
            styles[@event.Id] = style;
            laneInputs.Add(new LaneEventInput(@event, style.HeightPt));
        }

        if (LaneLayoutBuilder.TryBuild(laneInputs, laneMetrics).Layout is not { } laneLayout)
        {
            return Refused(SceneBuilderRefusal.InvalidLayoutSettings);
        }

        if (LaneEventLayout.TryBuild(laneInputs, laneLayout).Result is not { } placements)
        {
            return Refused(SceneBuilderRefusal.InvalidLayoutSettings);
        }

        List<ScenePrimitive> primitives = [];
        List<SceneWarning> warnings = [.. laneLayout.Warnings, .. placements.Warnings];

        // The critical overlay clips to the parent's post-plot-clip visible span, so
        // the map is filled over the span events before any overlay is built.
        Dictionary<GanttRowId, RectD> parentVisibleBounds = [];
        // The lane layout is lane-relative (it starts at y=0 for the first lane), but
        // every bar, marker, and overlay is placed in chart coordinates. The plot
        // top is therefore added exactly once, here, so no builder re-adds it and
        // the offset cannot be applied twice.
        BuildSpans(placements, styles, timeScale, plotBounds, primitives, warnings, parentVisibleBounds);
        BuildOverlaysAndMilestones(request, placements, styles, timeScale, plotBounds, parentVisibleBounds, primitives, warnings);
        return BuildFramePanelAndScene(request, timeScale, panelBounds, plotBounds, placements, delineators, delineatorStyles, parentVisibleBounds, primitives, warnings);
    }


    /// <summary>One event's resolved scene style and its resolved height.</summary>
    /// <param name="Style">The resolved scene style.</param>
    /// <param name="HeightPt">The resolved activity height from the catalogue.</param>
    private sealed record ResolvedEventStyle(SceneStyle Style, double HeightPt);

    private static void BuildSpans(
        LaneEventLayoutResult placements,
        IReadOnlyDictionary<GanttRowId, ResolvedEventStyle> styles,
        TimeScale timeScale,
        RectD plotBounds,
        List<ScenePrimitive> primitives,
        List<SceneWarning> warnings,
        Dictionary<GanttRowId, RectD> parentVisibleBounds)
    {
        foreach (LaneEventPlacement placement in placements.Placements)
        {
            GanttEvent @event = placement.Event;
            if (@event.Type is not (GanttEntityType.AsPlannedActivity
                or GanttEntityType.AsBuiltActivity
                or GanttEntityType.BaselineActivity
                or GanttEntityType.CustomActivity
                or GanttEntityType.AsPlannedProcurement
                or GanttEntityType.AsBuiltProcurement
                or GanttEntityType.BaselineProcurement
                or GanttEntityType.DelayEvent))
            {
                continue;
            }

            ResolvedEventStyle resolved = styles[@event.Id];

            // Lane-relative to chart-relative, applied once per placement. The zero
            // width fallback below must use the same centre, or a parentless overlay
            // would clip against a band the bar never had.
            var slotCentreY = placement.SlotCentreY + plotBounds.Top;
            SpanBarCreationOutcome bar = SpanBarBuilder.TryBuild(
                new SpanBarRequest(
                    @event,
                    resolved.Style,
                    slotCentreY,
                    resolved.HeightPt,
                    placement.LaneOrder,
                    placement.EffectiveStackIndex),
                timeScale);
            if (bar.Result is not { } result)
            {
                // A span bar that cannot be placed is not the same as a span that
                // was clipped away: the first is a broken dependency and the second
                // is ordinary geometry. Silently continuing made the two
                // indistinguishable, so a row whose bar vanished reported nothing at
                // all. The other three placement passes (overlay, milestone,
                // delineator) already warn on a refusal, and this is the fourth.
                warnings.Add(new SceneWarning(
                    SceneOwnerId.ForRow(@event.Id),
                    "SpanBarRefused",
                    "The activity bar could not be placed."));
                continue;
            }

            warnings.AddRange(result.Warnings);
            if (result.Primitive is not { } primitive)
            {
                // A span wholly outside the plot emits no bar, but a critical
                // interval parented to it still needs something to clip against, so
                // the slot band is recorded with zero width. The overlay builder
                // clips horizontally against the time scale itself, so only the
                // vertical extent needs to be right.
                parentVisibleBounds[@event.Id] =
                    new RectD(
                        timeScale.PlotLeftPt,
                        slotCentreY - (resolved.HeightPt / 2),
                        0,
                        resolved.HeightPt);
                continue;
            }

            primitives.Add(primitive);
            parentVisibleBounds[@event.Id] = result.VisibleBounds ?? primitive.Bounds;
        }
    }


    private static void BuildOverlaysAndMilestones(
        SceneBuildRequest request,
        LaneEventLayoutResult placements,
        IReadOnlyDictionary<GanttRowId, ResolvedEventStyle> styles,
        TimeScale timeScale,
        RectD plotBounds,
        Dictionary<GanttRowId, RectD> parentVisibleBounds,
        List<ScenePrimitive> primitives,
        List<SceneWarning> warnings)
    {
        foreach (LaneEventPlacement placement in placements.Placements)
        {
            GanttEvent @event = placement.Event;
            ResolvedEventStyle resolved = styles[@event.Id];

            if (@event.Type is GanttEntityType.CriticalInterval)
            {
                // A critical interval whose parent emitted no visible span has
                // nothing to clip against, so the plot is the outer bound and the
                // builder refuses rather than drawing a full-width overlay.
                RectD parentBounds = @event.ParentId is { } parentId
                    && parentVisibleBounds.TryGetValue(parentId, out RectD found)
                        ? found
                        : plotBounds;

                CriticalOverlayCreationOutcome overlay = CriticalOverlayBuilder.TryBuild(
                    new CriticalOverlayRequest(
                        @event,
                        resolved.Style,
                        parentBounds,
                        request.CriticalLinePt,
                        placement.LaneOrder,
                        placement.EffectiveStackIndex),
                    timeScale,
                    parentVisibleBounds);
                if (overlay.Result is not { } overlayResult)
                {
                    warnings.Add(new SceneWarning(
                        SceneOwnerId.ForRow(@event.Id),
                        "CriticalOverlayRefused",
                        "The critical interval could not be overlaid on its parent."));
                    continue;
                }

                warnings.AddRange(overlayResult.Warnings);
                if (overlayResult.Primitive is { } overlayPrimitive)
                {
                    primitives.Add(overlayPrimitive);
                }

                continue;
            }

            if (@event.Type is not (GanttEntityType.AsPlannedMilestone
                or GanttEntityType.AsBuiltMilestone
                or GanttEntityType.BaselineMilestone
                or GanttEntityType.CriticalMilestone))
            {
                continue;
            }

            MilestoneMarkerCreationOutcome marker = MilestoneMarkerBuilder.TryBuild(
                new MilestoneMarkerRequest(
                    @event,
                    resolved.Style,
                    request.MilestoneSizePt,
                    // The same single lane-relative to chart-relative conversion the
                    // span pass applies, so a marker and a bar on one row cannot
                    // disagree about where that row is.
                    placement.SlotCentreY + plotBounds.Top,
                    placement.LaneOrder,
                    placement.EffectiveStackIndex),
                timeScale);
            if (marker.Result is not { } markerResult)
            {
                warnings.Add(new SceneWarning(
                    SceneOwnerId.ForRow(@event.Id),
                    "MilestoneRefused",
                    "The milestone marker could not be placed."));
                continue;
            }

            warnings.AddRange(markerResult.Warnings);
            if (markerResult.Primitive is { } markerPrimitive)
            {
                primitives.Add(markerPrimitive);

                // The marker's own bounds are what its description label anchors to,
                // so they must be recorded exactly as a span's are. Without this the
                // label pass finds no entry for the row and skips the milestone's
                // description entirely -- §22 covers milestone description labels, and
                // §21's §22-priority rank for them is unreachable if they are never
                // planned. The same map is read by the critical-overlay pass, where a
                // milestone is never a parent, so this entry cannot affect clipping.
                parentVisibleBounds[@event.Id] = EnclosingBounds(markerPrimitive.Points);
            }
        }
    }


    private static SceneBuildOutcome BuildFramePanelAndScene(
        SceneBuildRequest request,
        TimeScale timeScale,
        RectD panelBounds,
        RectD plotBounds,
        LaneEventLayoutResult placements,
        List<GanttEvent> delineators,
        Dictionary<GanttRowId, SceneStyle> delineatorStyles,
        Dictionary<GanttRowId, RectD> parentVisibleBounds,
        List<ScenePrimitive> primitives,
        List<SceneWarning> warnings)
    {
        if (request.FrameTheme is not { } frameTheme)
        {
            return Refused(SceneBuilderRefusal.InvalidLayoutSettings);
        }

        // The title band is shown only when a real title exists, so a blank title
        // never becomes a visible empty strip.
        var showTitle = !string.IsNullOrWhiteSpace(request.Title);

        FrameBandsCreationOutcome frame = FrameBandsBuilder.TryBuild(
            new FrameBandsRequest(
                timeScale,
                request.Scale,
                request.PeriodLabelFormat,
                panelBounds,
                plotBounds,
                request.ChartOuterPaddingPt,
                request.TitleBandHeightPt,
                request.YearBandHeightPt,
                request.PeriodBandHeightPt,
                request.MinimumHeaderLabelWidthPt,
                request.GridLinePt,
                request.MajorBoundaryPt,
                showTitle,
                request.Title ?? string.Empty,
                request.AlternateBanding,
                request.ShowMinorGrid,
                request.ShowMajorGrid,
                frameTheme),
            new TextWidthMeasurer(request.Metrics!));
        if (frame.Result is not { } frameResult)
        {
            return Refused(SceneBuilderRefusal.InvalidLayoutSettings);
        }

        primitives.AddRange(frameResult.Primitives);
        warnings.AddRange(frameResult.Warnings);

        // The delineator and label passes run here, not earlier, because §24 and the
        // §22/§23 label planners all bound their boxes by the *derived* chart bounds
        // the frame builder produces. Running them earlier would mean passing an
        // unverified chart rectangle back in, which is the R3.15 defect all over again.
        BuildDelineators(
            request,
            timeScale,
            plotBounds,
            frameResult.Geometry.ChartBounds,
            delineators,
            delineatorStyles,
            primitives,
            warnings);

        BuildLabels(
            request,
            timeScale,
            plotBounds,
            frameResult.Geometry.ChartBounds,
            placements,
            parentVisibleBounds,
            primitives,
            warnings);

        if (request.Panel is { } panelTheme)
        {
            // Section 4 fixes the header band's bottom edge to the period header's
            // bottom, which R3.5 derives as PlotBounds.Y - YearBandHeightPt. The
            // panel builder cannot know the plot bounds, so it is supplied here and
            // SceneBuilderTests asserts the emitted bottom equals this value.
            //
            // The header band height is the §4 panel header's own height, not the
            // year band: YearBandHeightPt belongs to §5, and reusing it here made a
            // structural equality (the bottoms align) rest on an unrelated pairing.
            // The height is derived from the panel's own measured row height so the
            // header and the body rows share one metric.
            PanelBuildOutcome panel = PanelBuilder.TryBuild(
                new PanelBuildRequest(
                    request.Grid!,
                    [.. placements.Placements.Select(placement => new PanelRow(placement.Event.Id, Cells(placement.Event, request)))],
                    plotBounds,
                    plotBounds.Y - request.YearBandHeightPt,
                    request.Grid!.RowHeightPt,
                    panelTheme));

            // A panel refusal must not be dropped: an empty panel would read as a
            // scene with no data table, which is a silent data loss rather than an
            // error. Every PanelBuildRefusal is precondition-checked upstream (the
            // frame builder refuses a non-positive plot, and the grid guarantees a
            // positive row height and a unique row per placement), so this branch is
            // a defence against a future PanelBuilder refusal. It reports the existing
            // InvalidLayoutSettings rather than adding a member that could never be
            // positively tested -- AGENTS.md treats an unreachable validator as a
            // defect, exactly as R3.15 D2 removed the dead PlotOutsideChart guard.
            if (panel.Result is not { } panelResult)
            {
                return Refused(SceneBuilderRefusal.InvalidLayoutSettings);
            }

            primitives.AddRange(panelResult.Primitives);
        }

        SceneCreationOutcome scene = GanttScene.TryCreate(frameResult.Geometry.ChartBounds, plotBounds, primitives, warnings);
        return scene.Scene is { } built
            ? new SceneBuildOutcome(new SceneBuildResult(built, timeScale), null)
            : Refused(SceneBuilderRefusal.InvalidLayoutSettings);
    }

    /// <summary>
    /// Projects one event's cell texts in grid-column order, as §3 requires the
    /// emitted bounds to follow the measured column order.
    /// </summary>
    /// <param name="event">The validated event supplying the cell values.</param>
    /// <param name="request">The build request supplying the grid and the date format.</param>
    /// <returns>
    /// One entry per grid column. A column with no schema mapping, or a field the
    /// entity type does not use, is <see langword="null"/>: blank is legal cell
    /// data (R2.5 U2) and the builder already omits text for it.
    /// </returns>
    /// <remarks>
    /// Dates go through <see cref="GanttDateFormatting"/> with the request's
    /// approved format, so a panel cell and a §23 date label cannot disagree and
    /// no host culture can re-derive the pattern.
    /// </remarks>
    private static List<string?> Cells(GanttEvent @event, SceneBuildRequest request)
    {
        List<string?> cells = new(request.Grid!.Columns.Count);
        foreach (PanelColumn column in request.Grid.Columns)
        {
            cells.Add(column.Name switch
            {
                // Every schema column is mapped explicitly rather than relying on a
                // default: a panel cell must reproduce the worksheet value it stands
                // for, and a silently null column renders as a blank cell that reads
                // as an empty worksheet cell.
                "Id" => @event.Id.Value,
                "Type" => EntityTypeCatalog.GetDefinition(@event.Type)?.DisplayName,
                "Description" => @event.Description,
                "Start" => @event.Start is { } start ? GanttDateFormatting.Format(start, request.DateFormat) : null,
                "Finish" => @event.Finish is { } finish ? GanttDateFormatting.Format(finish, request.DateFormat) : null,
                "LaneId" => @event.LaneId?.Value,
                "StackIndex" => @event.StackIndex?.ToString(CultureInfo.InvariantCulture),
                "ParentId" => @event.ParentId?.Value,
                "StyleKey" => @event.StyleKey,
                // The worksheet holds the enum member name, which is also what
                // GanttRowValidator parses back, so the cell round-trips.
                "LabelPosition" => @event.LabelPosition?.ToString(),
                "FillColour" => @event.FillColour,
                "StrokeColour" => @event.StrokeColour,
                // ExcelCellConverter.ToText renders a bool as TRUE/FALSE, so the cell
                // matches the value the table reader would parse back.
                "Visible" => @event.Visible ? "TRUE" : "FALSE",
                "SortOrder" => @event.SortOrder?.ToString(CultureInfo.InvariantCulture),
                _ => null,
            });
        }

        return cells;
    }

    /// <summary>
    /// Groups the delineators §24 requires and builds each group's line and labels.
    /// </summary>
    /// <remarks>
    /// A group is one date plus one resolved line style, because §24 draws the line
    /// once per resolved line style: two same-date rows sharing a style share a line,
    /// and two sharing a date but not a style get one line each. Grouping by date
    /// alone would make <c>DelineatorLayout</c> refuse the mixed group, and grouping
    /// by style alone would emit several lines for one date.
    /// </remarks>
    private static void BuildDelineators(
        SceneBuildRequest request,
        TimeScale timeScale,
        RectD plotBounds,
        RectD chartBounds,
        List<GanttEvent> delineators,
        Dictionary<GanttRowId, SceneStyle> delineatorStyles,
        List<ScenePrimitive> primitives,
        List<SceneWarning> warnings)
    {
        // Grouping is by (date, style), and the members inside a group are ordered by
        // the stable row ID. The input is not ordered -- the R3.12 determinism contract
        // is that a shuffled input produces a byte-identical scene -- so using the
        // first member in *input* order would make the group's owning row depend on
        // the caller's row order. Grouping preserves first-appearance order, so an
        // explicit order-by is required for determinism, not just tidiness.
        foreach (IGrouping<(DateOnly Date, string Style), GanttEvent> group in delineators
            .OrderBy(@event => @event.Id.Value, StringComparer.Ordinal)
            .GroupBy(@event => (@event.Start!.Value, StyleKey(delineatorStyles, @event))))
        {
            DelineatorRequest[] requests =
            [
                .. group.Select(@event => new DelineatorRequest(
                    @event,
                    delineatorStyles[@event.Id],
                    request.DelineatorLinePt,
                    plotBounds,
                    chartBounds,
                    request.LabelGapPt,
                    request.Metrics!,
                    @event.LabelPosition ?? GanttLabelPosition.Auto)),
            ];

            DelineatorGroupCreationOutcome groupOutcome = DelineatorLayout.TryBuildGroup(
                new DelineatorGroupRequest(requests, request.DelineatorStackGapPt),
                timeScale);
            if (groupOutcome.Result is not { } groupResult)
            {
                // The warning is owned by the group's first row, deterministically, so
                // a refused group is reported once rather than per member.
                warnings.Add(new SceneWarning(
                    SceneOwnerId.ForRow(group.First().Id),
                    "DelineatorGroupRefused",
                    "A same-date delineator group could not be laid out."));
                continue;
            }

            primitives.AddRange(groupResult.Primitives);
            warnings.AddRange(groupResult.Warnings);
        }
    }

    /// <summary>
    /// Plans the §22 description label and the §23 start/finish date labels, in that
    /// order, so a date label is planned against the description label already placed
    /// for its row.
    /// </summary>
    /// <remarks>
    /// <para>
    /// §23 requires a date label to avoid the row's own description label, and §22
    /// requires labels to avoid each other. Both are satisfied by one growing
    /// <c>occupants</c> list that every placed label joins, which is also the
    /// contract <see cref="DateLabelRequest.Occupants"/> documents. A label is
    /// registered as an occupant only once it is actually emitted, so a suppressed
    /// label reserves nothing.
    /// </para>
    /// <para>
    /// Rows are visited in §22 placement-priority order, not placement order: the
    /// guide makes a critical milestone's label outrank a planned activity's, so the
    /// higher-priority label must be offered its box first. Within one priority the
    /// placement order already encodes lane, stack, subtype, sort order, and stable
    /// ID, so a stable sort by priority alone preserves that tie-break.
    /// </para>
    /// </remarks>
    private static void BuildLabels(
        SceneBuildRequest request,
        TimeScale timeScale,
        RectD plotBounds,
        RectD chartBounds,
        LaneEventLayoutResult placements,
        Dictionary<GanttRowId, RectD> parentVisibleBounds,
        List<ScenePrimitive> primitives,
        List<SceneWarning> warnings)
    {
        if (request.LabelStyle is not { } labelStyle)
        {
            return;
        }

        LabelMetrics metrics = new(
            plotBounds,
            chartBounds,
            request.LabelGapPt,
            request.LabelHeightPt,
            request.MaximumExternalLabelWidthPt);

        List<RectD> occupants = [];

        // OrderBy is a stable sort, so the placement order the lane layout already
        // fixed (lane, stack, subtype, sort order, stable ID) survives within one
        // §22 priority and the scene stays byte-identical across refreshes.
        foreach (LaneEventPlacement placement in placements.Placements
            .OrderBy(placement => LabelPriorityFor(placement.Event.Type)))
        {
            GanttEvent @event = placement.Event;
            if (!parentVisibleBounds.TryGetValue(@event.Id, out RectD shapeBounds))
            {
                // No emitted bar or marker: there is nothing to label. A span wholly
                // outside the plot still needs no label, because its date is not shown
                // on the shape either.
                continue;
            }

            // Only occupants sharing a vertical band with this row constrain it. A
            // label in another lane cannot overlap this one -- the planner's own
            // Blocked test is a rectangle intersection, so it already ignores them --
            // but the free-space measure is one-dimensional, and without this filter a
            // label far to the right of another lane's bar would shrink this row's
            // measured gap and truncate or suppress a label that has room.
            List<RectD> relevant = VerticalBand(occupants, shapeBounds);

            LabelPlanCreationOutcome description = LabelPlanner.TryPlan(
                new LabelRequest(
                    @event,
                    @event.Description,
                    @event.LabelPosition ?? GanttLabelPosition.Auto,
                    shapeBounds,
                    labelStyle,
                    request.Metrics!,
                    LaneOrder: placement.LaneOrder,
                    StackIndex: placement.EffectiveStackIndex),
                metrics,
                relevant);
            if (description.Result is { } described)
            {
                warnings.AddRange(described.Warnings);
                if (described.Primitive is { } descriptionText && described.Bounds is { } descriptionBounds)
                {
                    primitives.Add(descriptionText);
                    occupants.Add(descriptionBounds);
                }
            }

            if (!request.ShowDateLabels)
            {
                continue;
            }

            // The full span is the event's own unclipped start-to-finish geometry, so
            // the builder can tell a bar the plot cut from a whole one. Deriving it
            // from the dates rather than the plot width is what makes that comparison
            // mean "clipped" instead of "narrower than the plot".
            // §23 anchors a date label beside the bar, so it must see the description
            // label just placed for this same row. The band is therefore recomputed
            // after that placement rather than reused: a band captured before the
            // description was added would not contain it, and the date label would be
            // planned as if the row had no description at all -- which is exactly the
            // collision §23 requires the occupants list to prevent.
            DateLabelOutcome dates = DateLabelBuilder.TryBuild(
                new DateLabelRequest(
                    @event,
                    shapeBounds,
                    FullBoundsOf(@event, timeScale, shapeBounds),
                    metrics,
                    request.Metrics!,
                    request.DateFormat,
                    labelStyle,
                    Occupants: VerticalBand(occupants, shapeBounds)));
            if (dates.Result is not { } planned)
            {
                // A refused date label is a broken dependency, not a placement
                // decision: the overlay, milestone, and span passes all warn rather
                // than continuing, and dropping it silently here meant a row could
                // lose both date labels with nothing in the scene recording why.
                warnings.Add(new SceneWarning(
                    SceneOwnerId.ForRow(@event.Id),
                    "DateLabelRefused",
                    "The start or finish date label could not be placed."));
                continue;
            }

            primitives.AddRange(planned.Primitives);
            foreach (SceneText dateLabel in planned.Primitives)
            {
                occupants.Add(dateLabel.TextBounds);
            }
        }
    }

    /// <summary>
    /// The axis-aligned rectangle that encloses a marker's points, used as the shape
    /// bounds a milestone's description label anchors to.
    /// </summary>
    /// <param name="points">The polygon's vertices; a milestone diamond is axis-aligned.</param>
    /// <returns>The enclosing rectangle.</returns>
    /// <remarks>
    /// A milestone is a diamond, so the enclosing rectangle is the diamond's own
    /// tip-to-tip box. The point list is non-empty by <see cref="ScenePolygon"/>'s
    /// own construction guard, so the seed needs no fallback.
    /// </remarks>
    private static RectD EnclosingBounds(IReadOnlyList<PointD> points)
    {
        var left = points[0].X;
        var top = points[0].Y;
        var right = left;
        var bottom = top;

        for (var i = 1; i < points.Count; i++)
        {
            PointD point = points[i];
            left = Math.Min(left, point.X);
            top = Math.Min(top, point.Y);
            right = Math.Max(right, point.X);
            bottom = Math.Max(bottom, point.Y);
        }

        return new RectD(left, top, right - left, bottom - top);
    }

    /// <summary>
    /// Selects the occupants whose vertical range overlaps a row's band, which are
    /// the only ones whose horizontal space the row's labels compete for.
    /// </summary>
    /// <param name="occupants">Every label box placed so far, in any lane.</param>
    /// <param name="band">The row's own shape bounds.</param>
    /// <returns>The overlapping subset; the same list instance when all of them do.</returns>
    private static List<RectD> VerticalBand(List<RectD> occupants, RectD band)
    {
        // `skipped` is tracked separately from `relevant`: when the *first* occupant
        // is filtered out there is nothing to copy yet, so a null `relevant` alone
        // would mean "keep everything" and silently return the unfiltered list.
        var skipped = false;
        List<RectD>? relevant = null;
        for (var i = 0; i < occupants.Count; i++)
        {
            RectD occupant = occupants[i];
            if (occupant.Bottom <= band.Top + GeometryMath.Epsilon
                || occupant.Top >= band.Bottom - GeometryMath.Epsilon)
            {
                skipped = true;
                continue;
            }

            relevant ??= [.. occupants.Take(i)];
            relevant.Add(occupant);
        }

        return relevant ?? (skipped ? [] : occupants);
    }

    /// <summary>
    /// The §22 description-label placement priority for a type, highest first.
    /// </summary>
    /// <param name="type">The entity type whose label is being placed.</param>
    /// <returns>
    /// The priority rank: critical milestones, then other milestones, delay labels,
    /// actual labels, planned labels, baseline labels, and procurement/custom labels.
    /// </returns>
    /// <remarks>
    /// §22 states the order as "critical milestones, other milestones, delay labels,
    /// actual labels, planned labels, baseline labels, procurement/custom labels, date
    /// labels, and delineator labels". The date and delineator label passes are not
    /// planned here, so those ranks need no member. A type the list does not name
    /// sorts last, which is deterministic and cannot outrank a named one.
    /// </remarks>
    private static int LabelPriorityFor(GanttEntityType type) =>
        type switch
        {
            GanttEntityType.CriticalMilestone => 0,
            GanttEntityType.AsPlannedMilestone
                or GanttEntityType.AsBuiltMilestone
                or GanttEntityType.BaselineMilestone => 1,
            GanttEntityType.DelayEvent => 2,
            GanttEntityType.AsBuiltActivity or GanttEntityType.AsBuiltProcurement => 3,
            GanttEntityType.AsPlannedActivity or GanttEntityType.AsPlannedProcurement => 4,
            GanttEntityType.BaselineActivity or GanttEntityType.BaselineProcurement => 5,
            GanttEntityType.CustomActivity => 6,

            // A Critical Interval and a Delineator have no §22 description label of
            // their own (the first is an overlay, the second places its own corner
            // text), and a Splitter or Spacer emits no label at all. Listing them
            // explicitly rather than folding them into the default arm is what lets
            // the analyzer prove the switch covers the enum, so a new type cannot be
            // added without deciding its label priority.
            GanttEntityType.CriticalInterval => 7,
            GanttEntityType.Delineator => 8,
            GanttEntityType.Splitter or GanttEntityType.Spacer => 9,
            _ => int.MaxValue,
        };

    /// <summary>
    /// Reconstructs an entity's unclipped horizontal extent from its own dates, so
    /// §23's clipped-date rule fires only on a bar the plot actually cut.
    /// </summary>
    /// <param name="event">The validated event supplying the dates.</param>
    /// <param name="timeScale">The scale every date-to-X mapping came from.</param>
    /// <param name="visible">The bar's post-plot-clip bounds.</param>
    /// <returns>
    /// The unclipped bounds: the same vertical extent, spanning the event's own
    /// start-to-finish geometry. An entity the plot did not clip gets bounds equal
    /// to <paramref name="visible"/>, so it is never reported as clipped.
    /// </returns>
    /// <remarks>
    /// The previous form returned the whole plot width, which made every bar
    /// narrower than the plot compare unequal to its "full" bounds and therefore
    /// read as clipped -- so almost every date label took §23's never-suppress
    /// fallback instead of being planned. Deriving the extent from the dates is
    /// what makes the comparison mean "the plot cut this bar".
    /// </remarks>
    private static RectD FullBoundsOf(GanttEvent @event, TimeScale timeScale, RectD visible)
    {
        // The unclipped geometry is the time scale's own linear mapping applied
        // without its range test, so a date outside the plot range extrapolates to
        // the X it would occupy rather than collapsing onto the plot edge. Collapsing
        // is what made a clipped bar compare equal to its own full bounds.
        //
        // A point event (milestone, delineator) has no span, so its full width is zero
        // and the plot can never have shortened it: it is never reported as clipped.
        //
        // Both dates are required, and the missing-Finish case must return `visible`
        // rather than a zero-width rectangle. `DateLabelBuilder` refuses a
        // non-positive `FullBounds` outright, so a Start-only event that reached it
        // with a constructed zero-width bound was refused and its date label dropped
        // rather than planned. A point event has nothing the plot can shorten, so
        // `visible` is both the truthful and the usable answer.
        if (@event.Start is not { } start || @event.Finish is not { } finish)
        {
            return visible;
        }

        var left = timeScale.PlotLeftPt
            + ((start.DayNumber - timeScale.PlotStart.DayNumber) * timeScale.DayWidth);
        var width = (finish.DayNumber - start.DayNumber + 1) * timeScale.DayWidth;

        return new RectD(left, visible.Y, Math.Max(0, width), visible.Height);
    }

    private static string StyleKey(Dictionary<GanttRowId, SceneStyle> styles, GanttEvent @event) =>
        styles.TryGetValue(@event.Id, out SceneStyle? style) ? style.StyleKey : "None";

    private static bool TryResolveStyle(
        GanttStyleRegistry registry,
        GanttEvent @event,
        out ResolvedEventStyle? resolved)
    {
        if (!GanttStyleResolver.TryResolve(
                @event.Type,
                registry,
                @event.StyleKey,
                @event.FillColour,
                @event.StrokeColour,
                @event.LabelPosition,
                out GanttResolvedStyle? style,
                out _)
            || style is null)
        {
            resolved = null;
            return false;
        }

        // The resolved height comes from the catalogue definition rather than the
        // style key alone: a capability-only definition has no formatting, and
        // using it would contribute a zero-height lane that silently hides content.
        var hasDefinition = registry.TryGet(style.StyleKey, out GanttStyleDefinition? definition)
            && definition is { HasFormatting: true };
        var heightPt = hasDefinition ? definition!.ActivityHeightPt : 0;

        // The hatch pattern and the standard outline width are resolved style
        // tokens, not renderer choices: the guide requires a renderer's output to
        // match the scene, so a hatch the scene drops can never be drawn. Both come
        // from the same resolved definition as the height, and stay at their
        // defaults for a capability-only definition that supplied no formatting.
        resolved = new ResolvedEventStyle(
            new SceneStyle(
                style.StyleKey,
                style.FillColour,
                style.StrokeColour,
                hasDefinition ? definition!.StandardOutlinePt : null,
                hasDefinition ? definition!.HatchPattern ?? GanttHatchPattern.None : GanttHatchPattern.None),
            heightPt);
        return true;
    }

    private static SceneBuildOutcome Refused(SceneBuilderRefusal refusal) => new(null, refusal);
}
