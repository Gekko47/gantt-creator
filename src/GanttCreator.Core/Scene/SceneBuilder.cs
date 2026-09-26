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
                "Type" => @event.Type.ToString(),
                "Description" => @event.Description,
                "Start" => @event.Start is { } start ? GanttDateFormatting.Format(start, request.DateFormat) : null,
                "Finish" => @event.Finish is { } finish ? GanttDateFormatting.Format(finish, request.DateFormat) : null,
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
    /// §23 requires a date label to avoid the row's own description label, and §22
    /// requires labels to avoid each other. Both are satisfied by one growing
    /// <c>occupants</c> list that every placed label joins, which is also the
    /// contract <see cref="DateLabelRequest.Occupants"/> documents. A label is
    /// registered as an occupant only once it is actually emitted, so a suppressed
    /// label reserves nothing.
    /// </remarks>
    private static void BuildLabels(
        SceneBuildRequest request,
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
            plotBounds.Width);

        List<RectD> occupants = [];

        foreach (LaneEventPlacement placement in placements.Placements)
        {
            GanttEvent @event = placement.Event;
            if (!parentVisibleBounds.TryGetValue(@event.Id, out RectD shapeBounds))
            {
                // No emitted bar or marker: there is nothing to label. A span wholly
                // outside the plot still needs no label, because its date is not shown
                // on the shape either.
                continue;
            }

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
                occupants);
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

            // The full span is the post-plot-clip bounds with the clip released, so
            // the builder can tell a clipped event from a whole one. DateLabelBuilder
            // derives that itself from the two rectangles; the unclipped rectangle is
            // the visible bounds widened back to the time scale's own extent.
            DateLabelOutcome dates = DateLabelBuilder.TryBuild(
                new DateLabelRequest(
                    @event,
                    shapeBounds,
                    FullBoundsOf(shapeBounds, plotBounds),
                    metrics,
                    request.Metrics!,
                    request.DateFormat,
                    labelStyle,
                    Occupants: occupants));
            if (dates.Result is not { } planned)
            {
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
    /// Reconstructs a span's unclipped bounds from its clipped ones, so §23's
    /// clipped-date rule can fire on a bar the plot cut short.
    /// </summary>
    /// <param name="visible">The bar's post-plot-clip bounds.</param>
    /// <param name="plotBounds">The plot rectangle the clip was taken against.</param>
    /// <returns>The unclipped bounds: the same vertical extent, spanning the plot width.</returns>
    private static RectD FullBoundsOf(RectD visible, RectD plotBounds) =>
        new(plotBounds.Left, visible.Y, plotBounds.Width, visible.Height);

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
        var heightPt = registry.TryGet(style.StyleKey, out GanttStyleDefinition? definition) && definition is { HasFormatting: true }
            ? definition.ActivityHeightPt
            : 0;

        resolved = new ResolvedEventStyle(
            new SceneStyle(style.StyleKey, style.FillColour, style.StrokeColour),
            heightPt);
        return true;
    }

    private static SceneBuildOutcome Refused(SceneBuilderRefusal refusal) => new(null, refusal);
}
