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

    /// <summary>The caller-supplied plot, panel, or chart bounds were null.</summary>
    NullBounds = 4,

    /// <summary>The plot range could not be turned into a time scale.</summary>
    InvalidPlotRange = 5,

    /// <summary>The lane metrics or the resolved frame settings were refused.</summary>
    InvalidLayoutSettings = 6,

    /// <summary>The caller-supplied plot bounds fall outside the chart bounds.</summary>
    PlotOutsideChart = 7,

    /// <summary>One event's named style could not be resolved to a renderable style.</summary>
    UnresolvableStyle = 8,
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

    /// <summary>Gets the caller-supplied chart bounds enclosing the panel and plot.</summary>
    public RectD? ChartBounds { get; init; }

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
            || request.PlotBounds is not { } plotBounds
            || request.ChartBounds is not { } chartBounds)
        {
            return Refused(SceneBuilderRefusal.NullBounds);
        }

        // RectD has no rectangle-contains-rectangle, so containment is checked at
        // the two opposite corners. A plot outside the chart would let a frame
        // label at a chart edge fall outside the chart that owns it.
        if (!chartBounds.Contains(new PointD(plotBounds.Left, plotBounds.Top))
            || !chartBounds.Contains(new PointD(plotBounds.Right, plotBounds.Bottom)))
        {
            return Refused(SceneBuilderRefusal.PlotOutsideChart);
        }

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
        // lane geometry with no entity primitive. Both are excluded here so a lane
        // is never created for an entity that cannot be drawn.
        List<GanttEvent> renderable =
        [
            .. request.Events.Where(@event =>
                @event.Visible && @event.Type is not (GanttEntityType.Splitter or GanttEntityType.Spacer)),
        ];
        if (renderable.Count == 0)
        {
            return Refused(SceneBuilderRefusal.EmptyEvents);
        }

        List<LaneEventInput> laneInputs = [];
        Dictionary<GanttRowId, ResolvedEventStyle> styles = [];
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
        BuildSpans(placements, styles, timeScale, primitives, warnings, parentVisibleBounds);
        BuildOverlaysAndMilestones(request, placements, styles, timeScale, plotBounds, parentVisibleBounds, primitives, warnings);
        return BuildFramePanelAndScene(request, timeScale, panelBounds, plotBounds, chartBounds, placements, primitives, warnings);
    }


    /// <summary>One event's resolved scene style and its resolved height.</summary>
    /// <param name="Style">The resolved scene style.</param>
    /// <param name="HeightPt">The resolved activity height from the catalogue.</param>
    private sealed record ResolvedEventStyle(SceneStyle Style, double HeightPt);

    private static void BuildSpans(
        LaneEventLayoutResult placements,
        IReadOnlyDictionary<GanttRowId, ResolvedEventStyle> styles,
        TimeScale timeScale,
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
            SpanBarCreationOutcome bar = SpanBarBuilder.TryBuild(
                new SpanBarRequest(
                    @event,
                    resolved.Style,
                    placement.SlotCentreY,
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
                        placement.SlotCentreY - (resolved.HeightPt / 2),
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
                    placement.SlotCentreY,
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
        RectD chartBounds,
        LaneEventLayoutResult placements,
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

        if (request.Panel is { } panelTheme)
        {
            // Section 4 fixes the header band's bottom edge to the period header's
            // bottom, which R3.5 derives as PlotBounds.Y - YearBandHeightPt. The
            // panel builder cannot know the plot bounds, so it is supplied here and
            // SceneBuilderTests asserts the emitted bottom equals this value.
            PanelBuildOutcome panel = PanelBuilder.TryBuild(
                new PanelBuildRequest(
                    request.Grid!,
                    [.. placements.Placements.Select(placement => new PanelRow(placement.Event.Id, []))],
                    plotBounds,
                    plotBounds.Y - request.YearBandHeightPt,
                    request.YearBandHeightPt,
                    panelTheme));
            if (panel.Result is { } panelResult)
            {
                primitives.AddRange(panelResult.Primitives);
            }
        }

        SceneCreationOutcome scene = GanttScene.TryCreate(chartBounds, plotBounds, primitives, warnings);
        return scene.Scene is { } built
            ? new SceneBuildOutcome(new SceneBuildResult(built, timeScale), null)
            : Refused(SceneBuilderRefusal.InvalidLayoutSettings);
    }

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
