namespace GanttCreator.Core.Scene;

/// <summary>One delineator request.</summary>
/// <param name="Event">The validated delineator event. Only <c>Start</c> is read.</param>
/// <param name="LineStyle">The resolved named style; its outline width is replaced by <paramref name="LineWidthPt"/>.</param>
/// <param name="LineWidthPt">The line thickness, from the <c>DelineatorLinePt</c> token.</param>
/// <param name="PlotBounds">The plot rectangle; the line spans its top and bottom exactly.</param>
/// <param name="ChartBounds">The chart rectangle that must contain every label box.</param>
/// <param name="LabelGapPt">The offset between the line and a label box.</param>
/// <param name="LabelMetrics">The injected deterministic text-measuring seam.</param>
/// <param name="LabelPosition">The resolved label position, or <c>None</c> for no label.</param>
/// <param name="LaneOrder">The lane ordering value, when known.</param>
/// <param name="StackIndex">The stack ordering value, when known.</param>
public sealed record DelineatorRequest(
    GanttEvent Event,
    SceneStyle LineStyle,
    double LineWidthPt,
    RectD PlotBounds,
    RectD ChartBounds,
    double LabelGapPt,
    ITextMetrics LabelMetrics,
    GanttLabelPosition LabelPosition = GanttLabelPosition.Auto,
    int? LaneOrder = null,
    int? StackIndex = null
);

/// <summary>The result of building one delineator.</summary>
/// <param name="Primitive">The line, or <see langword="null"/> when the date lies outside the plot.</param>
/// <param name="X">The exact line X, or <see langword="null"/> when no line was emitted.</param>
/// <param name="Label">The corner label, or <see langword="null"/> when no label was emitted.</param>
/// <param name="LabelBounds">The resolved label box, or <see langword="null"/>.</param>
/// <param name="LabelPosition">The corner actually used, or <see langword="null"/>.</param>
/// <param name="Warnings">The deterministic non-blocking warnings.</param>
public sealed record DelineatorResult(
    SceneLine? Primitive,
    double? X,
    SceneText? Label,
    RectD? LabelBounds,
    GanttLabelPosition? LabelPosition,
    IReadOnlyList<SceneWarning> Warnings
);

/// <summary>The reason a delineator could not be built.</summary>
public enum DelineatorRefusal
{
    /// <summary>The request was null.</summary>
    NullRequest = 0,

    /// <summary>The event, its style, or its text-metrics seam was null.</summary>
    NullDependency = 1,

    /// <summary>The time scale was invalid.</summary>
    InvalidTimeScale = 2,

    /// <summary>A bound, gap, or width was not finite, or a bound was empty.</summary>
    InvalidGeometry = 3,

    /// <summary>The event carries no start date, so a full-height line has nothing to place.</summary>
    MissingEventDate = 4,

    /// <summary>The event is not a delineator type.</summary>
    NotADelineator = 5,

    /// <summary>The label position is not one the entity catalogue permits for a delineator.</summary>
    UnsupportedLabelPosition = 6,
}

/// <summary>The typed result of attempting to build one delineator.</summary>
/// <param name="Result">The successful result, or <see langword="null"/>.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/>.</param>
public sealed record DelineatorCreationOutcome(DelineatorResult? Result, DelineatorRefusal? Refusal)
{
    /// <summary>Gets whether construction succeeded.</summary>
    public bool Succeeded => Result is not null;
}

/// <summary>
/// Builds deterministic full-height delineator lines and their corner labels
/// (entity guide §24), with no Office dependency.
/// </summary>
/// <remarks>
/// <para>
/// §24 geometry: X is the exact date position and the line runs from
/// <c>PlotBounds.Top</c> to <c>PlotBounds.Bottom</c>, never through the title,
/// data-panel, or time-header bands, which is why the plot bounds — not the
/// chart bounds — are the span. Only <c>Start</c> is read; <c>Finish</c>,
/// <c>LaneId</c>, and <c>StackIndex</c> are never consulted for geometry.
/// </para>
/// <para>
/// The corner label positions belong to this builder rather than to
/// <see cref="LabelPlanner"/>: the planner explicitly refuses them and names
/// R3.10 as their owner, because a corner of a full-height line is a different
/// placement problem from the §22 span cascade. The single shared
/// <see cref="ITextMetrics"/> seam is still used, so measurement behaviour is
/// identical across every label in the scene.
/// </para>
/// </remarks>
public static class DelineatorBuilder
{
    /// <summary>The warning code emitted when a delineator date lies outside the plot.</summary>
    public const string OutsidePlotRangeCode = "DelineatorOutsidePlotRange";

    /// <summary>The warning code emitted when no permitted corner can hold the label.</summary>
    public const string LabelSuppressedCode = "DelineatorLabelSuppressed";

    /// <summary>
    /// The §24 <c>Auto</c> order: <c>TopRight</c> is tried first and
    /// <c>BottomLeft</c> last. The first corner contained within the chart
    /// bounds and clear of the registered occupants wins.
    /// </summary>
    private static readonly GanttLabelPosition[] _autoOrder =
    [
        GanttLabelPosition.TopRight,
        GanttLabelPosition.TopLeft,
        GanttLabelPosition.BottomRight,
        GanttLabelPosition.BottomLeft,
    ];

    /// <summary>Attempts to build one delineator line and its label.</summary>
    /// <param name="request">The typed delineator request.</param>
    /// <param name="timeScale">The validated time scale.</param>
    /// <param name="occupants">Already-placed label boxes this label must not intersect.</param>
    /// <param name="stackOffsetPt">
    /// A vertical offset applied away from the plot edge this label sits
    /// against, so several same-date labels stack instead of overlapping.
    /// </param>
    /// <returns>A typed result or refusal.</returns>
    public static DelineatorCreationOutcome TryBuild(
        DelineatorRequest? request,
        TimeScale? timeScale,
        IReadOnlyList<RectD>? occupants = null,
        double stackOffsetPt = 0)
    {
        if (request is null)
        {
            return Refused(DelineatorRefusal.NullRequest);
        }

        if (timeScale is null)
        {
            return Refused(DelineatorRefusal.InvalidTimeScale);
        }

        if (request.Event is null || request.LineStyle is null || request.LabelMetrics is null)
        {
            return Refused(DelineatorRefusal.NullDependency);
        }

        if (!double.IsFinite(request.LineWidthPt)
            || request.LineWidthPt <= 0
            || !double.IsFinite(request.LabelGapPt)
            || !double.IsFinite(stackOffsetPt)
            || !IsUsable(request.PlotBounds)
            || !IsUsable(request.ChartBounds))
        {
            return Refused(DelineatorRefusal.InvalidGeometry);
        }

        GanttEvent @event = request.Event;
        EntityTypeDefinition? definition = EntityTypeCatalog.GetDefinition(@event.Type);
        if (definition is null || definition.Kind != EntityKind.Delineator)
        {
            return Refused(DelineatorRefusal.NotADelineator);
        }

        // §24: Start is required and is the only date a delineator reads.
        if (@event.Start is not { } date)
        {
            return Refused(DelineatorRefusal.MissingEventDate);
        }

        // The catalogue is the authority: a position it does not permit for a
        // delineator is refused rather than rendered anyway.
        if (!definition.AllowedLabelPositions.Contains(request.LabelPosition))
        {
            return Refused(DelineatorRefusal.UnsupportedLabelPosition);
        }

        List<SceneWarning> warnings = [];
        var owner = SceneOwnerId.ForRow(@event.Id);

        // The `Try*` form is used deliberately: DateToX throws for a date
        // outside the scale, and an out-of-range delineator must warn, not throw.
        if (!timeScale.TryDateToX(date, out var x))
        {
            warnings.Add(new SceneWarning(
                owner,
                OutsidePlotRangeCode,
                "The delineator date lies outside the plot range, so no line was drawn."));
            return new DelineatorCreationOutcome(
                new DelineatorResult(null, null, null, null, null, warnings), null);
        }



        // §24: the line spans the plot top and bottom exactly. Using the chart
        // bounds here would run the line through the title and time-header
        // bands, which §24 forbids.
        var line = new SceneLine(
            ScenePrimitive.CreateId(owner, "delineator"),
            owner,
            ZLayer.Delineator,
            new PointD(x, request.PlotBounds.Top),
            new PointD(x, request.PlotBounds.Bottom),
            WithWidth(request.LineStyle, request.LineWidthPt),
            @event.Type,
            request.LaneOrder,
            request.StackIndex,
            @event.SortOrder);

        // §24 allows a blank description only when the label position is None.
        // The value is captured once here so the label path never re-reads a
        // nullable property the compiler cannot prove is non-null.
        var description = @event.Description;
        if (request.LabelPosition == GanttLabelPosition.None || string.IsNullOrEmpty(description))
        {
            return new DelineatorCreationOutcome(
                new DelineatorResult(line, x, null, null, null, warnings), null);
        }

        var placed = TryPlaceLabel(
            request,
            @event,
            description,
            x,
            occupants ?? [],
            stackOffsetPt,
            out SceneText? label,
            out RectD? labelBounds,
            out GanttLabelPosition? usedPosition);

        if (!placed)
        {
            warnings.Add(new SceneWarning(
                owner,
                LabelSuppressedCode,
                "No permitted corner held the delineator label, so no label was drawn."));
        }

        return new DelineatorCreationOutcome(
            new DelineatorResult(line, x, label, labelBounds, usedPosition, warnings), null);
    }

    private static bool TryPlaceLabel(
        DelineatorRequest request,
        GanttEvent @event,
        string text,
        double lineX,
        IReadOnlyList<RectD> occupants,
        double stackOffsetPt,
        out SceneText? label,
        out RectD? labelBounds,
        out GanttLabelPosition? usedPosition)
    {
        label = null;
        labelBounds = null;
        usedPosition = null;

        if (!request.LabelMetrics.TryMeasure(text, out TextMeasurement? measurement) || measurement is null)
        {
            return false;
        }

        IReadOnlyList<GanttLabelPosition> candidates = request.LabelPosition == GanttLabelPosition.Auto
            ? _autoOrder
            : [request.LabelPosition];

        foreach (GanttLabelPosition position in candidates)
        {
            RectD box = CornerBox(
                position, lineX, measurement, request.PlotBounds, request.LabelGapPt, stackOffsetPt);

            if (!Contains(request.ChartBounds, box) || IntersectsAny(box, occupants))
            {
                continue;
            }

            // §24 is silent on label alignment and SceneText requires a value.
            // The text hugs the line: a right-hand corner reads away from the
            // line, a left-hand corner is right-aligned so it also ends against
            // the line (product owner, 2026-09-26). ADR-0018 retypes this to
            // GanttTextAlignment; the rule and its rationale are unchanged.
            GanttTextAlignment alignment = position is GanttLabelPosition.TopLeft or GanttLabelPosition.BottomLeft
                ? GanttTextAlignment.Right
                : GanttTextAlignment.Left;

            var owner = SceneOwnerId.ForRow(@event.Id);
            label = new SceneText(
                ScenePrimitive.CreateId(owner, "delineator-label"),
                owner,
                ZLayer.DelineatorLabel,
                text,
                box,
                request.LineStyle,
                alignment,
                @event.Type,
                request.LaneOrder,
                request.StackIndex,
                @event.SortOrder);
            labelBounds = box;
            usedPosition = position;
            return true;
        }

        return false;
    }

    /// <summary>
    /// Returns the style with the delineator's own line width applied.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <see cref="SceneStyle"/> exposes read-only properties, so it is rebuilt
    /// rather than copied with a <c>with</c> expression. §24 fixes the width at
    /// the <c>DelineatorLinePt</c> token, so the width is never taken from the
    /// named style.
    /// </para>
    /// <para>
    /// Every member is forwarded, including the optional
    /// <c>textColour</c> (R4.11). Passing positionally made this call silently
    /// drop it, because it is the tenth optional parameter: a delineator whose
    /// named style declared a text colour rendered with the host's default. The
    /// golden snapshot could not catch that, since the reference fixture's
    /// delineator style declares no text colour and so had nothing to lose.
    /// </para>
    /// </remarks>
    private static SceneStyle WithWidth(SceneStyle style, double widthPt) =>
        new(
            style.StyleKey,
            style.FillColour,
            style.StrokeColour,
            widthPt,
            style.HatchPattern,
            style.FontFamily,
            style.FontSizePt,
            style.Bold,
            style.Alignment,
            style.TextColour);

    /// <summary>Builds one corner's box, offsetting the stack away from the plot edge.</summary>
    private static RectD CornerBox(
        GanttLabelPosition position,
        double lineX,
        TextMeasurement measurement,
        RectD plot,
        double labelGapPt,
        double stackOffsetPt)
    {
        var left = position is GanttLabelPosition.TopLeft or GanttLabelPosition.BottomLeft;
        var x = left ? lineX - labelGapPt - measurement.WidthPt : lineX + labelGapPt;

        // The stack grows away from the edge the label sits against, so a
        // Top-anchored label steps down and a Bottom-anchored label steps up;
        // neither walks out of the plot.
        var top = position is GanttLabelPosition.TopLeft or GanttLabelPosition.TopRight;
        var y = top ? plot.Top + stackOffsetPt : plot.Bottom - measurement.HeightPt - stackOffsetPt;

        return new RectD(x, y, measurement.WidthPt, measurement.HeightPt);
    }

    private static bool IsUsable(RectD bounds) =>
        double.IsFinite(bounds.X)
        && double.IsFinite(bounds.Y)
        && double.IsFinite(bounds.Width)
        && double.IsFinite(bounds.Height)
        && bounds.Width > GeometryMath.Epsilon
        && bounds.Height > GeometryMath.Epsilon;

    private static bool Contains(RectD outer, RectD inner) =>
        inner.Left >= outer.Left - GeometryMath.Epsilon
        && inner.Top >= outer.Top - GeometryMath.Epsilon
        && inner.Right <= outer.Right + GeometryMath.Epsilon
        && inner.Bottom <= outer.Bottom + GeometryMath.Epsilon;

    private static bool IntersectsAny(RectD candidate, IReadOnlyList<RectD> occupants)
    {
        foreach (RectD occupant in occupants)
        {
            if (candidate.IntersectsWith(occupant))
            {
                return true;
            }
        }

        return false;
    }

    private static DelineatorCreationOutcome Refused(DelineatorRefusal refusal) => new(null, refusal);
}

/// <summary>One same-date delineator group for the layout pass.</summary>
/// <param name="Requests">
/// The delineators sharing one date and one resolved line style. The order is
/// irrelevant: <see cref="DelineatorLayout"/> sorts it.
/// </param>
/// <param name="StackGapPt">The vertical gap between stacked same-date labels.</param>
public sealed record DelineatorGroupRequest(
    IReadOnlyList<DelineatorRequest> Requests,
    double StackGapPt
);

/// <summary>The result of laying out one same-date group.</summary>
/// <param name="Primitives">The line and labels, in deterministic emission order.</param>
/// <param name="Line">The single deduplicated line, or <see langword="null"/> when out of range.</param>
/// <param name="Warnings">The deterministic non-blocking warnings for the group.</param>
public sealed record DelineatorGroupResult(
    IReadOnlyList<ScenePrimitive> Primitives,
    SceneLine? Line,
    IReadOnlyList<SceneWarning> Warnings
);

/// <summary>The reason a same-date group could not be laid out.</summary>
public enum DelineatorGroupRefusal
{
    /// <summary>The request or its request list was null or empty.</summary>
    NullRequest = 0,

    /// <summary>The time scale was absent.</summary>
    InvalidTimeScale = 1,

    /// <summary>A request in the group carried different plot or chart bounds.</summary>
    InconsistentBounds = 2,

    /// <summary>Two requests in the group disagree on the label gap or metrics.</summary>
    InconsistentLabelSeam = 3,

    /// <summary>The group members are not all delineators on the same date.</summary>
    NotOneDelineatorDate = 4,

    /// <summary>
    /// A member's own <see cref="DelineatorBuilder.TryBuild"/> refused, so the
    /// group cannot be laid out honestly. Reported rather than skipped, because a
    /// skipped member would remove a contributing row from the scene silently.
    /// </summary>
    MemberRefused = 5,

    /// <summary>
    /// The members do not share one resolved line style. §24 draws the line once
    /// per resolved line style, so a mixed group has no single line to emit: the
    /// members must be grouped by style first, and silently styling them all from
    /// the first request would discard the others' resolved appearance.
    /// </summary>
    InconsistentLineStyle = 6,
}

/// <summary>The typed result of laying out one same-date group.</summary>
/// <param name="Result">The successful result, or <see langword="null"/>.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/>.</param>
public sealed record DelineatorGroupCreationOutcome(
    DelineatorGroupResult? Result,
    DelineatorGroupRefusal? Refusal)
{
    /// <summary>Gets whether construction succeeded.</summary>
    public bool Succeeded => Result is not null;
}

/// <summary>
/// Lays out the same-date groups §24 requires: one line per resolved line
/// style, owned by every row that contributed to it (ADR-0017), with distinct
/// labels stacked by <c>StackGapPt</c>.
/// </summary>
public static class DelineatorLayout
{
    /// <summary>Attempts to lay out one same-date group.</summary>
    /// <param name="request">The typed group request.</param>
    /// <param name="timeScale">The validated time scale.</param>
    /// <returns>A typed result or refusal.</returns>
    public static DelineatorGroupCreationOutcome TryBuildGroup(
        DelineatorGroupRequest? request,
        TimeScale? timeScale)
    {
        if (request is null || request.Requests is null || request.Requests.Count == 0)
        {
            return new DelineatorGroupCreationOutcome(null, DelineatorGroupRefusal.NullRequest);
        }

        if (timeScale is null)
        {
            return new DelineatorGroupCreationOutcome(null, DelineatorGroupRefusal.InvalidTimeScale);
        }

        if (!double.IsFinite(request.StackGapPt) || request.StackGapPt < 0)
        {
            return new DelineatorGroupCreationOutcome(null, DelineatorGroupRefusal.InconsistentLabelSeam);
        }

        DelineatorRequest first = request.Requests[0];
        if (request.Requests.Any(item => item.PlotBounds != first.PlotBounds || item.ChartBounds != first.ChartBounds))
        {
            return new DelineatorGroupCreationOutcome(null, DelineatorGroupRefusal.InconsistentBounds);
        }

        if (request.Requests.Any(item =>
            !ReferenceEquals(item.LabelMetrics, first.LabelMetrics) || item.LabelGapPt != first.LabelGapPt))
        {
            return new DelineatorGroupCreationOutcome(null, DelineatorGroupRefusal.InconsistentLabelSeam);
        }

        // §24 draws one line per *resolved line style*, so a group is only a group
        // when every member resolved the same style and width. The group emits a
        // single line taken from the first member; without this check a mixed group
        // would silently restyle every other member's line as the first's. The
        // caller groups by style before building, so mixed styles are a broken
        // request, not a placement decision. Compared by value because SceneStyle
        // is a record and two independently resolved equal styles are the same
        // resolved style.
        if (request.Requests.Any(item => item.LineStyle != first.LineStyle || item.LineWidthPt != first.LineWidthPt))
        {
            return new DelineatorGroupCreationOutcome(null, DelineatorGroupRefusal.InconsistentLineStyle);
        }

        if (request.Requests.Any(item => item.Event is null || item.Event.Start != first.Event?.Start))
        {
            return new DelineatorGroupCreationOutcome(null, DelineatorGroupRefusal.NotOneDelineatorDate);
        }

        // The scene's standard order: lane, then stack, then explicit sort
        // order, then the stable identifier. It is the same tiebreak the rest of
        // the scene uses, so the stack is reproducible across refreshes.
        List<DelineatorRequest> ordered =
        [
            .. request.Requests
                .OrderBy(item => item.LaneOrder ?? int.MaxValue)
                .ThenBy(item => item.StackIndex ?? int.MaxValue)
                .ThenBy(item => item.Event?.SortOrder ?? int.MaxValue)
                .ThenBy(item => item.Event?.Id.Value ?? string.Empty, StringComparer.Ordinal),
        ];

        List<GanttRowId> owners = [.. ordered.Select(item => item.Event.Id).Distinct()];
        SceneOwnerId owner = owners.Count > 1
            ? SceneOwnerId.ForRows(owners)
            : SceneOwnerId.ForRow(owners[0]);

        List<SceneWarning> warnings = [];
        List<ScenePrimitive> primitives = [];
        SceneLine? line = null;
        List<RectD> placed = [];

        // The cumulative stack offset, in points away from the plot edge the
        // label sits against. It advances by each *successfully placed* label's
        // own height plus StackGapPt, not by one gap per label: a gap is the
        // space between two labels, so advancing by the gap alone left the second
        // label overlapping the first whenever StackGapPt was smaller than the
        // label height, which §24's "stacked deterministically with StackGapPt"
        // forbids. A member whose label was suppressed advances nothing, because
        // it occupied no space.
        var stackOffsetPt = 0.0;

        foreach (DelineatorRequest item in ordered)
        {
            // Every row in the group shares the one line, so only the first builds
            // it; the rest contribute only their label. The shared owner makes the
            // line's identity change when the membership changes, which is what
            // lets refresh remove and recreate rather than mutate.
            //
            // The first member needs one TryBuild, not two: the very same outcome
            // carries both the group line and that member's label. Building it
            // twice ran the label placement pass twice for identical inputs, and
            // the first pass's label was discarded, so the member consumed a stack
            // offset's worth of gap it never occupied.
            var needsLabel =
                item.LabelPosition != GanttLabelPosition.None
                && !string.IsNullOrEmpty(item.Event?.Description);

            DelineatorCreationOutcome outcome = DelineatorBuilder.TryBuild(
                item,
                timeScale,
                placed,
                needsLabel ? stackOffsetPt : 0);

            // A refused member is a broken input, not an absent label. Silently
            // skipping it would drop a contributing row from the scene with no
            // warning, so the group refuses with the honest reason instead.
            if (outcome is not { Result: { } result, Refusal: null })
            {
                // A refused member is a broken input, not an absent label. Silently
                // skipping it would drop a contributing row from the scene with no
                // warning, so the group refuses with the honest reason instead.
                return new DelineatorGroupCreationOutcome(null, DelineatorGroupRefusal.MemberRefused);
            }

            // Warnings come from every outcome, not just the line-building one:
            // a member whose label no corner can hold emits a suppression warning
            // that would otherwise be lost.
            warnings.AddRange(result.Warnings);

            if (line is null && result.Primitive is { } builtLine)
            {
                line = Rebind(builtLine, owner);
            }

            if (!needsLabel)
            {
                continue;
            }

            if (result.Label is { } text)
            {
                primitives.Add(text);
            }

            if (result.LabelBounds is { } bounds)
            {
                placed.Add(bounds);

                // The next label clears this one entirely: the placed box's own
                // height is the space it occupies, and StackGapPt is the space left
                // between them. Advancing by the height rather than by one gap per
                // label is what keeps a stack correct when the gap is smaller than
                // the label height, and for labels of differing heights. The offset
                // is a distance from the plot edge whichever corner the label took,
                // so accumulating it is correct for a top and a bottom label alike.
                stackOffsetPt += bounds.Height + request.StackGapPt;
            }
        }

        if (line is not null)
        {
            primitives.Insert(0, line);
        }

        return new DelineatorGroupCreationOutcome(
            new DelineatorGroupResult(primitives, line, warnings), null);
    }

    private static SceneLine Rebind(SceneLine built, SceneOwnerId owner) =>
        new(
            ScenePrimitive.CreateId(owner, "delineator"),
            owner,
            built.ZLayer,
            built.From,
            built.To,
            built.Style,
            built.EntityType,
            built.LaneOrder,
            built.StackIndex,
            built.SortOrder);
}
