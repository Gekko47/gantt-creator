namespace GanttCreator.Core.Scene;

/// <summary>One delineator request.</summary>
/// <param name="Event">The validated delineator event. Only <c>Start</c> is read.</param>
/// <param name="LineStyle">The resolved named style; its outline width is replaced by <paramref name="LineWidthPt"/>.</param>
/// <param name="LabelStyle">
/// The resolved description-label text style, or <see langword="null"/>
/// for no label. This is the same predetermined style every other
/// description-label path is gated on (<see cref="SceneBuildRequest.LabelStyle"/>),
/// so a delineator label carries it too and renders in the typography the
/// measuring seam measured at; the line style governs the line alone.
/// </param>
/// <param name="LineWidthPt">The line thickness, from the <c>DelineatorLinePt</c> token.</param>
/// <param name="PlotBounds">The plot rectangle; the line spans its top and bottom exactly.</param>
/// <param name="ChartBounds">The chart rectangle that must contain every label box.</param>
/// <param name="LabelGapPt">The offset between the line and a label box.</param>
/// <param name="LabelMetrics">The injected deterministic text-measuring seam.</param>
/// <param name="LabelPosition">The resolved label position, or <c>None</c> for no label.</param>
/// <param name="LaneOrder">The lane ordering value, when known.</param>
/// <param name="StackIndex">The stack ordering value, when known.</param>
/// <param name="LineTopPt">
/// The line's top Y, or <see langword="null"/> to start it exactly at
/// <see cref="PlotBounds"/>'s top edge (ADR-0037 D1).
/// </param>
/// <param name="LineBottomPt">
/// The line's bottom Y, or <see langword="null"/> to end it exactly at
/// <see cref="PlotBounds"/>'s bottom edge (ADR-0038 D2).
/// </param>
/// <param name="DateFormat">
/// The approved invariant display format (ADR-0016) for the §23
/// start-date label the row's <see cref="GanttEvent.DateLabelPlacement"/>
/// selects.
/// </param>
/// <param name="ShowDateLabels">
/// Whether the §23 start-date label is requested at all (the date-labels
/// toggle gates every date-label path, the delineator's included).
/// </param>
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
    int? StackIndex = null,
    double? LineBottomPt = null,
    double? LineTopPt = null,
    SceneStyle? LabelStyle = null,
    GanttDateDisplayFormat DateFormat = GanttDateDisplayFormat.DdMMyyyy,
    bool ShowDateLabels = true
);

/// <summary>The result of building one delineator.</summary>
/// <param name="Primitive">The line, or <see langword="null"/> when the date lies outside the plot.</param>
/// <param name="X">The exact line X, or <see langword="null"/> when no line was emitted.</param>
/// <param name="Label">The corner label, or <see langword="null"/> when no label was emitted.</param>
/// <param name="LabelBounds">The resolved label box, or <see langword="null"/>.</param>
/// <param name="LabelPosition">The corner actually used, or <see langword="null"/>.</param>
/// <param name="DateLabel">
/// The §23 start-date label, or <see langword="null"/> when no date label
/// was emitted (hidden, not requested, or no permitted corner held it).
/// </param>
/// <param name="DateLabelBounds">The resolved date-label box, or <see langword="null"/>.</param>
/// <param name="Warnings">The deterministic non-blocking warnings.</param>
public sealed record DelineatorResult(
    SceneLine? Primitive,
    double? X,
    SceneText? Label,
    RectD? LabelBounds,
    GanttLabelPosition? LabelPosition,
    SceneText? DateLabel,
    RectD? DateLabelBounds,
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
    /// The warning code emitted when the requested date-label corner
    /// fits nowhere inside the chart (the explicit-position discipline:
    /// a requested box is never silently moved, so a miss is reported
    /// rather than re-placed).
    /// </summary>
    public const string DateLabelSuppressedCode = "DelineatorDateLabelSuppressed";

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
                new DelineatorResult(null, null, null, null, null, null, null, warnings), null);
        }



        // §24: the line spans the plot top and bottom exactly. Using the chart
        // bounds here would run the line through the title and time-header
        // bands, which §24 forbids.
        //
        // ADR-0038 D2: the bottom may be the CLOSING line's Y instead, so a
        // delineator shares the plot-spanning shapes' extended vertical span. A
        // delineator left at the plot's own bottom would be the one full-height line
        // still stopping at the old footer after a row insert -- the same defect as
        // the bands, one primitive out of step with them.
        //
        // ADR-0037 D1 EXTENDED HERE: the top is lifted into the header row for the
        // same reason the bands are. Measured live 2026-10-04
        // (scripts/probe-delineator-top.ps1): with the top exactly on the
        // header/body boundary a row added at the top of the body moved the line
        // TopDelta +15.75 with HeightDelta 0 -- it SLID, anchor D2 -> D3 -- while
        // the bands stretched correctly beside it. Lifting the top 0.5pt gave
        // TopDelta 0 / HeightDelta +15.75, anchor D1 -> D1. The delineator was
        // simply never given ADR-0037's lift, so it was the last plot-spanning
        // shape left on the boundary.
        //
        // THE LABEL CORNERS STILL RESOLVE AGAINST PlotBounds, at both ends. That is
        // why the line's extent is two separate members rather than a widened
        // PlotBounds: extending that box would drag every corner label with it.
        var line = new SceneLine(
            ScenePrimitive.CreateId(owner, "delineator"),
            owner,
            ZLayer.Delineator,
            new PointD(x, request.LineTopPt ?? request.PlotBounds.Top),
            new PointD(x, request.LineBottomPt ?? request.PlotBounds.Bottom),
            WithWidth(request.LineStyle, request.LineWidthPt),
            @event.Type,
            request.LaneOrder,
            request.StackIndex,
            @event.SortOrder);

        // §24 allows a blank description only when the label position is None.
        // The value is captured once here so the label path never re-reads a
        // nullable property the compiler cannot prove is non-null.
        //
        // A missing resolved label style is the same no-label decision every
        // other description-label path makes: BuildLabels and SplitterBuilder
        // both skip silently when the predetermined style is absent, so a
        // delineator label is gated on it too rather than falling back to the
        // line's style, which carries no font the measuring seam recognises.
        var description = @event.Description;
        SceneText? label = null;
        RectD? labelBounds = null;
        GanttLabelPosition? usedPosition = null;
        if (request.LabelPosition != GanttLabelPosition.None
            && !string.IsNullOrEmpty(description)
            && request.LabelStyle is not null)
        {
            var placed = TryPlaceLabel(
                request,
                @event,
                description,
                x,
                occupants ?? [],
                stackOffsetPt,
                warnings,
                out label,
                out labelBounds,
                out usedPosition);

            if (!placed)
            {
                warnings.Add(new SceneWarning(
                    owner,
                    LabelSuppressedCode,
                    "No permitted corner held the delineator label, so no label was drawn."));
            }
        }

        // §23's start-date label. A delineator reads Start only, and its
        // date label is positionable with the description label's corner set
        // (owner ruling 2026-10-14): the cell names a corner, not a
        // which-date choice, so Start/Finish never apply. It is planned
        // against the same occupants the description label saw plus that
        // label's own box, so the two labels of one row never overlap. The
        // same no-label style gate applies: the date label carries the
        // predetermined DefaultText style or it is not drawn at all.
        SceneText? dateLabel = null;
        RectD? dateLabelBounds = null;
        if (request.ShowDateLabels
            && request.LabelStyle is not null
            && @event.DateLabelPlacement is not GanttLabelPosition.None)
        {
            var datePlaced = TryPlaceDateLabel(
                request,
                @event,
                date,
                x,
                labelBounds,
                occupants ?? [],
                stackOffsetPt,
                warnings,
                out dateLabel,
                out dateLabelBounds);

            // The description label warns on any suppression, so the
            // date label does too: Auto means "show the date at the
            // first free corner", and a corner that holds neither
            // label is a loss the user should see reported.
            if (!datePlaced)
            {
                warnings.Add(new SceneWarning(
                    owner,
                    DateLabelSuppressedCode,
                    "No permitted corner held the delineator date label, so no date label was drawn."));
            }
        }

        return new DelineatorCreationOutcome(
            new DelineatorResult(
                line, x, label, labelBounds, usedPosition, dateLabel, dateLabelBounds, warnings),
            null);
    }

    private static bool TryPlaceLabel(
        DelineatorRequest request,
        GanttEvent @event,
        string text,
        double lineX,
        IReadOnlyList<RectD> occupants,
        double stackOffsetPt,
        List<SceneWarning> warnings,
        out SceneText? label,
        out RectD? labelBounds,
        out GanttLabelPosition? usedPosition)
    {
        label = null;
        labelBounds = null;
        usedPosition = null;

        IReadOnlyList<GanttLabelPosition> candidates = request.LabelPosition == GanttLabelPosition.Auto
            ? _autoOrder
            : [request.LabelPosition];

        foreach (GanttLabelPosition position in candidates)
        {
            // §22's overflow rule applies to a delineator label exactly
            // as it applies to a span's: the corner's available width is
            // the chart-bounded distance from the line (plus the gap) to
            // the chart's far edge, and a description that measures
            // wider is clipped with an ellipsis to that width — the one
            // truncation mechanism Core uses — never shrunk, wrapped, or
            // suppressed. Measuring the clipped text is what keeps the
            // box the host draws inside the box the scene measured.
            var left = position is GanttLabelPosition.TopLeft or GanttLabelPosition.BottomLeft;
            var availablePt = left
                ? lineX - request.LabelGapPt - request.ChartBounds.Left
                : request.ChartBounds.Right - lineX - request.LabelGapPt;
            if (availablePt <= 0)
            {
                continue;
            }

            var labelText = text;
            var truncated = false;
            if (LabelText.Overflows(text, availablePt, request.LabelMetrics))
            {
                labelText = LabelText.Ellipsize(text, availablePt, request.LabelMetrics);
                truncated = true;
            }

            if (!request.LabelMetrics.TryMeasure(labelText, out TextMeasurement? measurement) || measurement is null)
            {
                return false;
            }

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
                labelText,
                box,
                // The predetermined description-label style, the same
                // DefaultText style every other description label carries:
                // the line style governs the line alone. It carries the
                // font family and size the measuring seam measured at, so
                // the host draws the text the box was sized for.
                request.LabelStyle!,
                alignment,
                @event.Type,
                request.LaneOrder,
                request.StackIndex,
                @event.SortOrder);
            labelBounds = box;
            usedPosition = position;
            if (truncated)
            {
                warnings.Add(new SceneWarning(
                    owner,
                    LabelPlanner.TruncatedToFitCode,
                    "The label was truncated with an ellipsis to fit the available space."));
            }

            return true;
        }

        return false;
    }

    /// <summary>
    /// Places the §23 start-date label at the corner the row's
    /// <see cref="GanttEvent.DateLabelPlacement"/> names, defaulting to
    /// the §24 <c>Auto</c> corner order.
    /// </summary>
    /// <remarks>
    /// The date label shares the description label's corner geometry and
    /// its §24 alignment rule, but never its clipping discipline: a
    /// truncated date is not a date (D-G11), so a corner whose box would
    /// leave the chart — which is precisely a box wider than the available
    /// width — is declined rather than ellipsised. The occupant list
    /// already carries every prior label of the scene plus this row's own
    /// description label, so the two labels of one row never overlap.
    /// </remarks>
    /// <param name="request">The typed delineator request.</param>
    /// <param name="event">The validated delineator event.</param>
    /// <param name="start">The event's start date, the label's text source.</param>
    /// <param name="lineX">The line's exact X.</param>
    /// <param name="descriptionBounds">This row's description-label box, when placed.</param>
    /// <param name="occupants">Already-placed label boxes this label must not intersect.</param>
    /// <param name="stackOffsetPt">The row's stack offset, shared with its description label.</param>
    /// <param name="warnings">The warning list (unused here; suppression is reported by the caller).</param>
    /// <param name="dateLabel">The placed date-label text, or null.</param>
    /// <param name="dateLabelBounds">The placed date-label box, or null.</param>
    /// <returns>True when a corner held the full date.</returns>
    private static bool TryPlaceDateLabel(
        DelineatorRequest request,
        GanttEvent @event,
        DateOnly start,
        double lineX,
        RectD? descriptionBounds,
        IReadOnlyList<RectD> occupants,
        double stackOffsetPt,
        List<SceneWarning> warnings,
        out SceneText? dateLabel,
        out RectD? dateLabelBounds)
    {
        _ = warnings;
        dateLabel = null;
        dateLabelBounds = null;

        var dateText = GanttDateFormatting.Format(start, request.DateFormat);
        if (!request.LabelMetrics.TryMeasure(dateText, out TextMeasurement? measurement)
            || measurement is null)
        {
            return false;
        }

        GanttLabelPosition requested = @event.DateLabelPlacement ?? GanttLabelPosition.Auto;
        IReadOnlyList<GanttLabelPosition> candidates = requested == GanttLabelPosition.Auto
            ? _autoOrder
            : [requested];

        // This row's description label was placed first and blocks the
        // date label: §23 plans a date label against the description
        // label just placed for its row.
        List<RectD> dateOccupants = [.. occupants];
        if (descriptionBounds is { } placedDescription)
        {
            dateOccupants.Add(placedDescription);
        }

        foreach (GanttLabelPosition position in candidates)
        {
            RectD box = CornerBox(
                position,
                lineX,
                measurement,
                request.PlotBounds,
                request.LabelGapPt,
                stackOffsetPt);

            if (!Contains(request.ChartBounds, box) || IntersectsAny(box, dateOccupants))
            {
                continue;
            }

            // §24's alignment rule, shared with the description label:
            // the text hugs the line — a left-hand corner is right-aligned
            // so it ends against the line, a right-hand corner reads away
            // from it.
            GanttTextAlignment alignment = position is GanttLabelPosition.TopLeft
                or GanttLabelPosition.BottomLeft
                ? GanttTextAlignment.Right
                : GanttTextAlignment.Left;

            var owner = SceneOwnerId.ForRow(@event.Id);
            dateLabel = new SceneText(
                ScenePrimitive.CreateId(owner, DateLabelBuilder.StartRole),
                owner,
                ZLayer.DelineatorLabel,
                dateText,
                box,
                // The same predetermined description-label style the
                // description label carries: §23 renders date labels in
                // the body font and default text, and the style carries
                // the font family and size the measuring seam measured at.
                request.LabelStyle!,
                alignment,
                @event.Type,
                request.LaneOrder,
                request.StackIndex,
                @event.SortOrder);
            dateLabelBounds = box;
            return true;
        }

        return false;
    }
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
        if (request.Requests.Any(item =>
            item.PlotBounds != first.PlotBounds
            || item.ChartBounds != first.ChartBounds
            || item.LineBottomPt != first.LineBottomPt
            || item.LineTopPt != first.LineTopPt))
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
            //
            // The description label and the §23 date label share the row's stack
            // slot: they sit at different corners of the same row, so both are
            // placed at the same offset and the slot advances once, by the taller
            // of the two placed boxes.
            var contributesLabel =
                (item.LabelPosition != GanttLabelPosition.None
                    && !string.IsNullOrEmpty(item.Event?.Description))
                || (item.ShowDateLabels
                    && item.LabelStyle is not null
                    && item.Event?.DateLabelPlacement is not GanttLabelPosition.None);

            DelineatorCreationOutcome outcome = DelineatorBuilder.TryBuild(
                item,
                timeScale,
                placed,
                contributesLabel ? stackOffsetPt : 0);

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

            if (!contributesLabel)
            {
                continue;
            }

            if (result.Label is { } text)
            {
                primitives.Add(text);
            }

            if (result.DateLabel is { } dateText)
            {
                primitives.Add(dateText);
            }

            // A placed box reserves its space for every later label,
            // description and date alike. The row's slot advances once,
            // by the taller of its two placed boxes: the description and
            // date labels share the slot at different corners, so the
            // space the row occupies is the taller box, not their sum.
            var placedHeight = 0.0;
            if (result.LabelBounds is { } bounds)
            {
                placed.Add(bounds);
                placedHeight = bounds.Height;
            }

            if (result.DateLabelBounds is { } dateBounds)
            {
                placed.Add(dateBounds);
                placedHeight = Math.Max(placedHeight, dateBounds.Height);
            }

            if (placedHeight > 0)
            {
                // The next label clears this one entirely: the placed box's own
                // height is the space it occupies, and StackGapPt is the space left
                // between them. Advancing by the height rather than by one gap per
                // label is what keeps a stack correct when the gap is smaller than
                // the label height, and for labels of differing heights. The offset
                // is a distance from the plot edge whichever corner the label took,
                // so accumulating it is correct for a top and a bottom label alike.
                stackOffsetPt += placedHeight + request.StackGapPt;
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
