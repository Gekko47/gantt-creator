namespace GanttCreator.Core.Scene;

/// <summary>One label placement request.</summary>
/// <param name="Event">The validated event owning the label.</param>
/// <param name="Text">The label text; blank produces no label.</param>
/// <param name="ResolvedPosition">
/// The position from precedence: the explicit row value, else the named-style
/// default, else <see cref="GanttLabelPosition.Auto"/>.
/// </param>
/// <param name="ShapeBounds">The visible (post-clip) bounds of the owning shape.</param>
/// <param name="TextStyle">The resolved text style used inside the body.</param>
/// <param name="Metrics">The injected deterministic text-metrics seam.</param>
/// <param name="OutsideTextStyle">
/// The resolved text style for a label placed outside the body. §17 requires the
/// delay event to switch from <c>DelayText</c> inside the red body to
/// <c>DefaultText</c> outside it, so this is the style such a label carries. When
/// it is <see langword="null"/>, <paramref name="TextStyle"/> is used.
/// </param>
/// <param name="LaneOrder">The lane ordering value, when known.</param>
/// <param name="StackIndex">The stack ordering value, when known.</param>
public sealed record LabelRequest(
    GanttEvent Event,
    string? Text,
    GanttLabelPosition ResolvedPosition,
    RectD ShapeBounds,
    SceneStyle TextStyle,
    ITextMetrics Metrics,
    SceneStyle? OutsideTextStyle = null,
    int? LaneOrder = null,
    int? StackIndex = null
);

/// <summary>The resolved bounds and metrics a label planner needs.</summary>
/// <param name="PlotBounds">The visible plot bounds; they bound the free space.</param>
/// <param name="ChartBounds">The chart bounds enclosing panels, headers, and plot.</param>
/// <param name="LabelGapPt">The gap between a shape bound and an external label.</param>
/// <param name="LabelHeightPt">The one-line label box height.</param>
/// <param name="MaximumExternalLabelWidthPt">The maximum width of an external label.</param>
public sealed record LabelMetrics(
    RectD PlotBounds,
    RectD ChartBounds,
    double LabelGapPt,
    double LabelHeightPt,
    double MaximumExternalLabelWidthPt
);

/// <summary>The result of planning one label.</summary>
/// <param name="Primitive">The label text primitive, or <see langword="null"/> when the label is suppressed.</param>
/// <param name="Bounds">The resolved label bounds, or <see langword="null"/> when the label is suppressed.</param>
/// <param name="Position">The position actually used, or <see langword="null"/> when the label is suppressed.</param>
/// <param name="WasTruncated">Whether the resolved text was cut with an ellipsis.</param>
/// <param name="Warnings">The deterministic non-blocking warnings for this label.</param>
public sealed record LabelPlanResult(
    SceneText? Primitive,
    RectD? Bounds,
    GanttLabelPosition? Position,
    bool WasTruncated,
    IReadOnlyList<SceneWarning> Warnings
);

/// <summary>The reason a label could not be planned.</summary>
public enum LabelRefusal
{
    /// <summary>The request was null.</summary>
    NullRequest = 0,

    /// <summary>The event or its text style was null.</summary>
    NullDependency = 1,

    /// <summary>The text-metrics seam was absent or could not measure the text.</summary>
    MissingMetrics = 2,

    /// <summary>The metrics record was null.</summary>
    NullMetrics = 3,

    /// <summary>A metric or bound was not finite, or a gap/extent was negative.</summary>
    InvalidMetrics = 4,

    /// <summary>The position was not a defined <see cref="GanttLabelPosition"/> value.</summary>
    InvalidPosition = 5,

    /// <summary>The shape bounds were not finite.</summary>
    InvalidShapeBounds = 6,
}

/// <summary>The typed result of attempting to plan one label.</summary>
/// <param name="Result">The successful result, or <see langword="null"/>.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/>.</param>
public sealed record LabelPlanCreationOutcome(LabelPlanResult? Result, LabelRefusal? Refusal)
{
    /// <summary>Gets whether planning succeeded.</summary>
    public bool Succeeded => Result is not null;
}

/// <summary>
/// Plans deterministic description-label placement, with no Office dependency and
/// no renderer involvement.
/// </summary>
/// <remarks>
/// <para>
/// Precedence and the acceptance rules come from entity guide §22. The cascade
/// order, the delay-event style default, and the blocked-label outcome come from
/// ADR-0015, which amended §12, §17, and §22.
/// </para>
/// <para>
/// The scene resolves the side, the bounds, and the final string, so a renderer
/// never re-measures and never re-selects a label side.
/// </para>
/// </remarks>
public static class LabelPlanner
{
    /// <summary>The warning code emitted when a label is truncated to fit its space.</summary>
    public const string TruncatedToFitCode = "LabelTruncatedToFit";

    /// <summary>The warning code emitted when no measured gap can hold even an ellipsis.</summary>
    public const string SuppressedNoSpaceCode = "LabelSuppressedNoSpace";

    /// <summary>
    /// The span <c>Auto</c> order after ADR-0015: the external sides are tried
    /// before the body, so a fitting <c>Left</c> beats a fitting <c>Inside</c>.
    /// </summary>
    private static readonly GanttLabelPosition[] _spanAutoOrder =
    [
        GanttLabelPosition.Right,
        GanttLabelPosition.Left,
        GanttLabelPosition.Inside,
    ];

    /// <summary>
    /// The delay escape cascade. ADR-0015 D3 excludes <c>Inside</c> so the
    /// position the style default already rejected is never retried.
    /// </summary>
    private static readonly GanttLabelPosition[] _delayAutoOrder = [GanttLabelPosition.Right, GanttLabelPosition.Left];

    /// <summary>
    /// The milestone <c>Auto</c> order (ADR-0015 D2): <c>Right → Left → Above →
    /// Below</c>. <c>Inside</c> is excluded because it is invalid for a milestone
    /// at the initial minimum size, so trying it would be a position the entity
    /// may not use.
    /// </summary>
    private static readonly GanttLabelPosition[] _milestoneAutoOrder =
    [
        GanttLabelPosition.Right,
        GanttLabelPosition.Left,
        GanttLabelPosition.Above,
        GanttLabelPosition.Below,
    ];

    /// <summary>Attempts to plan one label.</summary>
    /// <param name="request">The typed label request.</param>
    /// <param name="metrics">The resolved label bounds and metrics.</param>
    /// <param name="occupants">
    /// The bounds that block a candidate: in-lane shape bounds and the
    /// already-placed higher-priority label bounds. <see cref="GanttLabelPosition.Inside"/>
    /// may occupy its own parent shape and ignores that one entry.
    /// </param>
    /// <returns>A typed result or refusal.</returns>
    public static LabelPlanCreationOutcome TryPlan(LabelRequest? request, LabelMetrics? metrics, IReadOnlyList<RectD>? occupants = null)
    {
        if (request is null)
        {
            return Refused(LabelRefusal.NullRequest);
        }

        if (request.Event is null || request.TextStyle is null)
        {
            return Refused(LabelRefusal.NullDependency);
        }

        if (request.Metrics is null)
        {
            return Refused(LabelRefusal.MissingMetrics);
        }

        if (metrics is null)
        {
            return Refused(LabelRefusal.NullMetrics);
        }

        if (!ValidMetrics(metrics))
        {
            return Refused(LabelRefusal.InvalidMetrics);
        }

        if (!Enum.IsDefined(request.ResolvedPosition))
        {
            return Refused(LabelRefusal.InvalidPosition);
        }

        if (!IsFinite(request.ShapeBounds))
        {
            return Refused(LabelRefusal.InvalidShapeBounds);
        }

        // A blank description is permitted for every type (R2.5 U2) and simply
        // produces no label: neither a warning nor a refusal.
        if (string.IsNullOrWhiteSpace(request.Text))
        {
            return Suppressed();
        }

        if (request.ResolvedPosition == GanttLabelPosition.None)
        {
            return Suppressed();
        }
        var text = request.Text;
        if (!request.Metrics.TryMeasure(text, out TextMeasurement? measured) || measured is null)
        {
            return Refused(LabelRefusal.MissingMetrics);
        }

        _ = measured;
        GanttEvent @event = request.Event;
        IReadOnlyList<RectD> blocked = occupants ?? [];
        var isDelay = @event.Type == GanttEntityType.DelayEvent;
        // §20/ADR-0015 D2: a milestone's Auto order is not the span order, so the
        // cascade is selected from the entity kind rather than assumed.
        var isMilestone = EntityTypeCatalog.GetDefinition(@event.Type)?.Kind == EntityKind.Milestone;

        LabelPlanCreationOutcome Suppressed()
        {
            return new(new LabelPlanResult(null, null, null, false, []), null);
        }

        LabelPlanCreationOutcome Placed(string source, RectD bounds, GanttLabelPosition position, bool truncated)
        {
            // `truncated` is authoritative: only the ADR-0015 widest-gap fallback
            // sets it, and it has already measured the gap. The overflow is
            // deliberately not re-derived here, because the fallback may hand
            // back a box wider than the original text and re-testing the
            // untruncated string against it would report no truncation at all.
            var finalText = truncated ? LabelText.Ellipsize(source, bounds.Width, request.Metrics) : source;
            IReadOnlyList<SceneWarning> warnings = truncated
                ?
                [
                    new SceneWarning(
                        SceneOwnerId.ForRow(@event.Id),
                        TruncatedToFitCode,
                        "The label was truncated with an ellipsis to fit the available space."
                    ),
                ]
                : [];

            // §17: the delay event's text is DelayText inside the red body and
            // DefaultText outside it. The switch is a property of the resolved
            // position, so it is applied here — the single point where the
            // position is final — rather than inside a candidate attempt.
            SceneStyle style = isDelay && position != GanttLabelPosition.Inside
                ? request.OutsideTextStyle ?? request.TextStyle
                : request.TextStyle;

            return new LabelPlanCreationOutcome(
                new LabelPlanResult(
                    new SceneText(
                        $"{@event.Id.Value}:label",
                        SceneOwnerId.ForRow(@event.Id),
                        ZLayer.Label,
                        finalText,
                        bounds,
                        style,
                        position,
                        @event.Type,
                        request.LaneOrder,
                        request.StackIndex,
                        @event.SortOrder
                    ),
                    bounds,
                    position,
                    truncated,
                    warnings
                ),
                null
            );
        }

        // An explicit row value is the single position honoured, whatever the
        // entity family: §22's precedence puts "explicit row value" above both the
        // style default and Auto, so a delay that names a position is never
        // re-placed by the delay default below.
        var hasExplicitPosition = request.ResolvedPosition is not (GanttLabelPosition.Auto or GanttLabelPosition.None);
        GanttLabelPosition styleDefault = request.TextStyle.Alignment ?? GanttLabelPosition.Inside;

        // ADR-0015 D3: the delay event's Inside is a style-level default that is
        // evaluated once and never re-entered. It is the first thing tried, and
        // the delay cascade below deliberately omits Inside. An explicit row
        // position skips this branch entirely, so the style default cannot
        // override what the user asked for.
        if (isDelay && !hasExplicitPosition)
        {
            if (TryAccept(styleDefault, out RectD insideBounds, out _))
            {
                return Placed(text, insideBounds, styleDefault, false);
            }
        }

        GanttLabelPosition[] cascade = hasExplicitPosition
            ? [request.ResolvedPosition]
            : isDelay
                ? _delayAutoOrder
                : isMilestone
                    ? _milestoneAutoOrder
                    : _spanAutoOrder;

        foreach (GanttLabelPosition position in cascade)
        {
            if (TryAccept(position, out RectD bounds, out var truncated))
            {
                return Placed(text, bounds, position, truncated);
            }
        }

        // ADR-0015 D4: no candidate accepted the full text. Measure the free gap
        // at each applicable position and take the widest, truncating to fit.
        return TryWidestGap(text, cascade, out RectD gapBounds, out GanttLabelPosition gapPosition, out var gapText)
            ? Placed(gapText, gapBounds, gapPosition, true)
            : new LabelPlanCreationOutcome(
            new LabelPlanResult(
                null,
                null,
                null,
                false,
                [
                    new SceneWarning(
                        SceneOwnerId.ForRow(request.Event.Id),
                        SuppressedNoSpaceCode,
                        "No available gap could hold the label, so no label was drawn."
                    ),
                ]
            ),
            null
        );
        bool TryAccept(GanttLabelPosition position, out RectD bounds, out bool truncated)
        {
            bounds = default;
            truncated = false;
            if (position is GanttLabelPosition.None or GanttLabelPosition.Auto)
            {
                return false;
            }

            if (!TryGeometry(position, request.ShapeBounds, metrics, blocked, out RectD geometry, out var freeWidth))
            {
                return false;
            }

            var width = Math.Min(freeWidth, metrics.MaximumExternalLabelWidthPt);

            // A cascade candidate must hold the *full* text. A position that
            // cannot is rejected so the cascade continues, and truncation happens
            // only in the ADR-0015 widest-gap fallback, which runs after every
            // candidate has been refused. This subsumes the §22 rule that Inside
            // is allowed only when the text fits the inner bounds, so Inside is
            // never truncated into the body either. `truncated` is only ever set
            // by that fallback.
            if (LabelText.Overflows(text, width, request.Metrics))
            {
                return false;
            }

            // The text fits, so the box hugs the measured text rather than
            // stretching across the whole free gap. Left stays anchored at the
            // shape's near edge and Right at its far edge; the three positions
            // §22 describes as centred (Inside, Above, Below) centre the fitted
            // text inside the shape's own horizontal extent.
            var textWidth = Math.Min(measured.WidthPt, width);
            var left = geometry.X;
            if (position == GanttLabelPosition.Left)
            {
                left = geometry.Right - textWidth;
            }
            else if (position is GanttLabelPosition.Inside or GanttLabelPosition.Above or GanttLabelPosition.Below)
            {
                left = geometry.X + ((geometry.Width - textWidth) / 2);
            }
            var effective = new RectD(left, geometry.Y, textWidth, geometry.Height);

            // Containment is against the chart bounds, which enclose the plot.
            if (!WithinBounds(effective, metrics.ChartBounds) || Blocked(effective, position))
            {
                return false;
            }

            bounds = effective;
            return true;
        }

        bool TryWidestGap(
            string source,
            GanttLabelPosition[] order,
            out RectD widest,
            out GanttLabelPosition widestPosition,
            out string widestText
        )
        {
            widest = default;
            widestPosition = GanttLabelPosition.None;
            widestText = source;
            var best = -1.0;
            var bestBounds = default(RectD);
            GanttLabelPosition bestPosition = GanttLabelPosition.None;

            foreach (GanttLabelPosition position in order)
            {
                if (!TryGeometry(position, request.ShapeBounds, metrics, blocked, out RectD geometry, out var freeWidth))
                {
                    continue;
                }

                // Free space never exceeds the approved external maximum.
                var usable = Math.Min(freeWidth, metrics.MaximumExternalLabelWidthPt);

                // A gap too small to hold even the ellipsis is not usable: the
                // ADR requires suppression, not a label drawn over an obstruction.
                if (
                    !request.Metrics.TryMeasure(LabelText.Ellipsis.ToString(), out TextMeasurement? marker)
                    || marker is null
                    || marker.WidthPt > usable
                )
                {
                    continue;
                }

                // The same anchor rule as TryAccept: Left is anchored by its right
                // edge to the gap boundary, every other position by its left edge.
                // Without this the truncated Left label would sit `freeWidth -
                // usable` to the left of the boundary the untruncated one uses.
                var left = position == GanttLabelPosition.Left ? geometry.Right - usable : geometry.X;
                var candidate = new RectD(left, geometry.Y, usable, geometry.Height);
                if (Blocked(candidate, position))
                {
                    continue;
                }

                // Ties break on the cascade order, so an equal gap keeps the
                // earlier position and placement is stable across refreshes.
                if (usable > best + GeometryMath.Epsilon)
                {
                    best = usable;
                    bestPosition = position;
                    bestBounds = candidate;
                }
            }

            if (bestPosition == GanttLabelPosition.None)
            {
                return false;
            }

            widest = bestBounds;
            widestPosition = bestPosition;
            widestText = LabelText.Ellipsize(source, best, request.Metrics);
            return true;
        }

        bool Blocked(RectD candidate, GanttLabelPosition position)
        {
            foreach (RectD occupant in blocked)
            {
                // An Inside label may occupy its own parent rectangle.
                if (position == GanttLabelPosition.Inside && SameRect(occupant, request.ShapeBounds))
                {
                    continue;
                }

                if (candidate.IntersectsWith(occupant))
                {
                    return true;
                }
            }

            return false;
        }
    }

    /// <summary>
    /// Builds the candidate geometry for a position and measures the free
    /// horizontal space available to it.
    /// </summary>
    /// <remarks>
    /// Free space is the distance from the shape's edge to the nearest
    /// obstruction, where obstructions are in-lane shape bounds, already-placed
    /// higher-priority label bounds, and the plot boundary alike (ADR-0015 D4).
    /// </remarks>
    private static bool TryGeometry(
        GanttLabelPosition position,
        RectD shape,
        LabelMetrics metrics,
        IReadOnlyList<RectD> blocked,
        out RectD bounds,
        out double freeWidth
    )
    {
        bounds = default;
        freeWidth = 0;

        // Every candidate is vertically centred on the shape's own band, and the
        // box height is the one-line label height.
        var top = shape.Top + ((shape.Height - metrics.LabelHeightPt) / 2);
        var height = metrics.LabelHeightPt;

        switch (position)
        {
            case GanttLabelPosition.Right:
                {
                    var left = shape.Right + metrics.LabelGapPt;
                    freeWidth = FreeRight(left, metrics.PlotBounds, blocked);
                    bounds = new RectD(left, top, freeWidth, height);
                    break;
                }

            case GanttLabelPosition.Left:
                {
                    // §22: "Left: label right = shape left − LabelGapPt". The gap
                    // runs leftwards from that edge, so the box is anchored at its
                    // right end; TryAccept then shrinks it to the text width while
                    // keeping that anchor.
                    var right = shape.Left - metrics.LabelGapPt;
                    freeWidth = FreeLeft(right, metrics.PlotBounds, blocked);
                    bounds = new RectD(right - freeWidth, top, freeWidth, height);
                    break;
                }

            case GanttLabelPosition.Inside:
                {
                    // Inside occupies the shape's own interior; the ADR-0015 tie-break
                    // treats it as the narrowest usable position by construction.
                    freeWidth = Math.Max(0, shape.Width);
                    bounds = new RectD(shape.X, top, freeWidth, height);
                    break;
                }

            case GanttLabelPosition.Above:
                {
                    // §22: "Above: horizontally centred; label bottom = shape top −
                    // LabelGapPt". The vertical placement therefore comes from the
                    // shape's top edge, not from the vertically-centred default.
                    freeWidth = Math.Max(0, shape.Width);
                    var aboveTop = shape.Top - metrics.LabelGapPt - metrics.LabelHeightPt;
                    bounds = new RectD(shape.X, aboveTop, freeWidth, height);
                    break;
                }

            case GanttLabelPosition.Below:
                {
                    // §22: "Below: horizontally centred; label top = shape bottom +
                    // LabelGapPt".
                    freeWidth = Math.Max(0, shape.Width);
                    var belowTop = shape.Bottom + metrics.LabelGapPt;
                    bounds = new RectD(shape.X, belowTop, freeWidth, height);
                    break;
                }

            case GanttLabelPosition.None:
            case GanttLabelPosition.Auto:
            case GanttLabelPosition.TopLeft:
            case GanttLabelPosition.TopRight:
            case GanttLabelPosition.BottomLeft:
            case GanttLabelPosition.BottomRight:
            case GanttLabelPosition.DataPanelLeft:
            case GanttLabelPosition.PlotCentre:
            case GanttLabelPosition.Both:
            default:
                // Every other position is either "no label" (None/Auto, handled by
                // the caller) or belongs to a different entity family: the
                // delineator corners, the splitter placements, and the date-label
                // positions are owned by R3.10 and R3.11 respectively, and are
                // refused here rather than silently rendered as a span label.
                return false;
        }

        return freeWidth > GeometryMath.Epsilon;
    }

    /// <summary>Measures the free space running right from a left edge to the nearest obstruction.</summary>
    private static double FreeRight(double left, RectD plot, IReadOnlyList<RectD> blocked)
    {
        var limit = plot.Right;
        foreach (RectD occupant in blocked)
        {
            // Only an occupant that starts to the right of the candidate limits
            // it; anything else is irrelevant to this direction.
            if (occupant.Left > left + GeometryMath.Epsilon)
            {
                limit = Math.Min(limit, occupant.Left);
            }
        }

        return Math.Max(0, limit - left);
    }

    /// <summary>
    /// Returns whether a candidate rectangle lies entirely within the allowed
    /// chart bounds, which is the §22 containment rule.
    /// </summary>
    private static bool WithinBounds(RectD candidate, RectD bounds) =>
        candidate.Left >= bounds.Left - GeometryMath.Epsilon
        && candidate.Top >= bounds.Top - GeometryMath.Epsilon
        && candidate.Right <= bounds.Right + GeometryMath.Epsilon
        && candidate.Bottom <= bounds.Bottom + GeometryMath.Epsilon;

    /// <summary>Measures the free space running left from a right edge to the nearest obstruction.</summary>
    private static double FreeLeft(double right, RectD plot, IReadOnlyList<RectD> blocked)
    {
        var limit = plot.Left;
        foreach (RectD occupant in blocked)
        {
            if (occupant.Right < right - GeometryMath.Epsilon)
            {
                limit = Math.Max(limit, occupant.Right);
            }
        }

        return Math.Max(0, right - limit);
    }

    private static bool IsFinite(RectD rect) =>
        double.IsFinite(rect.X) && double.IsFinite(rect.Y) && double.IsFinite(rect.Width) && double.IsFinite(rect.Height);

    private static bool SameRect(RectD left, RectD right) =>
        GeometryMath.ApproximatelyEqual(left.X, right.X)
        && GeometryMath.ApproximatelyEqual(left.Y, right.Y)
        && GeometryMath.ApproximatelyEqual(left.Width, right.Width)
        && GeometryMath.ApproximatelyEqual(left.Height, right.Height);

    private static bool ValidMetrics(LabelMetrics metrics) =>
        IsFinite(metrics.PlotBounds)
        && IsFinite(metrics.ChartBounds)
        && IsFiniteNonNegative(metrics.LabelGapPt)
        && IsFiniteNonNegative(metrics.LabelHeightPt)
        && IsFiniteNonNegative(metrics.MaximumExternalLabelWidthPt);

    private static bool IsFiniteNonNegative(double value) => double.IsFinite(value) && value >= 0;

    private static LabelPlanCreationOutcome Refused(LabelRefusal refusal) => new(null, refusal);
}
