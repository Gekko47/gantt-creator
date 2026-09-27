namespace GanttCreator.Core.Scene;

/// <summary>One span bar request: the validated event, its resolved style, and the slot it occupies.</summary>
/// <param name="Event">The validated span event.</param>
/// <param name="Style">The resolved scene style for the bar body.</param>
/// <param name="SlotCentreY">The vertical centre of the visual stack slot, in points.</param>
/// <param name="ResolvedHeightPt">The resolved activity height, in points.</param>
/// <param name="LaneOrder">The lane ordering value, when known.</param>
/// <param name="StackIndex">The stack ordering value, when known.</param>
public sealed record SpanBarRequest(
    GanttEvent Event,
    SceneStyle Style,
    double SlotCentreY,
    double ResolvedHeightPt,
    int? LaneOrder = null,
    int? StackIndex = null
);

/// <summary>The result of building one span bar.</summary>
/// <param name="Primitive">The bar rectangle, or <see langword="null"/> when the span lies wholly outside the plot.</param>
/// <param name="VisibleBounds">The clipped visible bounds, or <see langword="null"/> when no bar was emitted.</param>
/// <param name="Warnings">The deterministic non-blocking warnings for this bar.</param>
public sealed record SpanBarResult(SceneRect? Primitive, RectD? VisibleBounds, IReadOnlyList<SceneWarning> Warnings);

/// <summary>The reason a span bar could not be built.</summary>
public enum SpanBarRefusal
{
    /// <summary>The request was null.</summary>
    NullRequest = 0,

    /// <summary>The event, its style, or its owner was null.</summary>
    NullDependency = 1,

    /// <summary>The time scale was invalid.</summary>
    InvalidTimeScale = 2,

    /// <summary>The slot centre or resolved height was not finite, or the height was negative.</summary>
    InvalidGeometry = 3,

    /// <summary>The event is not a span: it has no start, or no finish after its start.</summary>
    NotASpan = 4,
}

/// <summary>The typed result of attempting to build a span bar.</summary>
/// <param name="Result">The successful result, or <see langword="null"/>.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/>.</param>
public sealed record SpanBarCreationOutcome(SpanBarResult? Result, SpanBarRefusal? Refusal)
{
    /// <summary>Gets whether construction succeeded.</summary>
    public bool Succeeded => Result is not null;
}

/// <summary>
/// Builds deterministic span-activity bar geometry, with no Office dependency
/// and no text measurement.
/// </summary>
public static class SpanBarBuilder
{
    /// <summary>The warning code emitted when a span is clipped at a plot edge.</summary>
    public const string ClippedToPlotCode = "SpanClippedToPlot";

    /// <summary>The warning code emitted when a span lies wholly outside the plot.</summary>
    public const string OutsidePlotRangeCode = "SpanOutsidePlotRange";

    /// <summary>Attempts to build one span bar.</summary>
    /// <param name="request">The typed span bar request.</param>
    /// <param name="timeScale">The validated time scale.</param>
    /// <returns>A typed result or refusal.</returns>
    public static SpanBarCreationOutcome TryBuild(SpanBarRequest? request, TimeScale? timeScale)
    {
        if (request is null)
        {
            return Refused(SpanBarRefusal.NullRequest);
        }

        if (timeScale is null)
        {
            return Refused(SpanBarRefusal.InvalidTimeScale);
        }

        if (request.Event is null || request.Style is null)
        {
            return Refused(SpanBarRefusal.NullDependency);
        }

        if (!double.IsFinite(request.SlotCentreY) || !double.IsFinite(request.ResolvedHeightPt) || request.ResolvedHeightPt < 0)
        {
            return Refused(SpanBarRefusal.InvalidGeometry);
        }

        if (request.Event.Start is not { } start)
        {
            return Refused(SpanBarRefusal.NotASpan);
        }

        GanttEvent @event = request.Event;
        var ownerId = SceneOwnerId.ForRow(@event.Id);
        List<SceneWarning> warnings = [];

        // A milestone or a dateless type reaching this builder is a scene
        // construction error, not a silently empty result: the inclusive span
        // rule needs both dates and the guide forbids inventing one. A finish
        // before the start is refused for the same reason — it is not a span, and
        // letting it reach the clip branch would let a reversed range produce a
        // bar from two collapsed edges.
        if (@event.Finish is not { } finish || finish < start)
        {
            return Refused(SpanBarRefusal.NotASpan);
        }

        // Entity guide §12: the inclusive span rule. `Finish` is never
        // incremented, and a one-day activity keeps exactly one day of width.
        if (!timeScale.TryDurationDays(start, finish, out var durationDays))
        {
            // The dates are ordered but at least one falls outside the scale, so
            // the span is clipped rather than refused. The `Try*` form is used
            // deliberately: DateToX throws for an out-of-range date, and a
            // clipped span must still produce a warning rather than an exception.
            (double Left, double Right)? visible = ClipToPlot(start, finish, timeScale);
            if (visible is not { } clipped)
            {
                warnings.Add(
                    new SceneWarning(ownerId, OutsidePlotRangeCode, "The activity lies wholly outside the plot range, so no bar was drawn.")
                );
                return new SpanBarCreationOutcome(new SpanBarResult(null, null, warnings), null);
            }

            RectD clippedBounds = BuildBounds(clipped.Left, clipped.Right, request.SlotCentreY, request.ResolvedHeightPt);
            warnings.Add(new SceneWarning(ownerId, ClippedToPlotCode, "The activity was clipped to the visible plot range."));
            return new SpanBarCreationOutcome(
                new SpanBarResult(CreatePrimitive(@event, request, clippedBounds), clippedBounds, warnings),
                null
            );
        }

        // TimeScale maps a date to the start of that day, so a left edge is the
        // start X and the width is the inclusive duration in day widths.
        var left = timeScale.DateToX(start);
        var right = left + (durationDays * timeScale.DayWidth);
        RectD unclipped = BuildBounds(left, right, request.SlotCentreY, request.ResolvedHeightPt);

        // The scale's own plot edges already bound any in-range span, so a bar
        // whose dates are inside the range cannot leave the plot. The clip is
        // therefore applied uniformly and the visible bounds are the bar bounds.
        RectD bounds = unclipped;
        return new SpanBarCreationOutcome(new SpanBarResult(CreatePrimitive(@event, request, bounds), bounds, warnings), null);
    }

    private static (double Left, double Right)? ClipToPlot(DateOnly start, DateOnly finish, TimeScale timeScale)
    {
        // A point-event style mapping is used deliberately instead of DateToX:
        // the range test is the branch, and the failure sentinel (x = 0) is
        // never inspected. An out-of-range edge collapses onto the plot boundary
        // and gains no day width, so a wholly-outside span keeps zero width
        // rather than a phantom single day.
        var left = ClampToScale(start, timeScale);
        var right = ClampFinishToScale(finish, timeScale);
        var clippedLeft = Math.Max(left, timeScale.PlotLeftPt);
        var clippedRight = Math.Min(right, timeScale.PlotRightPt);
        return clippedRight > clippedLeft + GeometryMath.Epsilon ? (clippedLeft, clippedRight) : null;
    }

    /// <summary>
    /// Maps an inclusive finish date to the exclusive right edge, adding one day
    /// width only when the finish is itself inside the scale. A finish before the
    /// plot contributes no width, which is what makes a wholly-outside span
    /// measure zero.
    /// </summary>
    private static double ClampFinishToScale(DateOnly date, TimeScale timeScale) =>
        date < timeScale.PlotStart || date > timeScale.PlotFinish
            ? ClampToScale(date, timeScale)
            : timeScale.DateToX(date) + timeScale.DayWidth;

    private static double ClampToScale(DateOnly date, TimeScale timeScale)
    {
        return date < timeScale.PlotStart
            ? timeScale.PlotLeftPt
            : date > timeScale.PlotFinish ? timeScale.PlotRightPt : timeScale.DateToX(date);
    }

    private static RectD BuildBounds(double left, double right, double centreY, double heightPt) =>
        new(left, centreY - (heightPt / 2), Math.Max(0, right - left), heightPt);

    private static SceneRect CreatePrimitive(GanttEvent @event, SpanBarRequest request, RectD bounds) =>
        new(
            $"{@event.Id.Value}:bar",
            SceneOwnerId.ForRow(@event.Id),
            ZLayer.ActivityBody,
            bounds,
            request.Style,
            @event.Type,
            request.LaneOrder,
            request.StackIndex,
            @event.SortOrder
        );

    private static SpanBarCreationOutcome Refused(SpanBarRefusal refusal) => new(null, refusal);
}
