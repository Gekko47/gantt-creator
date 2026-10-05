namespace GanttCreator.Core.Scene;

/// <summary>One critical-interval bar request.</summary>
/// <param name="Event">The validated critical-interval child event.</param>
/// <param name="Style">The resolved scene style for the bar.</param>
/// <param name="PredeterminedHeightPt">
/// The bar's <em>predetermined</em> height in points: the resolved style's
/// <c>ActivityHeightPt</c>. The drawn height is half of it, so there is no
/// thickness token to configure (ADR-0027 D4, owner ruling 2026-09-30).
/// </param>
/// <param name="SlotCentreY">
/// The centre Y of the visual slot this interval occupies, from
/// <c>LaneEventLayout</c>. The bar is top-aligned to the top a full-height
/// activity in this slot would have (ADR-0034 D3, owner ruling 2026-10-02),
/// so the critical interval needs nothing from its parent to be placed.
/// </param>
/// <param name="LaneOrder">The lane ordering value, when known.</param>
/// <param name="StackIndex">The stack ordering value, when known.</param>
/// <remarks>
/// <para>
/// <b>There is deliberately no parent in this request</b> (owner ruling
/// 2026-09-30). The critical interval is an ordinary span: its horizontal extent
/// comes from its <em>own</em> Start and Finish, and it may be shorter than,
/// equal to, or longer than its parent's span. Clipping it to the parent was a
/// guide rule that made the entity's geometry depend on a row it did not own.
/// </para>
/// <para>
/// The parent link still exists in the <em>hierarchy</em> — R4.7B projects the
/// child onto the parent's lane, which is what <paramref name="SlotCentreY"/>
/// already reflects. Sharing a lane is a layout fact, not a geometry dependency.
/// </para>
/// </remarks>
public sealed record CriticalOverlayRequest(
    GanttEvent Event,
    SceneStyle Style,
    double PredeterminedHeightPt,
    double SlotCentreY,
    int? LaneOrder = null,
    int? StackIndex = null
);

/// <summary>The result of building one critical-interval bar.</summary>
/// <param name="Primitive">The bar, or <see langword="null"/> when the interval lies wholly outside the plot.</param>
/// <param name="VisibleBounds">The drawn bounds, or <see langword="null"/> when no bar was emitted.</param>
/// <param name="WasClipped">Whether a portion of the interval fell outside the plot's visible span.</param>
/// <param name="Warnings">The deterministic non-blocking warnings for this bar.</param>
public sealed record CriticalOverlayResult(
    SceneRect? Primitive,
    RectD? VisibleBounds,
    bool WasClipped,
    IReadOnlyList<SceneWarning> Warnings
);

/// <summary>The reason a critical overlay could not be built.</summary>
public enum CriticalOverlayRefusal
{
    /// <summary>The request was null.</summary>
    NullRequest = 0,

    /// <summary>The event or its style was null.</summary>
    NullDependency = 1,

    /// <summary>The time scale was invalid.</summary>
    InvalidTimeScale = 2,

    /// <summary>
    /// The predetermined height, or the visual slot's centre Y, was not finite, or
    /// the predetermined height was not positive. Nothing about a parent is
    /// validated here because nothing about a parent is used.
    /// </summary>
    InvalidGeometry = 3,

    /// <summary>The event is not a critical interval.</summary>
    NotACriticalInterval = 4,

    /// <summary>The child interval carried no ordered start/finish pair.</summary>
    InvalidInterval = 5,
}

/// <summary>The typed result of attempting to build a critical overlay.</summary>
/// <param name="Result">The successful result, or <see langword="null"/>.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/>.</param>
public sealed record CriticalOverlayCreationOutcome(CriticalOverlayResult? Result, CriticalOverlayRefusal? Refusal)
{
    /// <summary>Gets whether construction succeeded.</summary>
    public bool Succeeded => Result is not null;
}

/// <summary>
/// Builds deterministic critical-interval bars from the interval's OWN dates and
/// its own visual slot, with no Office dependency and no text measurement.
/// </summary>
/// <remarks>
/// The parent plays no part in the geometry. See
/// <see cref="CriticalOverlayRequest"/> for why.
/// </remarks>
public static class CriticalOverlayBuilder
{
    /// <summary>The warning code emitted when a bar is clipped to the plotted range.</summary>
    public const string ClippedToPlotCode = "CriticalIntervalClippedToPlot";

    /// <summary>The warning code emitted when an interval lies wholly outside the plotted range.</summary>
    public const string OutsidePlotCode = "CriticalIntervalOutsidePlot";

    /// <summary>Attempts to build one critical-interval bar.</summary>
    /// <param name="request">The typed bar request.</param>
    /// <param name="timeScale">The validated time scale.</param>
    /// <returns>A typed result or refusal.</returns>
    public static CriticalOverlayCreationOutcome TryBuild(
        CriticalOverlayRequest? request,
        TimeScale? timeScale)
    {
        if (request is null)
        {
            return Refused(CriticalOverlayRefusal.NullRequest);
        }

        if (timeScale is null)
        {
            return Refused(CriticalOverlayRefusal.InvalidTimeScale);
        }

        if (request.Event is null || request.Style is null)
        {
            return Refused(CriticalOverlayRefusal.NullDependency);
        }

        // Nothing about the parent is validated here, because nothing about the
        // parent is used. A valid token plus a finite slot centre is sufficient to
        // draw the bar, which is what "independent" has to mean operationally.
        if (!double.IsFinite(request.PredeterminedHeightPt)
            || request.PredeterminedHeightPt <= 0
            || !double.IsFinite(request.SlotCentreY))
        {
            return Refused(CriticalOverlayRefusal.InvalidGeometry);
        }

        GanttEvent @event = request.Event;
        // A critical interval is a Span-kind row, so the type identity — not the
        // kind — is what distinguishes it from a plain activity.
        if (EntityTypeCatalog.GetDefinition(@event.Type) is null || @event.Type != GanttEntityType.CriticalInterval)
        {
            return Refused(CriticalOverlayRefusal.NotACriticalInterval);
        }

        if (@event.Start is not { } start || @event.Finish is not { } finish || finish < start)
        {
            return Refused(CriticalOverlayRefusal.InvalidInterval);
        }

        List<SceneWarning> warnings = [];

        // The interval is mapped with the inclusive rule and clamped to the plot,
        // exactly as any other span bar is. There is deliberately NO second,
        // parent-derived clip (owner ruling 2026-09-30): the critical interval may
        // be shorter than, equal to, or longer than its parent, and its horizontal
        // extent comes from its own dates alone. The `Try*` forms are used
        // throughout because DateToX throws for an out-of-range date, and a clipped
        // interval must warn rather than throw.
        var ownLeft = ClampStartToScale(start, timeScale);
        var ownRight = ClampFinishToScale(finish, timeScale);

        if (ownRight <= ownLeft + GeometryMath.Epsilon)
        {
            warnings.Add(
                new SceneWarning(
                    SceneOwnerId.ForRow(@event.Id),
                    OutsidePlotCode,
                    "The critical interval lies wholly outside the plotted range, so no bar was drawn."));
            return new CriticalOverlayCreationOutcome(
                new CriticalOverlayResult(null, null, true, warnings),
                null
            );
        }

        // Clipping is reported against the PLOT only. The old parent-clip warning
        // had no meaning here: there is no longer a parent edge to be clipped to.
        var wasClipped = (start < timeScale.PlotStart && ownLeft <= timeScale.PlotLeftPt + GeometryMath.Epsilon)
            || (finish > timeScale.PlotFinish && ownRight >= timeScale.PlotRightPt - GeometryMath.Epsilon);
        if (wasClipped)
        {
            warnings.Add(
                new SceneWarning(
                    SceneOwnerId.ForRow(@event.Id),
                    ClippedToPlotCode,
                    "The critical interval was clipped to the plotted range."));
        }

        // ADR-0027 D1 as AMENDED by ADR-0034 D3, owner ruling 2026-10-02: the rect is
        // what is drawn, it is filled, its height is HALF the predetermined
        // ActivityHeightPt, and it is TOP-ALIGNED with the top of a normal activity
        // shape in its slot rather than centred on the slot.
        //
        // It was centred (`SlotCentreY - height/2`), which sat it in the middle of the
        // parent bar. The owner ruled that a critical interval is OFF-CENTRE and shares
        // the activity shape's top edge, so the two read as one band when the child
        // overlays its parent. Expressed as "the top a full-height activity in this
        // slot would have" rather than as `SlotCentreY - (predetermined/2)` so it stays
        // true when the slot is taller than the activity -- the interval then keeps the
        // top edge the activity would have had instead of drifting with the slot.
        var height = request.PredeterminedHeightPt / 2;
        var activityTop = request.SlotCentreY - (request.PredeterminedHeightPt / 2);
        var overlay = new RectD(
            ownLeft,
            activityTop,
            ownRight - ownLeft,
            height
        );

        var primitive = new SceneRect(
            $"{@event.Id.Value}:critical",
            SceneOwnerId.ForRow(@event.Id),
            ZLayer.CriticalOverlay,
            overlay,
            request.Style,
            @event.Type,
            request.LaneOrder,
            request.StackIndex,
            @event.SortOrder
        );

        return new CriticalOverlayCreationOutcome(
            new CriticalOverlayResult(primitive, overlay, wasClipped, warnings),
            null
        );
    }

    /// <summary>
    /// Maps an interval start to its start-of-day X, collapsing a date before the
    /// plot onto the plot's left edge so a child crossing the left boundary keeps
    /// the portion that is visible.
    /// </summary>
    /// <param name="date">The interval start date.</param>
    /// <param name="timeScale">The validated time scale.</param>
    /// <returns>The start X, clamped to the plot.</returns>
    private static double ClampStartToScale(DateOnly date, TimeScale timeScale) =>
        timeScale.TryDateToX(date, out var x)
            ? x
            : date < timeScale.PlotStart ? timeScale.PlotLeftPt : timeScale.PlotRightPt;

    /// <summary>
    /// Maps an inclusive finish date to the exclusive right edge, adding one day
    /// width only when the finish is itself inside the scale, and collapsing a
    /// finish after the plot onto the plot's right edge.
    /// </summary>
    /// <param name="date">The interval finish date.</param>
    /// <param name="timeScale">The validated time scale.</param>
    /// <returns>The right edge, clamped to the plot.</returns>
    private static double ClampFinishToScale(DateOnly date, TimeScale timeScale) =>
        date < timeScale.PlotStart || date > timeScale.PlotFinish
            ? ClampStartToScale(date, timeScale)
            : timeScale.DateToX(date) + timeScale.DayWidth;

    private static CriticalOverlayCreationOutcome Refused(CriticalOverlayRefusal refusal) => new(null, refusal);
}
