namespace GanttCreator.Core.Scene;

/// <summary>One critical-interval overlay request.</summary>
/// <param name="Event">The validated critical-interval child event.</param>
/// <param name="Style">The resolved scene style for the overlay line.</param>
/// <param name="ParentVisibleBounds">
/// The parent's **visible (post-plot-clip)** bar bounds, as produced by
/// <see cref="SpanBarResult.VisibleBounds"/>. The overlay clips to these, never
/// to the parent's pre-clip dates.
/// </param>
/// <param name="CriticalLinePt">The overlay thickness in points.</param>
/// <param name="LaneOrder">The lane ordering value, when known.</param>
/// <param name="StackIndex">The stack ordering value, when known.</param>
public sealed record CriticalOverlayRequest(
    GanttEvent Event,
    SceneStyle Style,
    RectD ParentVisibleBounds,
    double CriticalLinePt,
    int? LaneOrder = null,
    int? StackIndex = null
);

/// <summary>The result of building one critical-interval overlay.</summary>
/// <param name="Primitive">The overlay line, or <see langword="null"/> when the interval lies wholly outside its parent.</param>
/// <param name="VisibleBounds">The clipped overlay bounds, or <see langword="null"/> when no overlay was emitted.</param>
/// <param name="WasClipped">Whether a portion of the interval fell outside the parent's visible span.</param>
/// <param name="Warnings">The deterministic non-blocking warnings for this overlay.</param>
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

    /// <summary>The parent bounds or the thickness was not finite, or the thickness was not positive.</summary>
    InvalidGeometry = 3,

    /// <summary>The event is not a critical interval.</summary>
    NotACriticalInterval = 4,

    /// <summary>
    /// The child carried no parent reference, or the parent bar was not
    /// resolvable. R2.5 blocks this at validation; the refusal is defence in
    /// depth so an orphan never throws at scene-build time.
    /// </summary>
    UnresolvedParent = 5,

    /// <summary>The child interval carried no ordered start/finish pair.</summary>
    InvalidInterval = 6,
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
/// Builds deterministic critical-interval overlays that sit on their parent's
/// visible bar span, with no Office dependency and no text measurement.
/// </summary>
public static class CriticalOverlayBuilder
{
    /// <summary>The warning code emitted when an overlay is clipped to its parent's visible span.</summary>
    public const string ClippedToParentCode = "CriticalIntervalClippedToParent";

    /// <summary>The warning code emitted when an interval lies wholly outside its parent.</summary>
    public const string OutsideParentCode = "CriticalIntervalOutsideParent";

    /// <summary>Attempts to build one critical-interval overlay.</summary>
    /// <param name="request">The typed overlay request.</param>
    /// <param name="timeScale">The validated time scale.</param>
    /// <param name="parentVisibleBounds">
    /// The parent bar's visible bounds, keyed by parent row ID. R3.6 produces
    /// these after plot clipping, which is what the overlay must clip to.
    /// </param>
    /// <returns>A typed result or refusal.</returns>
    public static CriticalOverlayCreationOutcome TryBuild(
        CriticalOverlayRequest? request,
        TimeScale? timeScale,
        IReadOnlyDictionary<GanttRowId, RectD>? parentVisibleBounds
    )
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

        if (
            !double.IsFinite(request.CriticalLinePt)
            || request.CriticalLinePt <= 0
            || !double.IsFinite(request.ParentVisibleBounds.X)
            || !double.IsFinite(request.ParentVisibleBounds.Width)
        )
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

        if (@event.ParentId is not { } parentId)
        {
            return Refused(CriticalOverlayRefusal.UnresolvedParent);
        }

        // The parent bar's *visible* bounds are authoritative, so a parent that
        // was itself plot-clipped constrains the overlay to what is on screen.
        if (parentVisibleBounds is null || !parentVisibleBounds.TryGetValue(parentId, out RectD parent))
        {
            return Refused(CriticalOverlayRefusal.UnresolvedParent);
        }

        if (@event.Start is not { } start || @event.Finish is not { } finish || finish < start)
        {
            return Refused(CriticalOverlayRefusal.InvalidInterval);
        }

        List<SceneWarning> warnings = [];

        // The child span is mapped with the inclusive rule and then intersected
        // with the parent's visible span. The `Try*` forms are used throughout:
        // DateToX throws for an out-of-range date, and a clipped interval must
        // warn rather than throw.
        var hasStart = timeScale.TryDateToX(start, out var startX);
        var hasFinish = timeScale.TryDurationDays(start, finish, out var durationDays);
        var ownRight = hasFinish ? startX + (durationDays * timeScale.DayWidth) : startX;
        var ownLeft = hasStart ? startX : timeScale.PlotLeftPt;

        var clippedLeft = Math.Max(Math.Min(ownLeft, ownRight), parent.Left);
        var clippedRight = Math.Min(Math.Max(ownLeft, ownRight), parent.Right);

        if (clippedRight <= clippedLeft + GeometryMath.Epsilon)
        {
            warnings.Add(
                new SceneWarning(
                    SceneOwnerId.ForRow(@event.Id),
                    OutsideParentCode,
                    "The critical interval lies wholly outside its parent span, so no overlay was drawn.")
            );
            return new CriticalOverlayCreationOutcome(
                new CriticalOverlayResult(null, null, true, warnings),
                null
            );
        }

        var wasClipped = ownLeft < parent.Left - GeometryMath.Epsilon || ownRight > parent.Right + GeometryMath.Epsilon;
        if (wasClipped)
        {
            warnings.Add(
                new SceneWarning(
                    SceneOwnerId.ForRow(@event.Id),
                    ClippedToParentCode,
                    "The critical interval was clipped to its parent's visible span.")
            );
        }

        // §16: a solid CriticalStroke line along the parent's top edge with its
        // centreline at parent.Top + CriticalLinePt/2, which keeps the whole
        // stroke inside the body. Thickness is CriticalLinePt.
        var overlay = new RectD(
            clippedLeft,
            parent.Top,
            clippedRight - clippedLeft,
            request.CriticalLinePt
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

    private static CriticalOverlayCreationOutcome Refused(CriticalOverlayRefusal refusal) => new(null, refusal);
}
