namespace GanttCreator.Core.Scene;

/// <summary>Which section 23 date labels a request should plan.</summary>
/// <param name="Event">
/// The validated event. A <see langword="null"/> <see cref="GanttEvent.Start"/>
/// or <see cref="GanttEvent.Finish"/> already means the entity type does not use
/// that field, so section 23's "no label for a field the event type does not use"
/// rule needs no catalogue lookup.
/// </param>
/// <param name="VisibleBounds">The bar's post-plot-clip bounds; section 23 anchors labels to the visible bar.</param>
/// <param name="FullBounds">
/// The bar's unclipped bounds. Comparing the two derives whether the entity is
/// plot-clipped, which drives the never-suppress rule. This is geometry, not a
/// caller-asserted flag, so the rule stays resolved in the scene.
/// </param>
/// <param name="Metrics">The plot and chart bounds plus the label metrics.</param>
/// <param name="TextMetrics">The injected deterministic text-metrics seam.</param>
/// <param name="DateFormat">The ADR-0016 approved invariant display format.</param>
/// <param name="TextStyle">The resolved date-label text style.</param>
/// <param name="ShowStart">Whether the start-date label is requested.</param>
/// <param name="ShowFinish">Whether the finish-date label is requested.</param>
/// <param name="StartPosition">The explicit start-label position, else <see cref="GanttLabelPosition.Auto"/>.</param>
/// <param name="FinishPosition">The explicit finish-label position, else <see cref="GanttLabelPosition.Auto"/>.</param>
/// <param name="Occupants">
/// Already-placed label boxes the date labels must not intersect — the row's own
/// description label and any higher-priority label, exactly the bounds the caller
/// passes to every other <see cref="LabelPlanner"/> caller. Without them the
/// planner is asked to place each date label in an empty world, so §23's
/// "collision with the description label" rule could never fire and overlapping
/// labels would be emitted rather than resolved.
/// </param>
public sealed record DateLabelRequest(
    GanttEvent Event,
    RectD VisibleBounds,
    RectD FullBounds,
    LabelMetrics Metrics,
    ITextMetrics TextMetrics,
    GanttDateDisplayFormat DateFormat,
    SceneStyle TextStyle,
    bool ShowStart = true,
    bool ShowFinish = true,
    GanttLabelPosition? StartPosition = null,
    GanttLabelPosition? FinishPosition = null,
    IReadOnlyList<RectD>? Occupants = null
);

/// <summary>The reason a date-label build was refused.</summary>
public enum DateLabelRefusal
{
    /// <summary>The request was null.</summary>
    NullRequest = 0,

    /// <summary>The event was null.</summary>
    NullEvent = 1,

    /// <summary>The metrics were null.</summary>
    NullMetrics = 2,

    /// <summary>The text-metrics seam was null.</summary>
    NullTextMetrics = 3,

    /// <summary>The text style was null.</summary>
    NullTextStyle = 4,

    /// <summary>The event type was not a defined value.</summary>
    InvalidEntityType = 5,

    /// <summary>The visible or full bounds were not finite, or had a non-positive size.</summary>
    InvalidBounds = 6,

    /// <summary>The planner refused the request rather than declining it.</summary>
    PlannerRefused = 7,
}

/// <summary>The planned date labels and any labels the planner declined.</summary>
/// <param name="Primitives">The emitted date-label text primitives, in start-then-finish order.</param>
/// <param name="Suppressed">
/// The roles that were requested but not emitted. Section 23's never-suppress
/// rule guarantees this is empty for a plot-clipped entity.
/// </param>
public sealed record DateLabelResult(
    IReadOnlyList<SceneText> Primitives,
    IReadOnlyList<string> Suppressed);

/// <summary>The typed result of building the date labels.</summary>
/// <param name="Result">The planned labels, or <see langword="null"/>.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/>.</param>
public sealed record DateLabelOutcome(DateLabelResult? Result, DateLabelRefusal? Refusal)
{
    /// <summary>Gets whether the labels were planned.</summary>
    public bool Succeeded => Result is not null;
}
/// <summary>
/// Plans section 23's start-date and finish-date labels through the shared
/// <see cref="LabelPlanner"/>.
/// </summary>
/// <remarks>
/// <para>
/// The two labels are separate scene text entities with role-derived IDs
/// (<c>&lt;rowId&gt;:date-start</c> and <c>&lt;rowId&gt;:date-finish</c>), so a
/// row's description label and its two date labels never collide (R3.11 D4).
/// </para>
/// <para>
/// <b>The never-suppress rule.</b> Section 23 (revision 5, product owner
/// 2026-09-26, D-G11) requires that a plot-clipped event always displays its
/// <em>true</em> date, and that the off-plot date is never suppressed. The
/// label text is therefore always the real date, never a clamped or elided one,
/// and it is anchored to the bar's <em>visible</em> bounds. When the planner
/// cannot place such a label, this builder falls back to a deterministic
/// unclamped box beside the visible edge so the date is still emitted. The rule
/// is resolved here in the scene and is never re-derived per renderer.
/// </para>
/// </remarks>
public static class DateLabelBuilder
{
    /// <summary>The start-date label role.</summary>
    public const string StartRole = "date-start";

    /// <summary>The finish-date label role.</summary>
    public const string FinishRole = "date-finish";

    /// <summary>Attempts to plan both date labels.</summary>
    /// <param name="request">The typed date-label request.</param>
    /// <returns>A typed result or refusal.</returns>
    public static DateLabelOutcome TryBuild(DateLabelRequest? request)
    {
        if (request is null)
        {
            return Refused(DateLabelRefusal.NullRequest);
        }

        if (request.Event is null)
        {
            return Refused(DateLabelRefusal.NullEvent);
        }

        if (request.Metrics is null)
        {
            return Refused(DateLabelRefusal.NullMetrics);
        }

        if (request.TextMetrics is null)
        {
            return Refused(DateLabelRefusal.NullTextMetrics);
        }

        if (request.TextStyle is null)
        {
            return Refused(DateLabelRefusal.NullTextStyle);
        }

        if (!Enum.IsDefined(request.Event.Type))
        {
            return Refused(DateLabelRefusal.InvalidEntityType);
        }

        if (!IsPositive(request.VisibleBounds) || !IsPositive(request.FullBounds))
        {
            return Refused(DateLabelRefusal.InvalidBounds);
        }

        // Geometry-derived, not caller-asserted: a narrower visible span than
        // full span means the plot clipped the entity.
        var clipped = request.VisibleBounds != request.FullBounds;
        List<SceneText> emitted = [];
        List<string> suppressed = [];

        // The caller's already-placed labels block both date labels, and the
        // start label this builder places then blocks the finish label, so the
        // two dates of one row are planned against each other rather than each
        // being planned in isolation and overlapping.
        List<RectD> occupied = [.. request.Occupants ?? []];

        DateLabelRefusal? refusal = null;
        if (request.ShowStart && request.Event.Start is { } start)
        {
            refusal = Plan(request, start, StartRole, request.StartPosition, clipped, occupied, emitted, suppressed);
        }

        if (refusal is null && request.ShowFinish && request.Event.Finish is { } finish)
        {
            refusal = Plan(request, finish, FinishRole, request.FinishPosition, clipped, occupied, emitted, suppressed);
        }

        // A planner refusal is a broken dependency, not a placement decision, so
        // it is never reported as a merely missing date.
        return refusal is { } reason ? Refused(reason) : new DateLabelOutcome(new DateLabelResult(emitted, suppressed), null);
    }
    /// <summary>
    /// Plans one date label.
    /// </summary>
    /// <returns>
    /// A refusal when the planner itself refused; otherwise <see langword="null"/>,
    /// with the outcome recorded in <paramref name="emitted"/> or
    /// <paramref name="suppressed"/>.
    /// </returns>
    private static DateLabelRefusal? Plan(
        DateLabelRequest request,
        DateOnly date,
        string role,
        GanttLabelPosition? explicitPosition,
        bool clipped,
        List<RectD> occupied,
        List<SceneText> emitted,
        List<string> suppressed)
    {
        // Section 23: the start label anchors Left of the visible bar and the
        // finish label Right, unless the row names an explicit position.
        GanttLabelPosition position =
            explicitPosition
            ?? (role == StartRole ? GanttLabelPosition.Left : GanttLabelPosition.Right);

        var text = GanttDateFormatting.Format(date, request.DateFormat);
        GanttEvent @event = request.Event;
        var owner = SceneOwnerId.ForRow(@event.Id);

        LabelPlanCreationOutcome planned = LabelPlanner.TryPlan(
            new LabelRequest(
                @event,
                text,
                position,
                request.VisibleBounds,
                request.TextStyle,
                request.TextMetrics,
                Role: role),
            request.Metrics,
            occupied);

        // A truncated date is not a date. D-G11 and §23 require the *true* date,
        // so an ellipsised "05/01/20…" is treated exactly as a declined
        // placement: it falls through to the clipped/unclipped and suppressed
        // paths below rather than being emitted as a plausible-looking wrong date.
        if (planned.Succeeded && planned.Result!.Primitive is { } primitive && !planned.Result.WasTruncated)
        {
            emitted.Add(primitive);
            if (planned.Result.Bounds is { } placed)
            {
                // Register the placed box so the finish label is planned against
                // the start label rather than overlapping it.
                occupied.Add(placed);
            }

            return null;
        }

        // A planner *refusal* is a broken dependency, not a placement decision,
        // so it surfaces as a typed refusal rather than a silently missing date.
        if (planned.Refusal is not null)
        {
            return DateLabelRefusal.PlannerRefused;
        }

        if (clipped)
        {
            SceneText fallback = Unclipped(request, text, role, position, owner, @event);
            emitted.Add(fallback);

            // The fallback box is an emitted label like any other, so it is
            // registered as an occupant. Without this the finish label would be
            // planned as if the start date had never been drawn, and the two
            // could land on the same box -- the exact overlap the occupant list
            // exists to prevent.
            occupied.Add(fallback.TextBounds);
            return null;
        }

        suppressed.Add(role);
        return null;
    }

    /// <summary>
    /// Emits the true date for a clipped event at a deterministic box beside the
    /// visible edge, because section 23 forbids suppressing an off-plot date.
    /// </summary>
    private static SceneText Unclipped(
        DateLabelRequest request,
        string text,
        string role,
        GanttLabelPosition position,
        SceneOwnerId owner,
        GanttEvent @event)
    {
        if (!request.TextMetrics.TryMeasure(text, out TextMeasurement? measured) || measured is null)
        {
            measured = new TextMeasurement(text.Length * 4.0, request.Metrics.LabelHeightPt);
        }

        var height = request.Metrics.LabelHeightPt;
        var top = request.VisibleBounds.Y + ((request.VisibleBounds.Height - height) / 2);
        var left =
            position is GanttLabelPosition.Right
                ? request.VisibleBounds.Right + request.Metrics.LabelGapPt
                : request.VisibleBounds.Left - request.Metrics.LabelGapPt - measured.WidthPt;

        return new SceneText(
            $"{@event.Id.Value}:{role}",
            owner,
            ZLayer.Label,
            text,
            new RectD(left, top, measured.WidthPt, height),
            request.TextStyle,
            position is GanttLabelPosition.Right ? GanttTextAlignment.Left : GanttTextAlignment.Right,
            @event.Type);
    }

    private static bool IsPositive(RectD rect) =>
        double.IsFinite(rect.X)
        && double.IsFinite(rect.Y)
        && rect.Width > 0
        && rect.Height > 0;

    private static DateLabelOutcome Refused(DateLabelRefusal refusal) => new(null, refusal);
}
