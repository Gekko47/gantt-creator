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
/// <param name="Warnings">
/// The deterministic non-blocking warnings. A never-suppress fallback that had to
/// move its box to keep the label on the chart reports
/// <see cref="DateLabelBuilder.PositionChangedCode"/>, because a position the
/// caller never asked for must be visible in the scene rather than inferred from
/// the emitted geometry.
/// </param>
public sealed record DateLabelResult(
    IReadOnlyList<SceneText> Primitives,
    IReadOnlyList<string> Suppressed,
    IReadOnlyList<SceneWarning> Warnings);

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

    /// <summary>
    /// The warning code emitted when the never-suppress fallback had to place an
    /// off-plot date label somewhere other than the requested position.
    /// </summary>
    public const string PositionChangedCode = "DateLabelPositionChanged";

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

        // Geometry-derived, not caller-asserted: the plot shortened the entity's
        // horizontal extent. Compared on width alone, with the epsilon, because
        // that is precisely what clipping changes -- a caller that supplies
        // rectangles differing only in X or Y has not clipped anything, and
        // reading that as clipped would divert the label onto the never-suppress
        // fallback for a bar the plot never touched.
        var clipped = request.VisibleBounds.Width
            < request.FullBounds.Width - GeometryMath.Epsilon;
        List<SceneText> emitted = [];
        List<string> suppressed = [];

        // The caller's already-placed labels block both date labels, and the
        // start label this builder places then blocks the finish label, so the
        // two dates of one row are planned against each other rather than each
        // being planned in isolation and overlapping.
        List<RectD> occupied = [.. request.Occupants ?? []];
        List<SceneWarning> warnings = [];

        DateLabelRefusal? refusal = null;
        if (request.ShowStart && request.Event.Start is { } start)
        {
            refusal = Plan(request, start, StartRole, request.StartPosition, clipped, occupied, emitted, suppressed, warnings);
        }

        if (refusal is null && request.ShowFinish && request.Event.Finish is { } finish)
        {
            refusal = Plan(request, finish, FinishRole, request.FinishPosition, clipped, occupied, emitted, suppressed, warnings);
        }

        // A planner refusal is a broken dependency, not a placement decision, so
        // it is never reported as a merely missing date.
        return refusal is { } reason ? Refused(reason) : new DateLabelOutcome(new DateLabelResult(emitted, suppressed, warnings), null);
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
        List<string> suppressed,
        List<SceneWarning> warnings)
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
            SceneText fallback = Unclipped(
                request,
                text,
                role,
                position,
                owner,
                @event,
                out GanttLabelPosition used);
            emitted.Add(fallback);

            // A position the caller did not ask for is recorded, not hidden: the
            // scene is the only place a reader can see that the never-suppress
            // fallback moved the date off the requested side to keep it on the
            // chart. Silently flipping would make the emitted geometry disagree
            // with the resolved position with nothing to explain the difference.
            if (used != position)
            {
                warnings.Add(new SceneWarning(
                    owner,
                    PositionChangedCode,
                    "The off-plot date label was placed on a different side so it would stay inside the chart."));
            }

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
    /// Emits the true date for a clipped event at a deterministic box, because
    /// section 23 forbids suppressing an off-plot date.
    /// </summary>
    /// <param name="request">The typed date-label request.</param>
    /// <param name="text">The true, unclamped date text.</param>
    /// <param name="role">The date-label role, which names the emitted primitive.</param>
    /// <param name="position">The requested position.</param>
    /// <param name="owner">The owning scene row.</param>
    /// <param name="event">The validated event the label belongs to.</param>
    /// <param name="used">The position the box was finally placed at.</param>
    /// <returns>The emitted scene text.</returns>
    /// <remarks>
    /// <para>
    /// The never-suppress rule requires a label, not a particular box, so this
    /// fallback owns three things the planner's own contract assumes: the box is
    /// contained within the chart bounds, an explicit position is honoured rather
    /// than silently replaced by the start role's Left, and the position actually
    /// used is reported back so a caller can record the change.
    /// </para>
    /// <para>
    /// A clipped bar sits against a plot edge, which is exactly where an external
    /// box escapes the chart. Each position is tried, then its opposite side, and
    /// only then is the box clamped into the chart. Clamping is the last resort
    /// because §23 requires the true date to be emitted and a box outside the
    /// chart would be invisible; it shifts the box rather than shortening it, so
    /// the text is never elided into a plausible-looking wrong date (D-G11).
    /// </para>
    /// </remarks>
    private static SceneText Unclipped(
        DateLabelRequest request,
        string text,
        string role,
        GanttLabelPosition position,
        SceneOwnerId owner,
        GanttEvent @event,
        out GanttLabelPosition used)
    {
        if (!request.TextMetrics.TryMeasure(text, out TextMeasurement? measured) || measured is null)
        {
            measured = new TextMeasurement(text.Length * 4.0, request.Metrics.LabelHeightPt);
        }

        var height = request.Metrics.LabelHeightPt;
        var gap = request.Metrics.LabelGapPt;
        RectD visible = request.VisibleBounds;
        RectD chart = request.Metrics.ChartBounds;

        // The §22 candidate geometry, resolved here because this path is reached
        // only after the planner has already declined every candidate.
        var top = visible.Top + ((visible.Height - height) / 2);
        var centred = visible.Left + ((visible.Width - measured.WidthPt) / 2);

        // Every enum member is listed rather than folded into a default arm, so a
        // position added to GanttLabelPosition later cannot silently acquire this
        // builder's geometry. The corners, the splitter placements, and None/Auto
        // are not date-label positions; they take the Left box, which is what the
        // two-case form this replaces did for every non-Right value.
        RectD Box(GanttLabelPosition candidate)
        {
            return candidate switch
            {
                GanttLabelPosition.Right => new RectD(visible.Right + gap, top, measured.WidthPt, height),
                GanttLabelPosition.Inside => new RectD(centred, top, measured.WidthPt, height),

                GanttLabelPosition.None
                or GanttLabelPosition.Auto
                or GanttLabelPosition.Left
                or GanttLabelPosition.TopLeft
                or GanttLabelPosition.TopRight
                or GanttLabelPosition.BottomLeft
                or GanttLabelPosition.BottomRight
                or GanttLabelPosition.DataPanelLeft
                or GanttLabelPosition.PlotCentre
                or GanttLabelPosition.Both
                    => new RectD(visible.Left - gap - measured.WidthPt, top, measured.WidthPt, height),
                _ => throw new ArgumentOutOfRangeException(nameof(candidate)),
            };
        }

        // Every remaining position tries the two external sides. Inside is
        // deliberately absent: a box inside a clipped sliver is unreadable, and
        // §23 wants the date legible. Above and Below used to flip to each other
        // here; both are retired (owner ruling 2026-09-30).
        GanttLabelPosition[] order = position switch
        {
            GanttLabelPosition.None
            or GanttLabelPosition.Auto
            or GanttLabelPosition.Left
            or GanttLabelPosition.Right
            or GanttLabelPosition.Inside
            or GanttLabelPosition.TopLeft
            or GanttLabelPosition.TopRight
            or GanttLabelPosition.BottomLeft
            or GanttLabelPosition.BottomRight
            or GanttLabelPosition.DataPanelLeft
            or GanttLabelPosition.PlotCentre
            or GanttLabelPosition.Both => [position, GanttLabelPosition.Right, GanttLabelPosition.Left],
            _ => throw new ArgumentOutOfRangeException(nameof(position)),
        };

        used = order[0];
        foreach (GanttLabelPosition candidate in order)
        {
            if (Contains(chart, Box(candidate)))
            {
                used = candidate;
                return Emit(Box(candidate), candidate);
            }
        }

        // Nothing fits: keep the requested position and clamp, so the true date is
        // still emitted inside a chart too small to hold it beside the bar.
        return Emit(Clamp(chart, Box(position)), position);

        SceneText Emit(RectD box, GanttLabelPosition placed)
        {
            return new SceneText(
                $"{@event.Id.Value}:{role}",
                owner,
                ZLayer.Label,
                text,
                box,
                request.TextStyle,

                // §22's alignment rule: a right-hand box reads away from the shape,
                // a left-hand box ends against it, and the centred positions centre.
                placed is GanttLabelPosition.Left ? GanttTextAlignment.Right
                    : placed is GanttLabelPosition.Right ? GanttTextAlignment.Left
                    : GanttTextAlignment.Centre,
                @event.Type);
        }
    }

    /// <summary>Returns whether the box lies entirely within the chart bounds.</summary>
    private static bool Contains(RectD outer, RectD inner) =>
        inner.Left >= outer.Left - GeometryMath.Epsilon
        && inner.Top >= outer.Top - GeometryMath.Epsilon
        && inner.Right <= outer.Right + GeometryMath.Epsilon
        && inner.Bottom <= outer.Bottom + GeometryMath.Epsilon;

    /// <summary>
    /// Shifts a box the minimum distance needed to bring it inside the chart,
    /// keeping its size.
    /// </summary>
    private static RectD Clamp(RectD chart, RectD box)
    {
        var left = Math.Clamp(box.Left, chart.Left, Math.Max(chart.Left, chart.Right - box.Width));
        var top = Math.Clamp(box.Top, chart.Top, Math.Max(chart.Top, chart.Bottom - box.Height));
        return new RectD(left, top, box.Width, box.Height);
    }

    private static bool IsPositive(RectD rect) =>
        double.IsFinite(rect.X)
        && double.IsFinite(rect.Y)
        && rect.Width > 0
        && rect.Height > 0;

    private static DateLabelOutcome Refused(DateLabelRefusal refusal) => new(null, refusal);
}
