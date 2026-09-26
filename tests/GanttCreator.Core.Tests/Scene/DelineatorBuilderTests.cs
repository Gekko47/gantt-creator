using GanttCreator.Core;
using GanttCreator.Core.Scene;
using Xunit;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>Contract tests for the R3.10 delineator builder (entity guide §24).</summary>
/// <remarks>
/// Every guard has a positive test that constructs the bad input, so a
/// validator cannot ship silently inert (AGENTS.md).
/// </remarks>
public sealed class DelineatorBuilderTests
{
    private const double Left = 100.0;
    private const double Right = 300.0;
    private const double Day = 20.0;

    private static readonly RectD Plot = new(Left, 40.0, Right - Left, 60.0);
    private static readonly RectD Chart = new(0.0, 0.0, 400.0, 200.0);

    private static int s_next;

    // One shared measuring seam, so a group of requests compares as the same
    // seam. A fresh instance per request would make every group inconsistent.
    private static readonly ITextMetrics Metrics = new FakeTextMetrics();

    private static GanttRowId Id() => GanttRowId.Parse($"G-{s_next++:x32}");

    private static TimeScale Scale() =>
        TimeScale.TryCreate(new DateOnly(2026, 9, 10), new DateOnly(2026, 9, 19), Left, Right).Scale!;

    private static GanttEvent Event(
        DateOnly? start = null,
        string? description = "Gate",
        GanttEntityType type = GanttEntityType.Delineator,
        DateOnly? finish = null,
        int? sortOrder = null,
        bool noStart = false) =>
        new(
            1, Id(), null, null, type, description, noStart ? null : start ?? new DateOnly(2026, 9, 12), finish,
            null, null, null, null, null, true, sortOrder);

    private static DelineatorRequest Req(
        GanttEvent e,
        ITextMetrics? metrics = null,
        GanttLabelPosition position = GanttLabelPosition.Auto,
        double? width = null) =>
        new(
            e,
            new SceneStyle("DefaultDelineator", strokeColour: ColourHex.Parse("#404040")),
            width ?? 0.75,
            Plot,
            Chart,
            2.0,
            metrics ?? Metrics,
            position);

    private static DelineatorResult Build(
        DelineatorRequest request,
        IReadOnlyList<RectD>? occupants = null,
        double stack = 0)
    {
        DelineatorCreationOutcome outcome = DelineatorBuilder.TryBuild(request, Scale(), occupants, stack);
        Assert.Null(outcome.Refusal);
        return outcome.Result!;
    }

    private static DelineatorResult Build(
        GanttEvent e,
        ITextMetrics? metrics = null,
        GanttLabelPosition position = GanttLabelPosition.Auto) => Build(Req(e, metrics, position));

    [Fact]
    public void The_line_x_is_the_exact_date_position()
    {
        DelineatorResult r = Build(Event(new DateOnly(2026, 9, 12)));

        Assert.Equal(Left + (2 * Day), r.X);
        Assert.Equal(Left + (2 * Day), r.Primitive!.From.X);
    }

    [Fact]
    public void The_line_spans_the_plot_height_and_never_the_title_or_header_bands()
    {
        // Positive test for §24's exclusion rule: the chart bounds are present
        // and larger, so a line spanning the chart would start at 0, not 40.
        DelineatorResult r = Build(Event());

        Assert.Equal(Plot.Top, r.Primitive!.From.Y);
        Assert.Equal(Plot.Bottom, r.Primitive.To.Y);
        Assert.True(r.Primitive.From.Y > Chart.Top);
        Assert.True(r.Primitive.To.Y < Chart.Bottom);
    }

    [Fact]
    public void The_line_uses_the_delineator_line_width_and_stroke()
    {
        DelineatorResult r = Build(Req(Event(), width: 1.25));

        Assert.Equal(1.25, r.Primitive!.Style.OutlineWidthPt);
        Assert.Equal(ColourHex.Parse("#404040"), r.Primitive.Style.StrokeColour);
    }

    [Theory]
    [InlineData(2026, 9, 10, Left)]  // first day, on the left boundary
    [InlineData(2026, 9, 19, 280.0)] // last day: day index 9, so Left + 9 * Day
    public void The_inclusive_plot_edges_place_a_line_exactly_on_the_boundary(int y, int m, int d, double expected)
    {
        // The plot is ten inclusive days wide, so the last date's point X is
        // Left + 9 * Day. The plot's right *boundary* (300) is the inclusive
        // activity right edge, which a point event never reaches: §24 shares the
        // point-event rule that no extra day is added.
        DelineatorResult r = Build(Event(new DateOnly(y, m, d)));

        Assert.Equal(expected, r.X);
        Assert.Empty(r.Warnings);
    }

    [Fact]
    public void An_out_of_range_date_emits_no_line_and_exactly_one_warning()
    {
        // Positive test for §24 D2: warn, never throw, and never use the Try*
        // failure sentinel as a coordinate.
        DelineatorResult r = Build(Event(new DateOnly(2026, 9, 25)));

        Assert.Null(r.Primitive);
        Assert.Null(r.X);
        Assert.Equal(DelineatorBuilder.OutsidePlotRangeCode, Assert.Single(r.Warnings).Code);
    }

    [Fact]
    public void A_delineator_reads_only_start_and_never_finish()
    {
        DelineatorResult withFinish = Build(Event(finish: new DateOnly(2026, 12, 31)));
        DelineatorResult without = Build(Event());

        Assert.Equal(without.X, withFinish.X);
    }

    [Fact]
    public void The_line_and_label_carry_their_role_derived_ids()
    {
        GanttEvent e = Event();
        DelineatorResult r = Build(e);

        Assert.Equal($"{e.Id.Value}:delineator", r.Primitive!.PrimitiveId);
        Assert.Equal($"{e.Id.Value}:delineator-label", r.Label!.PrimitiveId);
    }

    [Fact]
    public void The_line_sits_below_activity_bodies_and_the_label_above_both()
    {
        DelineatorResult r = Build(Event());

        Assert.True(ZLayer.Delineator < ZLayer.ActivityBody);
        Assert.Equal(ZLayer.Delineator, r.Primitive!.ZLayer);
        Assert.Equal(ZLayer.DelineatorLabel, r.Label!.ZLayer);
        Assert.True(ZLayer.ActivityBody < ZLayer.DelineatorLabel);
    }

    private static readonly GanttLabelPosition[] Corners =
    [
        GanttLabelPosition.TopRight,
        GanttLabelPosition.TopLeft,
        GanttLabelPosition.BottomRight,
        GanttLabelPosition.BottomLeft,
    ];

    [Theory]
    [InlineData(GanttLabelPosition.TopRight)]
    [InlineData(GanttLabelPosition.TopLeft)]
    [InlineData(GanttLabelPosition.BottomRight)]
    [InlineData(GanttLabelPosition.BottomLeft)]
    public void Every_permitted_corner_is_honoured_when_explicitly_requested(GanttLabelPosition position)
    {
        DelineatorResult r = Build(Event(), position: position);

        Assert.Equal(position, r.LabelPosition);
        Assert.NotNull(r.LabelBounds);
    }

    [Fact]
    public void A_none_position_emits_a_line_but_no_text()
    {
        DelineatorResult r = Build(Event(), position: GanttLabelPosition.None);

        Assert.NotNull(r.Primitive);
        Assert.Null(r.Label);
        Assert.Empty(r.Warnings);
    }

    [Fact]
    public void The_auto_cascade_starts_at_top_right()
    {
        Assert.Equal(GanttLabelPosition.TopRight, Build(Event()).LabelPosition);
    }

    [Fact]
    public void An_occupied_top_right_corner_flips_to_the_next_candidate()
    {
        RectD blocker = Build(Event()).LabelBounds!.Value;
        DelineatorResult r = Build(Req(Event()), [blocker]);

        Assert.Equal(GanttLabelPosition.TopLeft, r.LabelPosition);
        Assert.False(r.LabelBounds!.Value.IntersectsWith(blocker));
    }

    [Fact]
    public void A_corner_label_is_offset_from_the_line_by_the_label_gap()
    {
        // "Gate" is four characters at the fake metrics' 4pt advance, so 16pt.
        RectD box = Build(Event()).LabelBounds!.Value;

        Assert.Equal(Left + (2 * Day) + 2.0, box.Left);
        Assert.Equal(Plot.Top, box.Top);
        Assert.Equal(16.0, box.Width);
    }

    [Fact]
    public void Every_corner_label_stays_within_the_chart_bounds()
    {
        foreach (GanttLabelPosition position in Corners)
        {
            RectD box = Build(Event(), position: position).LabelBounds!.Value;

            Assert.True(box.Left >= Chart.Left, $"{position} escaped left.");
            Assert.True(box.Right <= Chart.Right, $"{position} escaped right.");
            Assert.True(box.Top >= Chart.Top, $"{position} escaped top.");
            Assert.True(box.Bottom <= Chart.Bottom, $"{position} escaped bottom.");
        }
    }

    [Fact]
    public void A_label_no_corner_can_hold_is_suppressed_with_exactly_one_warning()
    {
        // Positive test for the suppression path: four blockers leave no corner,
        // and the builder warns once rather than drawing over an obstruction.
        List<RectD> blockers = [.. Corners.Select(c => Build(Event(), position: c).LabelBounds!.Value)];
        DelineatorResult r = Build(Req(Event()), blockers);

        Assert.NotNull(r.Primitive);
        Assert.Null(r.Label);
        Assert.Equal(DelineatorBuilder.LabelSuppressedCode, Assert.Single(r.Warnings).Code);
    }

    [Fact]
    public void A_blank_description_emits_no_label_and_no_warning()
    {
        DelineatorResult r = Build(Event(description: null));

        Assert.NotNull(r.Primitive);
        Assert.Null(r.Label);
        Assert.Empty(r.Warnings);
    }

    [Fact]
    public void The_stack_offset_moves_a_label_away_from_the_plot_edge()
    {
        RectD first = Build(Req(Event())).LabelBounds!.Value;
        RectD stacked = Build(Req(Event()), null, 6.0).LabelBounds!.Value;

        Assert.Equal(first.Top + 6.0, stacked.Top);
    }

    // ------------------------------------------------------------------
    // Same-date deduplication and stacking (§24 D4, ADR-0017)
    // ------------------------------------------------------------------

    private static DelineatorGroupResult Group(
        params GanttEvent[] events) => DelineatorLayout
        .TryBuildGroup(new DelineatorGroupRequest([.. events.Select(e => Req(e))], 2.0), Scale()).Result!;

    [Fact]
    public void A_same_date_group_draws_one_line_owned_by_every_contributing_row()
    {
        // Positive test for §24 D4: two rows on one date collapse to a single
        // line and the shared owner records both (ADR-0017).
        DelineatorGroupResult r = Group(Event(description: "Alpha"), Event(description: "Bravo"));

        SceneLine line = Assert.IsType<SceneLine>(r.Primitives[0]);
        Assert.Single(r.Primitives.OfType<SceneLine>());
        Assert.Equal(SceneOwnerKind.Rows, line.OwnerId.Kind);
        Assert.Equal(2, line.OwnerId.OwnedRows.Count);
    }

    [Fact]
    public void A_single_row_group_keeps_a_plain_row_owner()
    {
        SceneLine line = Assert.IsType<SceneLine>(Group(Event()).Primitives[0]);

        Assert.Equal(SceneOwnerKind.Row, line.OwnerId.Kind);
    }

    [Fact]
    public void Same_date_distinct_labels_stack_by_the_stack_gap()
    {
        DelineatorGroupResult r = DelineatorLayout.TryBuildGroup(
            new DelineatorGroupRequest(
                [Req(Event(description: "Alpha", sortOrder: 1)), Req(Event(description: "Bravo", sortOrder: 2))],
                6.0),
            Scale()).Result!;

        List<SceneText> labels = [.. r.Primitives.OfType<SceneText>()];

        Assert.Equal(2, labels.Count);
        Assert.Equal(6.0, Math.Abs(labels[1].TextBounds.Top - labels[0].TextBounds.Top));
        Assert.False(labels[0].TextBounds.IntersectsWith(labels[1].TextBounds));
    }

    [Fact]
    public void A_group_is_independent_of_the_caller_sequence()
    {
        GanttEvent a = Event(description: "Alpha", sortOrder: 1);
        GanttEvent b = Event(description: "Bravo", sortOrder: 2);

        string Snap(IReadOnlyList<GanttEvent> events)
        {
            DelineatorGroupResult r = DelineatorLayout
                .TryBuildGroup(new DelineatorGroupRequest([.. events.Select(e => Req(e))], 2.0), Scale()).Result!;
            return SceneSnapshot.Serialize(GanttScene.TryCreate(Chart, Plot, r.Primitives, []).Scene!);
        }

        Assert.Equal(Snap([a, b]), Snap([b, a]));
    }

    [Fact]
    public void Building_the_same_request_repeatedly_is_byte_identical()
    {
        // The R3.12 golden snapshot can only be stable if this holds.
        GanttEvent e = Event();
        string first = Serialize(Req(e));

        for (var i = 0; i < 3; i++)
        {
            Assert.Equal(first, Serialize(Req(e)));
        }
    }

    private static string Serialize(DelineatorRequest request)
    {
        DelineatorResult r = Build(request);
        List<ScenePrimitive> primitives = [];
        if (r.Primitive is not null)
        {
            primitives.Add(r.Primitive);
        }

        if (r.Label is not null)
        {
            primitives.Add(r.Label);
        }

        return SceneSnapshot.Serialize(GanttScene.TryCreate(Chart, Plot, primitives, []).Scene!);
    }

    // ------------------------------------------------------------------
    // Guard refusals, each with its positive test
    // ------------------------------------------------------------------

    [Fact]
    public void A_null_request_or_time_scale_is_refused()
    {
        Assert.Equal(DelineatorRefusal.NullRequest, DelineatorBuilder.TryBuild(null, Scale()).Refusal);
        Assert.Equal(
            DelineatorRefusal.InvalidTimeScale, DelineatorBuilder.TryBuild(Req(Event()), null).Refusal);
    }

    [Fact]
    public void Non_finite_or_empty_geometry_is_refused()
    {
        DelineatorRequest good = Req(Event());

        Assert.Equal(
            DelineatorRefusal.InvalidGeometry,
            DelineatorBuilder.TryBuild(good with { LineWidthPt = double.NaN }, Scale()).Refusal);
        Assert.Equal(
            DelineatorRefusal.InvalidGeometry,
            DelineatorBuilder.TryBuild(good with { LineWidthPt = 0 }, Scale()).Refusal);
        Assert.Equal(
            DelineatorRefusal.InvalidGeometry,
            DelineatorBuilder.TryBuild(good with { LabelGapPt = double.PositiveInfinity }, Scale()).Refusal);
        Assert.Equal(
            DelineatorRefusal.InvalidGeometry,
            DelineatorBuilder.TryBuild(good with { PlotBounds = new RectD(0, 0, 0, 0) }, Scale()).Refusal);
        Assert.Equal(
            DelineatorRefusal.InvalidGeometry,
            DelineatorBuilder.TryBuild(good with { ChartBounds = new RectD(0, 0, 0, 200) }, Scale())
                .Refusal);
        Assert.Equal(
            DelineatorRefusal.InvalidGeometry, DelineatorBuilder.TryBuild(good, Scale(), null, double.NaN).Refusal);
    }

    [Fact]
    public void A_missing_start_date_is_refused()
    {
        Assert.Equal(
            DelineatorRefusal.MissingEventDate,
            DelineatorBuilder.TryBuild(Req(Event(noStart: true)), Scale()).Refusal);
    }

    [Fact]
    public void A_non_delineator_type_is_refused()
    {
        // Defence in depth: an activity must never become a full-height line.
        Assert.Equal(
            DelineatorRefusal.NotADelineator,
            DelineatorBuilder.TryBuild(Req(Event(type: GanttEntityType.AsBuiltActivity)), Scale()).Refusal);
    }

    [Fact]
    public void A_label_position_the_catalogue_does_not_permit_is_refused()
    {
        // The catalogue is the authority: Inside and the splitter placements are
        // not delineator corners and must not be drawn as one.
        Assert.Equal(
            DelineatorRefusal.UnsupportedLabelPosition,
            DelineatorBuilder.TryBuild(Req(Event(), position: GanttLabelPosition.Inside), Scale()).Refusal);
        Assert.Equal(
            DelineatorRefusal.UnsupportedLabelPosition,
            DelineatorBuilder.TryBuild(Req(Event(), position: GanttLabelPosition.PlotCentre), Scale()).Refusal);
    }

    [Fact]
    public void A_group_refuses_mismatched_members()
    {
        DelineatorRequest first = Req(Event());

        Assert.Equal(
            DelineatorGroupRefusal.NotOneDelineatorDate,
            DelineatorLayout.TryBuildGroup(
                new DelineatorGroupRequest([first, Req(Event(new DateOnly(2026, 9, 13)))], 2.0), Scale()).Refusal);
        Assert.Equal(
            DelineatorGroupRefusal.InconsistentBounds,
            DelineatorLayout.TryBuildGroup(
                new DelineatorGroupRequest(
                    [first, first with { PlotBounds = new RectD(0, 0, 10, 10) }], 2.0), Scale()).Refusal);
        Assert.Equal(
            DelineatorGroupRefusal.InconsistentLabelSeam,
            DelineatorLayout.TryBuildGroup(
                new DelineatorGroupRequest([first, first with { LabelMetrics = new FakeTextMetrics() }], 2.0), Scale())
                .Refusal);
        Assert.Equal(
            DelineatorGroupRefusal.InvalidTimeScale,
            DelineatorLayout.TryBuildGroup(new DelineatorGroupRequest([first], 2.0), null).Refusal);
        Assert.Equal(
            DelineatorGroupRefusal.NullRequest,
            DelineatorLayout.TryBuildGroup(new DelineatorGroupRequest([], 2.0), Scale()).Refusal);
    }
}
