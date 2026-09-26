using System.Globalization;
using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

public sealed class DateLabelBuilderTests
{
    private static readonly ITextMetrics _metrics = new FakeTextMetrics(_ => 10.0, 10.0);
    private static readonly SceneStyle _style = new("DefaultText", fontFamily: "Aptos", fontSizePt: 8.0);
    private static readonly LabelMetrics _labelMetrics = new(
        new RectD(0.0, 0.0, 400.0, 200.0),
        new RectD(-10.0, -10.0, 420.0, 220.0),
        4.0,
        10.0,
        120.0);

    [Fact]
    public void A_span_emits_both_labels_with_the_stable_role_derived_ids()
    {
        // The expected ids come from the request that was actually built, not
        // from a second Row(): GanttRowId.New() mints a distinct id per call, so
        // a separately built event would never match.
        DateLabelRequest request = Request();
        DateLabelResult result = DateLabelBuilder.TryBuild(request).Result!;
        var row = request.Event.Id.Value;

        Assert.Equal(
            [$"{row}:date-start", $"{row}:date-finish"],
            [.. result.Primitives.Select(label => label.PrimitiveId)]);
    }

    [Fact]
    public void The_start_label_anchors_left_and_the_finish_label_right()
    {
        DateLabelResult result = Build().Result!;

        // Section 23: start anchors Left of the visible bar, finish Right.
        SceneText start = result.Primitives[0];
        SceneText finish = result.Primitives[1];
        Assert.True(start.TextBounds.Right <= Visible().Left, "the start label should sit left of the bar");
        Assert.True(finish.TextBounds.Left >= Visible().Right, "the finish label should sit right of the bar");
    }

    [Fact]
    public void The_dates_render_in_the_approved_invariant_format()
    {
        CultureInfo original = CultureInfo.CurrentCulture;
        try
        {
            string[][] observed = [.. _cultures.Select(culture =>
            {
                CultureInfo.CurrentCulture = new CultureInfo(culture);
                return Build().Result!.Primitives.Select(label => label.Text).ToArray();
            })];

            Assert.All(observed, texts => Assert.Equal(_expectedDates, texts));
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }

    private static readonly string[] _cultures = ["en-GB", "en-US"];

    private static readonly string[] _expectedDates = ["05/01/2024", "09/01/2024"];

    [Fact]
    public void Visibility_is_independent_so_finish_can_be_hidden_while_start_shows()
    {
        // The positive independent-visibility pin: one flag off, the other on.
        DateLabelResult result = Build(showFinish: false).Result!;

        SceneText only = Assert.Single(result.Primitives);
        Assert.EndsWith(":date-start", only.PrimitiveId, StringComparison.Ordinal);
        Assert.Empty(result.Suppressed);
    }

    [Fact]
    public void A_point_event_never_gets_a_finish_date_label()
    {
        // Section 23: no start/finish label for a field the event type does not
        // use. A milestone's Finish is null, so the rule needs no lookup.
        DateLabelResult result = Build(hasFinish: false).Result!;

        SceneText only = Assert.Single(result.Primitives);
        Assert.EndsWith(":date-start", only.PrimitiveId, StringComparison.Ordinal);
    }

    [Fact]
    public void A_clipped_event_still_shows_its_true_off_plot_date()
    {
        // D-G11: a plot-clipped event always displays the true date and the
        // off-plot date is never suppressed. Visible must be *narrower* than Full
        // -- the builder derives "clipped" by comparing the two, so passing the
        // same rectangle twice would leave clipped false and this test would
        // never reach the clipped path it names. The label text is the real date,
        // not a clamped one.
        DateLabelResult result = Build(full: new RectD(150.0, 100.0, 180.0, 12.0)).Result!;

        Assert.Empty(result.Suppressed);
        Assert.Equal(2, result.Primitives.Count);
        Assert.Equal("05/01/2024", result.Primitives[0].Text);
    }

    [Fact]
    public void A_clipped_event_whose_label_cannot_be_placed_is_still_emitted()
    {
        // The never-suppress fallback: a bar hard against the chart's left edge
        // leaves no room outside it, so the planner declines, but the date must
        // still be emitted rather than dropped.
        RectD pinned = new(LabelEdge().Left, 100.0, 40.0, 12.0);
        DateLabelResult result = DateLabelBuilder
            .TryBuild(Request(visible: pinned, full: new RectD(-500.0, 100.0, 600.0, 12.0)))
            .Result!;

        Assert.Empty(result.Suppressed);
        Assert.Equal(2, result.Primitives.Count);
    }

    [Fact]
    public void A_same_date_start_and_finish_resolve_deterministically()
    {
        // One request, built once: Row() mints a new id per call, so rebuilding
        // the request per run would compare three different entities and prove
        // nothing about determinism.
        DateLabelRequest request = Request(start: new DateOnly(2024, 1, 5), finish: new DateOnly(2024, 1, 5));
        string[][] runs =
        [
            .. Enumerable.Range(0, 3).Select(_ => DateLabelBuilder.TryBuild(request).Result!.Primitives
                .Select(label => $"{label.PrimitiveId}@{label.TextBounds}")
                .ToArray()),
        ];

        Assert.Equal(runs[0], runs[1]);
        Assert.Equal(runs[0], runs[2]);
        Assert.Equal(2, runs[0].Length);
        Assert.Contains(
            $"{request.Event.Id.Value}:date-start",
            runs[0][0],
            StringComparison.Ordinal);
    }

    [Fact]
    public void The_placed_start_label_blocks_the_finish_label()
    {
        // The finish label is planned after the start label, so the start label's
        // placed box must be an occupant. The default Left/Right anchors point
        // away from each other and could never collide, which would make the
        // assertion vacuous, so both roles are given the same explicit position:
        // without the registration the two land on the identical box.
        // The start label takes the box; the finish label is then planned against
        // it, finds that position occupied, and is declined rather than drawn on
        // top. One primitive and one suppressed role is the correct outcome.
        DateLabelResult result = DateLabelBuilder
            .TryBuild(Request() with
            {
                StartPosition = GanttLabelPosition.Right,
                FinishPosition = GanttLabelPosition.Right,
            })
            .Result!;

        SceneText start = Assert.Single(result.Primitives);
        Assert.EndsWith(":date-start", start.PrimitiveId, StringComparison.Ordinal);
        Assert.Equal([DateLabelBuilder.FinishRole], result.Suppressed);
    }

    [Fact]
    public void A_caller_supplied_occupant_blocks_a_date_label()
    {
        // The Occupants member carries the bounds every other LabelPlanner caller
        // passes -- here the row's own description label. An occupant sitting
        // exactly where the start label would go must stop it being placed there.
        DateLabelRequest bare = new(
            Row(_start, _finish),
            Visible(),
            Full(),
            _labelMetrics,
            new FakeTextMetrics(_ => 10.0, 10.0),
            GanttDateDisplayFormat.DdMMyyyy,
            _style,
            ShowFinish: false);

        // Where the start label lands with nothing to avoid, established from the
        // same request so the occupant is the real planned box.
        SceneText unblocked = Assert.Single(DateLabelBuilder.TryBuild(bare).Result!.Primitives);

        // With that box occupied, the only §22 position for a start date is
        // blocked, so the planner declines and the role is reported suppressed
        // rather than a second label being drawn on top of the occupant.
        DateLabelResult blocked = DateLabelBuilder.TryBuild(bare with { Occupants = [unblocked.TextBounds] }).Result!;

        Assert.Empty(blocked.Primitives);
        Assert.Equal([DateLabelBuilder.StartRole], blocked.Suppressed);
    }

    [Fact]
    public void A_truncated_date_is_never_emitted_as_a_plausible_wrong_date()
    {
        // D-G11 and §23 require the *true* date. The planner's widest-gap
        // fallback truncates with an ellipsis, which would render "05/01/20…"
        // -- a plausible-looking wrong date that reads as a real date. A
        // truncated result must therefore take the declined path, never the
        // emitted one.
        //
        // The advance table makes every date character wide (the 10-character
        // date needs 400pt, far past the 120pt external maximum) while the
        // ellipsis stays 2pt, so the gap holds the ellipsis and the planner
        // truncates rather than declining outright. That is the only path that
        // produces a truncated result, so it is the path under test.
        var metrics = new LabelMetrics(
            new RectD(0.0, 0.0, 400.0, 200.0),
            new RectD(-10.0, -10.0, 420.0, 220.0),
            4.0,
            10.0,
            120.0);
        var wide = new FakeTextMetrics(c => c == '…' ? 2.0 : 40.0, 10.0);
        var request = Request() with { TextMetrics = wide, Metrics = metrics };

        // The planner really does offer a truncated label for these inputs, or
        // this test would be asserting nothing.
        LabelPlanCreationOutcome plan = LabelPlanner.TryPlan(
            new LabelRequest(
                request.Event,
                "05/01/2024",
                GanttLabelPosition.Left,
                request.VisibleBounds,
                request.TextStyle,
                wide,
                Role: DateLabelBuilder.StartRole),
            metrics);
        Assert.True(plan.Result!.WasTruncated, "the fixture must reach the truncation path");

        DateLabelResult result = DateLabelBuilder.TryBuild(request).Result!;

        Assert.DoesNotContain(
            result.Primitives,
            label => label.Text.Contains('…', StringComparison.Ordinal));
        Assert.Equal(
            [DateLabelBuilder.StartRole, DateLabelBuilder.FinishRole],
            result.Suppressed);
    }

    [Fact]
    public void A_null_request_is_refused() =>
        Assert.Equal(DateLabelRefusal.NullRequest, DateLabelBuilder.TryBuild(null).Refusal);

    [Fact]
    public void A_null_event_is_refused() =>
        Assert.Equal(DateLabelRefusal.NullEvent, DateLabelBuilder.TryBuild(Request() with { Event = null! }).Refusal);

    [Fact]
    public void Null_metrics_are_refused()
    {
        Assert.Equal(DateLabelRefusal.NullMetrics, DateLabelBuilder.TryBuild(Request() with { Metrics = null! }).Refusal);
        Assert.Equal(DateLabelRefusal.NullTextMetrics, DateLabelBuilder.TryBuild(Request() with { TextMetrics = null! }).Refusal);
    }

    [Fact]
    public void A_null_text_style_is_refused() =>
        Assert.Equal(DateLabelRefusal.NullTextStyle, DateLabelBuilder.TryBuild(Request() with { TextStyle = null! }).Refusal);

    [Fact]
    public void An_undefined_entity_type_is_refused() =>
        Assert.Equal(
            DateLabelRefusal.InvalidEntityType,
            DateLabelBuilder.TryBuild(Request() with { Event = Row(_start, _finish) with { Type = (GanttEntityType)99 } }).Refusal);

    [Fact]
    public void Zero_extent_bounds_are_refused() =>
        Assert.Equal(
            DateLabelRefusal.InvalidBounds,
            DateLabelBuilder.TryBuild(Request() with { VisibleBounds = new RectD(100.0, 100.0, 0.0, 12.0) }).Refusal);

    [Fact]
    public void A_planner_refusal_is_surfaced_rather_than_reported_as_a_missing_date()
    {
        // A broken dependency must not look like a placement decision, or a
        // missing date would read as an ordinary suppression.
        Assert.Equal(
            DateLabelRefusal.PlannerRefused,
            DateLabelBuilder.TryBuild(Request() with { TextMetrics = new FakeTextMetrics(_ => double.NaN, 10.0) }).Refusal);
    }

    [Fact]
    public void A_blank_role_is_refused_by_the_planner()
    {
        // The new LabelRequest.Role guard: a blank role would emit a primitive
        // whose ID ends in a bare separator.
        LabelPlanCreationOutcome outcome = LabelPlanner.TryPlan(
            new LabelRequest(Row(_start, _finish), "x", GanttLabelPosition.Left, Visible(), _style, _metrics, Role: "  "),
            _labelMetrics);

        Assert.Equal(LabelRefusal.BlankRole, outcome.Refusal);
    }

    private static DateLabelOutcome Build(
        DateOnly? start = null,
        DateOnly? finish = null,
        bool hasFinish = true,
        bool showFinish = true,
        RectD? visible = null,
        RectD? full = null) =>
        DateLabelBuilder.TryBuild(Request(start, finish, hasFinish, showFinish, visible, full));

    private static readonly DateOnly _start = new(2024, 1, 5);

    private static readonly DateOnly _finish = new(2024, 1, 9);

    private static DateLabelRequest Request(
        DateOnly? start = null,
        DateOnly? finish = null,
        bool hasFinish = true,
        bool showFinish = true,
        RectD? visible = null,
        RectD? full = null) =>
        new(
            Row(start ?? _start, hasFinish ? finish ?? _finish : null),
            visible ?? Visible(),
            full ?? Full(),
            _labelMetrics,
            _metrics,
            GanttDateDisplayFormat.DdMMyyyy,
            _style,
            ShowFinish: showFinish);

    private static GanttEvent Row(DateOnly start, DateOnly? finish) =>
        new(
            1,
            GanttRowId.New(),
            null,
            null,
            GanttEntityType.AsPlannedActivity,
            "Activity",
            start,
            finish,
            null,
            null,
            null,
            null,
            null,
            true,
            null);

    private static RectD Visible() => new(200.0, 100.0, 40.0, 12.0);

    private static RectD Full() => Visible();

    private static RectD LabelEdge() => _labelMetrics.ChartBounds;
}
