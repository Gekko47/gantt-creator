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
    public void An_unclipped_fallback_label_stays_inside_the_chart_bounds()
    {
        // §24 requires a label box to be "contained within chart bounds" and §22
        // makes containment a condition of acceptance. A bar pinned hard against
        // the chart's left edge has no room for a Left box, so the fallback used to
        // emit one starting left of the chart -- a label no renderer can draw,
        // which is the same as suppressing it while claiming never to suppress.
        // The box is still emitted (the rule holds); it is only kept on the chart.
        RectD chart = LabelEdge();
        RectD pinned = new(chart.Left, 100.0, 40.0, 12.0);
        DateLabelResult result = DateLabelBuilder
            .TryBuild(Request(visible: pinned, full: new RectD(-500.0, 100.0, 600.0, 12.0)))
            .Result!;

        Assert.Empty(result.Suppressed);
        Assert.Equal(2, result.Primitives.Count);
        foreach (SceneText label in result.Primitives)
        {
            Assert.True(
                label.TextBounds.Left >= chart.Left && label.TextBounds.Right <= chart.Right
                && label.TextBounds.Top >= chart.Top && label.TextBounds.Bottom <= chart.Bottom,
                $"The {label.PrimitiveId} fallback box {label.TextBounds} escapes the chart {chart}.");
        }
    }

    [Fact]
    public void An_unclipped_fallback_keeps_the_true_date_rather_than_fitting_it_to_the_chart()
    {
        // The counterpart to the containment rule: bringing the box inside the chart
        // must not shorten the text. §23 and D-G11 require the *true* date, and an
        // ellipsised "05/01/20…" is a plausible-looking wrong date, so the fallback
        // shifts the box and keeps every character.
        RectD chart = LabelEdge();
        RectD pinned = new(chart.Left, 100.0, 40.0, 12.0);
        DateLabelResult result = DateLabelBuilder
            .TryBuild(Request(visible: pinned, full: new RectD(-500.0, 100.0, 600.0, 12.0)))
            .Result!;

        Assert.All(result.Primitives, label => Assert.DoesNotContain('…', label.Text));
        Assert.Contains(result.Primitives, label => label.Text == "05/01/2024");
    }

    [Fact]
    public void An_unclipped_fallback_records_a_position_it_could_not_keep()
    {
        // A position the caller never asked for must be visible in the scene rather
        // than inferable only from the geometry. Without this the fallback could
        // flip a date to the opposite side of a clipped bar and nothing in the
        // emitted scene would say so.
        RectD chart = LabelEdge();
        RectD pinned = new(chart.Left, 100.0, 40.0, 12.0);
        DateLabelResult result = DateLabelBuilder
            .TryBuild(Request(visible: pinned, full: new RectD(-500.0, 100.0, 600.0, 12.0)))
            .Result!;

        Assert.Contains(result.Warnings, warning => warning.Code == DateLabelBuilder.PositionChangedCode);
    }

    [Fact]
    public void A_clipped_fallback_that_keeps_the_requested_position_warns_about_nothing()
    {
        // The control on the warning above: a fallback that honoured the position it
        // was asked for must not emit the change warning, or every clipped date
        // label would report a change that never happened.
        DateLabelResult result = DateLabelBuilder
            .TryBuild(Request(full: new RectD(150.0, 100.0, 180.0, 12.0)))
            .Result!;

        Assert.DoesNotContain(result.Warnings, warning => warning.Code == DateLabelBuilder.PositionChangedCode);
    }

    /// <summary>
    /// Above and Below were retired by owner ruling 2026-09-30, so the date
    /// label's vertical cases are gone. Every remaining external position routes
    /// through the left/right cascade, which is what this pins: a stored
    /// "Above" no longer resolves to a bespoke vertical box, and cannot silently
    /// keep producing one.
    /// </summary>
    [Fact]
    public void A_retired_vertical_date_position_is_not_constructible()
    {
        // The enum member is gone, so the old assertion cannot even be written.
        // This is the compile-time half of the contract, expressed at runtime so
        // a re-introduction is a failing test rather than a silent capability.
        Assert.False(Enum.TryParse("Above", ignoreCase: false, out GanttLabelPosition _));
        Assert.False(Enum.TryParse("Below", ignoreCase: false, out GanttLabelPosition _));

        // And the surviving positions all reach the builder.
        foreach (GanttLabelPosition position in new[]
        {
            GanttLabelPosition.Left,
            GanttLabelPosition.Right,
            GanttLabelPosition.Inside,
        })
        {
            DateLabelResult outcome = DateLabelBuilder
                .TryBuild(Request(full: new RectD(150.0, 100.0, 180.0, 12.0)) with
                {
                    StartPosition = position,
                    FinishPosition = position,
                })
                .Result!;

            Assert.Equal(2, outcome.Primitives.Count);
        }
    }

    [Fact]
    public void An_explicit_inside_date_label_position_is_not_routed_through_the_left_branch()
    {
        // Inside is the third §22 position with its own geometry: the label is
        // centred on the visible rectangle. The old two-case branch sent it to the
        // left of the bar, which is exactly the position §22's Inside exists to
        // distinguish, so centring on the bar's centre is the property under test.
        //
        // The box is wider than the 40pt bar because the measured date is 100pt and
        // §23 forbids suppressing it. §22 allows an Inside label only when the text
        // fits the inner bounds, so containment within the bar is NOT asserted here
        // -- what must hold is that the box is centred on the bar rather than
        // displaced beside it.
        RectD visible = Visible();
        DateLabelResult result = DateLabelBuilder
            .TryBuild(Request(full: new RectD(150.0, 100.0, 180.0, 12.0)) with
            {
                StartPosition = GanttLabelPosition.Inside,
                FinishPosition = GanttLabelPosition.Inside,
            })
            .Result!;

        var centre = visible.Left + (visible.Width / 2);
        foreach (SceneText label in result.Primitives)
        {
            Assert.Equal(centre, label.TextBounds.Left + (label.TextBounds.Width / 2));
            Assert.False(
                label.TextBounds.Right <= visible.Left,
                $"The {label.PrimitiveId} Inside label sits left of the bar rather than on it.");
        }
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
