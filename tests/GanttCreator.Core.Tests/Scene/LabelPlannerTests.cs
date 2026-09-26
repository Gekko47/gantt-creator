using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

public sealed class LabelPlannerTests
{
    private static readonly GanttRowId _rowId = GanttRowId.New();
    private static readonly GanttRowId _laneId = GanttRowId.New();

    // Every character advances 4pt, so a 5-character label is exactly 20pt wide
    // and the tests can pin gap arithmetic by hand.
    private static readonly FakeTextMetrics _metrics = new(_ => 4.0, 10.0);

    // A 310pt-wide plot inside chart bounds that leave 6pt of margin on each side.
    private static readonly LabelMetrics _labelMetrics = new(
        new RectD(0, 0, 310, 100),
        new RectD(-6, -6, 322, 112),
        LabelGapPt: 3,
        LabelHeightPt: 10,
        MaximumExternalLabelWidthPt: 144
    );

    [Fact]
    public void Auto_places_a_span_label_to_the_right_of_the_shape()
    {
        LabelPlanResult result = Plan(Shape(100, 20, 50), "abcde");

        Assert.Equal(GanttLabelPosition.Right, result.Position);
        Assert.Equal(153, result.Bounds!.Value.Left);
        Assert.False(result.WasTruncated);
        Assert.Empty(result.Warnings);
    }

    [Fact]
    public void Auto_prefers_a_fitting_left_over_a_fitting_inside()
    {
        // ADR-0015 D1: a bar whose right edge is 7pt from the plot's right edge
        // cannot fit a 20pt label to its right, so Right is refused and Left is
        // chosen even though Inside (60pt) would also fit. This assertion is the
        // whole point of the amendment and fails against the superseded
        // Right -> Inside -> Left order, which would have chosen Inside.
        LabelPlanResult result = Plan(Shape(240, 20, 63), "abcde");

        Assert.Equal(GanttLabelPosition.Left, result.Position);
        // §22: the label's right edge sits at shape.left - LabelGapPt = 237, and
        // the 20pt-wide box is anchored to that edge.
        Assert.Equal(237, result.Bounds!.Value.Right);
        Assert.Equal(217, result.Bounds!.Value.Left);
    }

    [Fact]
    public void Auto_falls_back_to_inside_only_when_both_external_sides_are_blocked()
    {
        // A 100pt bar boxed in by occupants: the right one starts at 203, leaving
        // no room for a 20pt label, and the left one ends at 97, likewise. Inside
        // (100pt) is then the only position that can hold the text.
        LabelPlanResult result = Plan(
            Shape(100, 20, 100),
            "abcde",
            occupants:
            [
                new RectD(0, 20, 97, 10),
                new RectD(203, 20, 107, 10),
            ]
        );

        Assert.Equal(GanttLabelPosition.Inside, result.Position);
    }

    [Fact]
    public void Inside_is_rejected_by_the_cascade_when_the_text_does_not_fit_the_inner_bounds()
    {
        // A 10pt-wide bar boxed in on both sides so Inside is the only remaining
        // cascade candidate. Because 160pt of text cannot fit a 10pt interior,
        // the cascade refuses Inside outright rather than overflowing the body.
        LabelPlanResult result = Plan(
            Shape(100, 20, 10),
            new string('x', 40),
            occupants:
            [
                new RectD(0, 20, 97, 10),
                new RectD(113, 20, 197, 10),
            ]
        );

        // The widest-gap fallback then finds only the same 10pt interior, which
        // does hold an ellipsis, so the label is placed there truncated. The
        // rejection proven here is the cascade's, not the fallback's.
        Assert.Equal(GanttLabelPosition.Inside, result.Position);
        Assert.True(result.WasTruncated);
        Assert.Equal(LabelText.Ellipsis, result.Primitive!.Text[^1]);
    }

    [Fact]
    public void An_explicit_position_beats_the_auto_cascade()
    {
        // Right is requested explicitly on a bar sitting 10pt from the plot's
        // right edge. The 20pt text cannot fit the 7pt gap to its right, so the
        // cascade refuses it; the explicit position is still the only candidate
        // honoured, and the widest-gap measure across its own cascade truncates
        // rather than silently switching to Left or Inside.
        LabelPlanResult result = Plan(Shape(290, 20, 10), "abcde", GanttLabelPosition.Right);

        Assert.Equal(GanttLabelPosition.Right, result.Position);
        Assert.True(result.WasTruncated);
    }

    [Fact]
    public void The_delay_event_resolves_its_style_default_inside_once()
    {
        LabelRequest request = Request(
            Shape(100, 20, 200),
            "abcde",
            GanttLabelPosition.Auto,
            GanttEntityType.DelayEvent,
            alignment: GanttLabelPosition.Inside
        );

        LabelPlanResult result = Plan(request);

        // The style default is tried first, so a delay bar wide enough for the
        // text keeps its label inside the red body.
        Assert.Equal(GanttLabelPosition.Inside, result.Position);
        Assert.False(result.WasTruncated);
    }

    [Fact]
    public void The_delay_event_never_retries_inside_after_its_style_default_fails()
    {
        // A 10pt-wide delay bar: the style default Inside cannot fit 20pt of text.
        LabelRequest request = Request(
            Shape(100, 20, 10),
            "abcde",
            GanttLabelPosition.Auto,
            GanttEntityType.DelayEvent,
            alignment: GanttLabelPosition.Inside
        );

        LabelPlanResult result = Plan(request);

        // ADR-0015 D3: the escape cascade is Right then Left, and Inside is never
        // re-entered, so the position that just failed cannot succeed on a
        // second pass through the same evaluation.
        Assert.NotEqual(GanttLabelPosition.Inside, result.Position);
        Assert.Equal(GanttLabelPosition.Right, result.Position);
    }

    [Fact]
    public void The_delay_event_escapes_to_right_then_left()
    {
        // Right is blocked by an occupant starting at 120, so the escape cascade
        // must continue to Left rather than retrying the rejected Inside.
        LabelRequest request = Request(
            Shape(100, 20, 10),
            "abcde",
            GanttLabelPosition.Auto,
            GanttEntityType.DelayEvent,
            alignment: GanttLabelPosition.Inside
        );

        LabelPlanResult result = Plan(request, [new RectD(120, 20, 6, 10)]);

        Assert.Equal(GanttLabelPosition.Left, result.Position);
    }

    [Fact]
    public void The_delay_event_switches_text_colour_between_inside_and_outside()
    {
        // §17: the delay's text is DelayText inside the red body and DefaultText
        // outside it, so the same request must resolve to two different styles
        // depending only on which side the label lands.
        SceneStyle inside = new("Delay", strokeColour: ColourHex.Parse("#FF0000"), alignment: GanttLabelPosition.Inside);
        SceneStyle outside = new("Default");

        // Wide enough for the 20pt text inside, so the style default applies.
        LabelPlanResult inner = Plan(
            Request(
                Shape(100, 20, 200),
                "abcde",
                GanttLabelPosition.Auto,
                GanttEntityType.DelayEvent,
                alignment: GanttLabelPosition.Inside,
                outsideStyle: outside,
                textStyle: inside
            )
        );
        // Too narrow for the text inside, so it escapes to the right of the body.
        LabelPlanResult outer = Plan(
            Request(
                Shape(100, 20, 10),
                "abcde",
                GanttLabelPosition.Auto,
                GanttEntityType.DelayEvent,
                alignment: GanttLabelPosition.Inside,
                outsideStyle: outside,
                textStyle: inside
            )
        );

        Assert.Equal(GanttLabelPosition.Inside, inner.Position);
        Assert.Same(inside, inner.Primitive!.Style);

        Assert.Equal(GanttLabelPosition.Right, outer.Position);
        Assert.Same(outside, outer.Primitive!.Style);
    }

    [Fact]
    public void A_delay_with_an_explicit_position_ignores_the_style_default()
    {
        // §22 precedence: an explicit row value outranks the named-style default,
        // so a delay that names Right must not be moved back inside the body by
        // the delay style default.
        LabelPlanResult result = Plan(
            Request(
                Shape(100, 20, 200),
                "abcde",
                GanttLabelPosition.Right,
                GanttEntityType.DelayEvent,
                alignment: GanttLabelPosition.Inside
            )
        );

        Assert.Equal(GanttLabelPosition.Right, result.Position);
    }

    [Fact]
    public void A_milestone_never_uses_inside_and_falls_through_to_above()
    {
        // §20/ADR-0015 D2: a milestone's Auto order is Right → Left → Above →
        // Below, so a milestone boxed in on both sides lands Above and never
        // considers the Inside position a span would use.
        LabelPlanResult result = Plan(
            Request(
                Shape(100, 20, 20),
                "abcde",
                GanttLabelPosition.Auto,
                GanttEntityType.AsPlannedMilestone
            ),
            [
                new RectD(0, 20, 97, 10),
                new RectD(123, 20, 187, 10),
            ]
        );

        Assert.Equal(GanttLabelPosition.Above, result.Position);
    }

    [Fact]
    public void An_inside_label_is_centred_in_the_shape()
    {
        // §22: Inside is centred in the visible rectangle, so the fitted 20pt text
        // in a 100pt bar starts 40pt in rather than flush against the bar's edge.
        LabelPlanResult result = Plan(Shape(100, 20, 100), "abcde", GanttLabelPosition.Inside);

        Assert.Equal(GanttLabelPosition.Inside, result.Position);
        Assert.Equal(140, result.Bounds!.Value.Left);
        Assert.Equal(160, result.Bounds.Value.Right);
    }

    [Fact]
    public void An_above_label_sits_one_gap_above_the_shape_top_edge()
    {
        // §22: "Above: horizontally centred; label bottom = shape top −
        // LabelGapPt", so the 10pt label spans 7..17 on a bar whose top is 20.
        LabelPlanResult result = Plan(Shape(100, 20, 100), "abcde", GanttLabelPosition.Above);

        Assert.Equal(GanttLabelPosition.Above, result.Position);
        Assert.Equal(17, result.Bounds!.Value.Bottom);
        Assert.Equal(7, result.Bounds.Value.Top);
        // Horizontally centred: 100 + ((100 - 20) / 2).
        Assert.Equal(140, result.Bounds.Value.Left);
    }

    [Fact]
    public void A_below_label_sits_one_gap_below_the_shape_bottom_edge()
    {
        // §22: "Below: horizontally centred; label top = shape bottom + LabelGapPt",
        // and the bar's bottom is 28, so the label starts at 31.
        LabelPlanResult result = Plan(Shape(100, 20, 100), "abcde", GanttLabelPosition.Below);

        Assert.Equal(GanttLabelPosition.Below, result.Position);
        Assert.Equal(31, result.Bounds!.Value.Top);
        Assert.Equal(140, result.Bounds.Value.Left);
    }

    [Fact]
    public void The_widest_gap_fallback_anchors_a_left_label_to_the_gap_boundary()
    {
        // A 10pt bar at 250 leaves a 247pt gap to its left, which the 144pt
        // external maximum caps. The truncated label is anchored by its right edge
        // to shape.left − LabelGapPt = 247, exactly as an untruncated Left label
        // would be, rather than starting at the far end of the whole gap.
        LabelPlanResult result = Plan(Shape(250, 20, 10), new string('x', 40));

        Assert.Equal(GanttLabelPosition.Left, result.Position);
        Assert.True(result.WasTruncated);
        Assert.Equal(247, result.Bounds!.Value.Right);
        Assert.Equal(103, result.Bounds.Value.Left);
    }

    [Fact]
    public void The_widest_gap_fallback_takes_the_interior_when_it_is_the_widest()
    {
        // Box the bar in so no cascade position can hold the full 40 characters
        // (160pt): Right is pinned to a 4pt gap at 58, Left to a 4pt gap ending
        // at 42, and the interior is 10pt. The interior is the widest gap, so
        // ADR-0015 D4 puts the truncated label there, and §22's rule that Inside
        // is only used when the text fits does not apply to the fallback, whose
        // whole purpose is to cut the text to the space available.
        LabelPlanResult result = Plan(
            Shape(45, 20, 10),
            new string('x', 40),
            occupants:
            [
                new RectD(0, 20, 38, 10),
                new RectD(58, 20, 252, 10),
            ]
        );

        Assert.Equal(GanttLabelPosition.Inside, result.Position);
        Assert.True(result.WasTruncated);
        Assert.Equal(LabelText.Ellipsis, result.Primitive!.Text[^1]);
        Assert.Equal(LabelPlanner.TruncatedToFitCode, Assert.Single(result.Warnings).Code);
    }

    [Fact]
    public void The_widest_gap_fallback_takes_the_widest_of_the_three_positions()
    {
        // Pin the right side shut with an occupant starting at 303, so Right has
        // no usable gap while Left keeps 237pt and Inside only 10pt.
        LabelPlanResult result = Plan(
            Shape(290, 20, 10),
            "abcde",
            occupants: [new RectD(303, 20, 7, 10)]
        );

        // The widest gap is Left, so the label goes left and fits uncut.
        Assert.Equal(GanttLabelPosition.Left, result.Position);
        Assert.False(result.WasTruncated);
    }

    [Fact]
    public void The_widest_gap_fallback_breaks_an_equal_gap_tie_on_the_cascade_order()
    {
        // A 50pt bar centred in a 310pt plot leaves equal ~128pt gaps on both
        // sides, so the Right-first cascade order must win the tie.
        LabelPlanResult result = Plan(Shape(130, 20, 50), new string('x', 30));

        Assert.Equal(GanttLabelPosition.Right, result.Position);
    }

    [Fact]
    public void The_widest_gap_fallback_never_places_above_the_chart_top_edge()
    {
        // A 200pt milestone band whose top edge sits 5pt below the chart's top
        // edge. The 60-character text is 240pt, so no cascade position holds it
        // and the ADR-0015 widest-gap measure runs: Above and Below each offer the
        // shape's full 200pt, which beats the 97pt Left gap, and the cascade order
        // tries Above first. §22's "Above: label bottom = shape top − LabelGapPt"
        // puts that box at 5 − 3 − 10 = −8, i.e. above the chart's top edge of 0 —
        // and the gap measure is blind to the chart, so containment is the only
        // thing that keeps the label on the panel. Below is chosen instead.
        var topEdgeMetrics = _labelMetrics with { ChartBounds = new RectD(-6, 0, 322, 106) };

        LabelPlanResult result = PlanOutcome(
                Request(
                    Shape(100, 5, 200),
                    new string('x', 60),
                    GanttLabelPosition.Auto,
                    GanttEntityType.AsPlannedMilestone
                ),
                topEdgeMetrics
            )
            .Result!;

        Assert.Equal(GanttLabelPosition.Below, result.Position);
        Assert.True(result.WasTruncated);
        // Below: label top = shape bottom + LabelGapPt = 13 + 3 = 16.
        Assert.Equal(16, result.Bounds!.Value.Top);
        Assert.True(result.Bounds.Value.Top >= topEdgeMetrics.ChartBounds.Top);
    }

    [Fact]
    public void The_widest_gap_fallback_never_places_below_the_chart_bottom_edge()
    {
        // The mirror case, and the one that proves the containment check rather
        // than the tie-break: the band sits near the chart's bottom, so Below
        // would land at 96 + 3 = 99 and end at 109, past the chart's bottom edge
        // of 106. Above is blocked by an in-lane occupant, so before the fix the
        // 200pt Below gap was selected and the label was drawn off the panel.
        var bottomEdgeMetrics = _labelMetrics with { ChartBounds = new RectD(-6, 0, 322, 106) };

        LabelPlanResult result = PlanOutcome(
                Request(
                    Shape(100, 88, 200),
                    new string('x', 60),
                    GanttLabelPosition.Auto,
                    GanttEntityType.AsPlannedMilestone
                ),
                bottomEdgeMetrics,
                [new RectD(100, 70, 200, 10)]
            )
            .Result!;

        Assert.NotEqual(GanttLabelPosition.Below, result.Position);
        Assert.Equal(GanttLabelPosition.Left, result.Position);
        Assert.True(result.Bounds!.Value.Bottom <= bottomEdgeMetrics.ChartBounds.Bottom);
    }

    [Fact]
    public void Suppresses_the_label_with_one_warning_when_no_gap_holds_an_ellipsis()
    {
        // A 10pt-wide bar in a 120pt plot, boxed in so that neither side leaves
        // room for even the 4pt single-character ellipsis. Right would start at
        // 113 and Left would end at 97, so both are pinned shut.
        LabelPlanResult result = Plan(
            Shape(50, 20, 10),
            "abcde",
            occupants:
            [
                new RectD(40, 20, 10, 10),
                new RectD(63, 20, 67, 10),
            ]
        );

        Assert.Null(result.Primitive);
        SceneWarning warning = Assert.Single(result.Warnings);
        Assert.Equal(LabelPlanner.SuppressedNoSpaceCode, warning.Code);
    }

    [Fact]
    public void A_candidate_blocked_by_a_higher_priority_label_is_rejected()
    {
        // The single free candidate (Right) is occupied, so the plan falls to the
        // widest-gap measure, which must not hand back the blocked space.
        LabelPlanResult result = Plan(
            Shape(100, 20, 10),
            "abcde",
            occupants:
            [
                new RectD(113, 20, 4, 10),
                new RectD(0, 20, 98, 10),
            ]
        );

        Assert.NotEqual(GanttLabelPosition.Right, result.Position);
    }

    [Fact]
    public void Truncation_never_splits_a_surrogate_pair()
    {
        // Each emoji is two UTF-16 code units, so an odd-length cut would land
        // between a high and low surrogate. 4pt per code unit is 8pt per emoji.
        var text = string.Concat(Enumerable.Repeat("\U0001F600", 6));
        var wideningMetrics = new FakeTextMetrics(
            character => char.IsLowSurrogate(character) ? 0 : 4.0,
            10.0
        );

        var truncated = LabelText.Ellipsize(text, 20, wideningMetrics);

        Assert.False(char.IsHighSurrogate(truncated[^2]), "The truncation must not leave an orphaned high surrogate.");
        Assert.False(char.IsLowSurrogate(truncated[^1]), "The truncation must not leave an orphaned low surrogate.");
        Assert.Equal(LabelText.Ellipsis, truncated[^1]);
    }

    [Fact]
    public void Ellipsize_truncates_the_text_and_appends_the_marker()
    {
        // Five characters need 20pt, but the reserved ellipsis means only four
        // fit inside a 20pt box, so the result is the text cut to fit with the
        // single-character marker appended.
        var truncated = LabelText.Ellipsize("abcde", 20, _metrics);

        Assert.Equal("abcd…", truncated);
    }

    [Fact]
    public void Ellipsize_returns_the_bare_marker_when_nothing_fits()
    {
        var truncated = LabelText.Ellipsize("abcde", 2, _metrics);

        Assert.Equal(LabelText.Ellipsis.ToString(), truncated);
    }

    [Fact]
    public void Emits_the_label_at_the_label_layer_with_a_role_derived_id()
    {
        LabelPlanResult result = Plan(Shape(100, 20, 50), "abcde");

        SceneText label = Assert.IsType<SceneText>(result.Primitive);
        Assert.Equal(ZLayer.Label, label.ZLayer);
        Assert.EndsWith(":label", label.PrimitiveId, StringComparison.Ordinal);
    }

    [Fact]
    public void A_blank_description_produces_no_label_and_no_warning()
    {
        LabelPlanResult result = Plan(Shape(100, 20, 50), "   ");

        Assert.Null(result.Primitive);
        Assert.Empty(result.Warnings);
    }

    [Fact]
    public void The_none_position_produces_no_label()
    {
        LabelPlanResult result = Plan(Shape(100, 20, 50), "abcde", GanttLabelPosition.None);

        Assert.Null(result.Primitive);
        Assert.Empty(result.Warnings);
    }

    [Fact]
    public void Placement_is_identical_across_repeated_runs()
    {
        for (var run = 0; run < 3; run++)
        {
            LabelPlanResult result = Plan(Shape(130, 20, 50), new string('x', 30));

            Assert.Equal(GanttLabelPosition.Right, result.Position);
            // Right starts at shape.right + LabelGapPt = 180 + 3.
            Assert.Equal(183, result.Bounds!.Value.Left);
        }
    }

    [Fact]
    public void Refuses_a_null_request()
    {
        LabelPlanCreationOutcome outcome = LabelPlanner.TryPlan(null, _labelMetrics);

        Assert.False(outcome.Succeeded);
        Assert.Equal(LabelRefusal.NullRequest, outcome.Refusal);
    }

    [Fact]
    public void Refuses_null_metrics()
    {
        LabelPlanCreationOutcome outcome = LabelPlanner.TryPlan(Request(Shape(100, 20, 50), "abcde"), null);

        Assert.False(outcome.Succeeded);
        Assert.Equal(LabelRefusal.NullMetrics, outcome.Refusal);
    }

    [Fact]
    public void Refuses_an_undefined_position()
    {
        LabelPlanCreationOutcome outcome = LabelPlanner.TryPlan(
            Request(Shape(100, 20, 50), "abcde", (GanttLabelPosition)999),
            _labelMetrics
        );

        Assert.False(outcome.Succeeded);
        Assert.Equal(LabelRefusal.InvalidPosition, outcome.Refusal);
    }

    [Fact]
    public void Refuses_shape_bounds_that_cannot_be_finite() =>
        // RectD itself rejects a non-finite coordinate at construction, so the
        // planner's own guard is a defence in depth against a future caller
        // that bypasses the value object. The value object is the contract.
        _ = Assert.Throws<ArgumentOutOfRangeException>(() => new RectD(double.NaN, 20, 50, 8));

    [Fact]
    public void Refuses_a_non_finite_gap()
    {
        LabelPlanCreationOutcome outcome = LabelPlanner.TryPlan(
            Request(Shape(100, 20, 50), "abcde"),
            _labelMetrics with
            {
                LabelGapPt = double.PositiveInfinity,
            }
        );

        Assert.False(outcome.Succeeded);
        Assert.Equal(LabelRefusal.InvalidMetrics, outcome.Refusal);
    }

    [Fact]
    public void Refuses_a_negative_label_height()
    {
        LabelPlanCreationOutcome outcome = LabelPlanner.TryPlan(
            Request(Shape(100, 20, 50), "abcde"),
            _labelMetrics with
            {
                LabelHeightPt = -1,
            }
        );

        Assert.False(outcome.Succeeded);
        Assert.Equal(LabelRefusal.InvalidMetrics, outcome.Refusal);
    }

    private static RectD Shape(double x, double y, double width) => new(x, y, width, 8);

    private static LabelPlanCreationOutcome PlanOutcome(
        LabelRequest request,
        IReadOnlyList<RectD>? occupants = null
    ) => LabelPlanner.TryPlan(request, _labelMetrics, occupants);

    private static LabelPlanCreationOutcome PlanOutcome(LabelRequest request, LabelMetrics metrics) =>
        LabelPlanner.TryPlan(request, metrics);

    private static LabelPlanCreationOutcome PlanOutcome(
        LabelRequest request,
        LabelMetrics metrics,
        IReadOnlyList<RectD> occupants
    ) => LabelPlanner.TryPlan(request, metrics, occupants);

    private static LabelPlanResult Plan(LabelRequest request, IReadOnlyList<RectD>? occupants = null) =>
        PlanOutcome(request, occupants).Result!;

    private static LabelPlanResult Plan(
        RectD shape,
        string text,
        GanttLabelPosition position = GanttLabelPosition.Auto,
        IReadOnlyList<RectD>? occupants = null
    ) => Plan(Request(shape, text, position), occupants);

    private static LabelPlanResult Plan(
        RectD shape,
        string text,
        IReadOnlyList<RectD>? occupants
    ) => PlanOutcome(Request(shape, text), occupants).Result!;

    private static LabelRequest Request(
        RectD shape,
        string text,
        GanttLabelPosition position = GanttLabelPosition.Auto,
        GanttEntityType type = GanttEntityType.AsPlannedActivity,
        GanttLabelPosition? alignment = null,
        SceneStyle? outsideStyle = null,
        SceneStyle? textStyle = null
    ) =>
        new(
            new GanttEvent(
                1,
                _rowId,
                _laneId,
                null,
                type,
                text,
                new DateOnly(2024, 1, 5),
                new DateOnly(2024, 1, 9),
                null,
                null,
                null,
                null,
                null,
                true,
                null
            ),
            text,
            position,
            shape,
            textStyle ?? new SceneStyle("Label", fontFamily: "Aptos", fontSizePt: 8, alignment: alignment),
            _metrics,
            outsideStyle
        );
}
