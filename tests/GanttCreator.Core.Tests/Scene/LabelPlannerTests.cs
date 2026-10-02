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
        // One worksheet row (owner ruling). It was 10, a single-line height, which
        // produced a label box shorter than the row it sat in.
        RowHeightPt: 18,
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
            styleDefaultPosition: GanttLabelPosition.Inside
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
            styleDefaultPosition: GanttLabelPosition.Inside
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
            styleDefaultPosition: GanttLabelPosition.Inside
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
        SceneStyle inside = new("Delay", strokeColour: ColourHex.Parse("#FF0000"));
        SceneStyle outside = new("Default");

        // Wide enough for the 20pt text inside, so the style default applies.
        LabelPlanResult inner = Plan(
            Request(
                Shape(100, 20, 200),
                "abcde",
                GanttLabelPosition.Auto,
                GanttEntityType.DelayEvent,
                styleDefaultPosition: GanttLabelPosition.Inside,
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
                styleDefaultPosition: GanttLabelPosition.Inside,
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
                styleDefaultPosition: GanttLabelPosition.Inside
            )
        );

        Assert.Equal(GanttLabelPosition.Right, result.Position);
    }

    [Fact]
    public void A_milestone_auto_order_is_right_then_left_and_never_inside()
    {
        // §20/ADR-0015 D2: a milestone's Auto order is Right → Left. Above and
        // Below were the old third and fourth entries and are retired (owner
        // ruling 2026-09-30), so a milestone boxed in on both sides has no
        // vertical escape and must fall back rather than reach for one.
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

        // With the vertical positions gone there is no third option, so a
        // milestone boxed in on both sides SUPPRESSES its label with one warning
        // rather than reaching above or below the shape. That is the ruling's
        // intended consequence: a fixed lane height cannot guarantee vertical
        // room, so vertical placement is refused rather than approximated.
        Assert.Null(result.Position);
        Assert.Null(result.Primitive);
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

    /// <summary>
    /// Above and Below were retired by owner ruling 2026-09-30. This pins that
    /// they are not merely denied by the catalogue but are <em>not parseable</em>
    /// from a stored cell value — which is what makes ADR-0029 D6's "reported,
    /// never coerced" achievable, since there is no value left to coerce.
    /// </summary>
    [Theory]
    [InlineData("Above")]
    [InlineData("Below")]
    public void A_retired_label_position_name_does_not_parse(string stored)
    {
        Assert.False(Enum.TryParse(stored, ignoreCase: false, out GanttLabelPosition _));
        Assert.DoesNotContain(
            Enum.GetValues<GanttLabelPosition>(),
            position => string.Equals(position.ToString(), stored, StringComparison.Ordinal));
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
    public void The_widest_gap_fallback_cannot_leave_the_chart_vertically()
    {
        // This scenario is WHY Above and Below were retired (owner ruling
        // 2026-09-30): a 200pt band whose top sits 5pt below the chart's top edge
        // let the vertical positions place a label at 5 − 3 − 10 = −8, outside the
        // chart. With those positions gone the fallback can only choose Left or
        // Right, so vertical escape is now unrepresentable rather than merely
        // defended against by a containment check.
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

        Assert.True(result.Position is GanttLabelPosition.Left or GanttLabelPosition.Right);
        Assert.True(result.Bounds!.Value.Bottom <= topEdgeMetrics.ChartBounds.Bottom);
        Assert.True(result.Bounds.Value.Top >= topEdgeMetrics.ChartBounds.Top);
    }

    [Fact]
    public void The_widest_gap_fallback_never_places_below_the_chart_bottom_edge()
    {
        // The mirror case, near the chart's bottom. The result must stay inside the
        // chart AND must be a horizontal position.
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

        Assert.True(result.Position is GanttLabelPosition.Left or GanttLabelPosition.Right);
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
                RowHeightPt = -1,
            }
        );

        Assert.False(outcome.Succeeded);
        Assert.Equal(LabelRefusal.InvalidMetrics, outcome.Refusal);
    }

    private static RectD Shape(double x, double y, double width) => new(x, y, width, 8);

    /// <summary>
    /// The label box is one worksheet row tall and centred on the row.
    /// </summary>
    [Fact]
    public void The_label_box_is_one_row_tall_and_centred_on_the_shape()
    {
        // A bar 8pt tall inside an 18pt row, offset so centring cannot pass by
        // coincidence. The owner ruled the box height is one row, not the 10pt
        // single-line height it used to be.
        RectD shape = new(100, 40, 50, 8);

        LabelPlanResult result = Plan(shape, "abcde", GanttLabelPosition.Right);

        Assert.Equal(18, result.Bounds!.Value.Height);
        Assert.Equal(shape.Top + ((shape.Height - 18) / 2), result.Bounds.Value.Top);

        // Centred on the shape, so it overhangs the 8pt bar by 5pt on each side.
        // This is the intent of a row-height box, not an error to be corrected.
        Assert.Equal(shape.Top - 5, result.Bounds.Value.Top);
        Assert.Equal(shape.Bottom + 5, result.Bounds.Value.Bottom);
    }

    /// <summary>
    /// A row-height box centred on an 8pt bar overhangs the bar vertically, and it
    /// must still be placed - the overhang is the point of centring on the ROW.
    /// </summary>
    /// <remarks>
    /// This is the regression test for the failure the row-height change caused. The
    /// box is 5pt taller than the bar on each side, so a collision test that
    /// compared the candidate against the shape's own rectangle would reject every
    /// position and suppress the label outright. The occupants passed here are the
    /// row's own bar bounds, which is what <c>SceneBuilder</c> seeds for a row whose
    /// description label was placed first.
    /// </remarks>
    [Fact]
    public void A_row_height_box_is_not_blocked_by_the_owning_shape_it_is_centred_on()
    {
        RectD shape = new(100, 40, 50, 8);

        LabelPlanResult result = Plan(shape, "abcde", GanttLabelPosition.Right, [shape]);

        Assert.Equal(GanttLabelPosition.Right, result.Position);
        Assert.False(result.WasTruncated);
        Assert.Equal("abcde", result.Primitive!.Text);
    }

    /// <summary>
    /// A stack sibling's label box does not block this row's label (ADR-0033 D2).
    /// </summary>
    /// <remarks>
    /// <para>
    /// The reported defect: two events on one lane are stacked, each with its own
    /// description label, and both labels came out as unreadable stubs. Once the
    /// label box became a full row tall and centred, the two boxes necessarily
    /// overlap - the slots are only a few points apart - so <c>Blocked</c> rejected
    /// every candidate and both fell through to the ADR-0015 widest-gap fallback,
    /// which ellipsised them to almost nothing.
    /// </para>
    /// <para>
    /// A stack is a deliberate layout, not an overlap, so the owner's ruling is that
    /// it must not be treated as a collision. The occupant below is in the SAME lane
    /// and a DIFFERENT slot, and sits exactly where the candidate's box would land.
    /// </para>
    /// </remarks>
    [Fact]
    public void A_stack_siblings_label_does_not_block_this_rows_label()
    {
        RectD shape = new(100, 40, 50, 8);

        // The sibling's box occupies the whole Right candidate: same lane, other slot.
        LabelOccupant sibling = new(new RectD(153, 35, 60, 18), LaneOrder: 2, StackIndex: 1);

        LabelPlanResult result = PlanOutcome(
            Request(shape, "abcde", GanttLabelPosition.Right) with
            {
                LaneOrder = 2,
                StackIndex = 0,
            },
            [sibling]).Result!;

        Assert.Equal(GanttLabelPosition.Right, result.Position);
        Assert.False(result.WasTruncated);
        Assert.Equal("abcde", result.Primitive!.Text);
    }

    /// <summary>
    /// The exemption is narrow: a same-slot occupant in the same lane still blocks.
    /// </summary>
    [Fact]
    public void A_same_slot_occupant_in_the_same_lane_still_blocks()
    {
        RectD shape = new(100, 40, 50, 8);
        LabelOccupant sameSlot = new(new RectD(153, 35, 60, 18), LaneOrder: 2, StackIndex: 0);

        LabelPlanResult result = PlanOutcome(
            Request(shape, "abcde", GanttLabelPosition.Right) with
            {
                LaneOrder = 2,
                StackIndex = 0,
            },
            [sameSlot]).Result!;

        // A blocked candidate cannot come back clean at its own position. The exact
        // fallback (truncated at the widest gap, or suppressed entirely) is a
        // placement detail; what matters is that the exemption did NOT fire.
        Assert.True(
            result.WasTruncated || result.Position != GanttLabelPosition.Right,
            "A same-slot collision must still prevent a clean Right placement.");
        Assert.NotEqual("abcde", result.Primitive?.Text);
    }

    /// <summary>
    /// The exemption is narrow: an occupant in a different lane still blocks.
    /// </summary>
    [Fact]
    public void An_occupant_in_a_different_lane_still_blocks()
    {
        RectD shape = new(100, 40, 50, 8);
        LabelOccupant otherLane = new(new RectD(153, 35, 60, 18), LaneOrder: 7, StackIndex: 1);

        LabelPlanResult result = PlanOutcome(
            Request(shape, "abcde", GanttLabelPosition.Right) with
            {
                LaneOrder = 2,
                StackIndex = 0,
            },
            [otherLane]).Result!;

        // Same invariant as the same-slot case: the cross-lane occupant still blocks, so
        // the full text cannot be placed cleanly at Right.
        Assert.True(
            result.WasTruncated || result.Position != GanttLabelPosition.Right,
            "A cross-lane collision must still prevent a clean Right placement.");
        Assert.NotEqual("abcde", result.Primitive?.Text);
    }

    /// <summary>
    /// How much a stacked lane's labels now overlap, measured rather than asserted.
    /// </summary>
    /// <remarks>
    /// This is the diagnostic the owner asked for. With the box at one row tall and
    /// centred, two stacked slots a few points apart produce boxes that overlap
    /// VERTICALLY by design - that is what centring on the row means. What the
    /// exemption guarantees is that this overlap no longer truncates or suppresses
    /// anything: both labels are emitted whole. The horizontal separation is what
    /// keeps them legible, so it is reported here rather than left implicit.
    /// </remarks>
    [Fact]
    public void Stacked_labels_overlap_vertically_by_design_but_are_both_emitted_whole()
    {
        // A realistic 18pt row split into two slots with a 2pt gap: 5pt each, which is
        // the geometry that produced the unreadable stubs.
        var slotHeight = 5.0;
        RectD upperShape = new(100, 40, 50, slotHeight);
        RectD lowerShape = new(200, 40 + slotHeight + 2, 50, slotHeight);

        LabelPlanResult upper = PlanOutcome(
            Request(upperShape, "upper", GanttLabelPosition.Right) with { LaneOrder = 1, StackIndex = 0 },
            []).Result!;
        LabelPlanResult lower = PlanOutcome(
            Request(lowerShape, "lower", GanttLabelPosition.Right) with { LaneOrder = 1, StackIndex = 1 },
            [new LabelOccupant(upper.Bounds!.Value, LaneOrder: 1, StackIndex: 0)]).Result!;

        Assert.Equal("upper", upper.Primitive!.Text);
        Assert.Equal("lower", lower.Primitive!.Text);
        Assert.False(upper.WasTruncated);
        Assert.False(lower.WasTruncated);

        // The overlap is real and expected; report it rather than assert it away.
        RectD upperBox = upper.Bounds!.Value;
        RectD lowerBox = lower.Bounds!.Value;
        double verticalOverlap =
            Math.Min(upperBox.Bottom, lowerBox.Bottom) - Math.Max(upperBox.Top, lowerBox.Top);

        Assert.True(
            verticalOverlap > 0,
            "Row-height boxes in adjacent slots are expected to overlap vertically.");

        // Horizontal separation is what keeps them readable, so assert it is real
        // rather than the boxes sitting on top of each other.
        Assert.True(
            upperBox.Right <= lowerBox.Left || lowerBox.Right <= upperBox.Left,
            $"Stacked labels must not share horizontal space: {upperBox} and {lowerBox}.");

        System.Console.WriteLine(
            $"stacked label boxes overlap {verticalOverlap:0.##}pt vertically and are separated horizontally");
    }

    private static LabelPlanCreationOutcome PlanOutcome(
        LabelRequest request,
        IReadOnlyList<LabelOccupant>? occupants = null
    ) => LabelPlanner.TryPlan(request, _labelMetrics, occupants);

    private static LabelPlanCreationOutcome PlanOutcome(LabelRequest request, LabelMetrics metrics) =>
        LabelPlanner.TryPlan(request, metrics);

    private static LabelPlanCreationOutcome PlanOutcome(
        LabelRequest request,
        LabelMetrics metrics,
        IReadOnlyList<LabelOccupant> occupants
    ) => LabelPlanner.TryPlan(request, metrics, occupants);

    private static LabelPlanResult Plan(LabelRequest request, IReadOnlyList<LabelOccupant>? occupants = null) =>
        PlanOutcome(request, occupants).Result!;

    private static LabelPlanResult Plan(
        RectD shape,
        string text,
        GanttLabelPosition position = GanttLabelPosition.Auto,
        IReadOnlyList<LabelOccupant>? occupants = null
    ) => Plan(Request(shape, text, position), occupants);

    private static LabelPlanResult Plan(
        RectD shape,
        string text,
        IReadOnlyList<LabelOccupant>? occupants
    ) => PlanOutcome(Request(shape, text), occupants).Result!;

    [Theory]
    [InlineData(GanttLabelPosition.Left, GanttTextAlignment.Right)]
    [InlineData(GanttLabelPosition.Right, GanttTextAlignment.Left)]
    [InlineData(GanttLabelPosition.Inside, GanttTextAlignment.Centre)]
    public void The_emitted_alignment_follows_the_product_rule_not_the_position_enum(
        GanttLabelPosition position,
        GanttTextAlignment expected)
    {
        // ADR-0018 D4. §22 fixes the label *box* but never says how text sits
        // inside it, so this mapping is a product decision: text is aligned to
        // end against the shape (the §24 "hugs the line" rule generalised), and
        // the centred placements of §22 centre their text. Before the retype
        // the planner passed the raw `GanttLabelPosition` straight through as
        // the text alignment, which this pins.
        LabelPlanResult result = Plan(Shape(100, 20, 200), "abcde", position);

        Assert.Equal(position, result.Position);
        Assert.NotNull(result.Primitive);
        Assert.Equal(expected, result.Primitive!.Alignment);
    }

    [Fact]
    public void The_delay_label_keeps_its_own_text_colour_inside_the_body()
    {
        // Section 17: a delay label inside the red body is DelayText (white). The
        // planner must carry the inside style through unchanged.
        LabelPlanResult result = Plan(Request(
            Shape(100, 20, 200),
            "abcde",
            GanttLabelPosition.Auto,
            GanttEntityType.DelayEvent,
            styleDefaultPosition: GanttLabelPosition.Inside,
            textStyle: new SceneStyle("DelayText", textColour: ColourHex.Parse("#FFFFFF")),
            outsideStyle: new SceneStyle("DefaultText", textColour: ColourHex.Parse("#000000"))));

        Assert.Equal(GanttLabelPosition.Inside, result.Position);
        Assert.Equal("#FFFFFF", result.Primitive!.Style.TextColour!.ToString());
    }

    [Fact]
    public void The_delay_label_switches_to_the_outside_text_colour_when_it_leaves_the_body()
    {
        // The other half of section 17's switch, and the row that was unreachable
        // while no caller supplied an outside style at all. A 10pt delay bar cannot
        // hold 20pt of text, so the label escapes the body and must stop being
        // white - white on the chart background is unreadable.
        LabelPlanResult result = Plan(Request(
            Shape(100, 20, 10),
            "abcde",
            GanttLabelPosition.Auto,
            GanttEntityType.DelayEvent,
            styleDefaultPosition: GanttLabelPosition.Inside,
            textStyle: new SceneStyle("DelayText", textColour: ColourHex.Parse("#FFFFFF")),
            outsideStyle: new SceneStyle("DefaultText", textColour: ColourHex.Parse("#000000"))));

        Assert.NotEqual(GanttLabelPosition.Inside, result.Position);
        Assert.Equal("#000000", result.Primitive!.Style.TextColour!.ToString());
    }

    [Fact]
    public void A_non_delay_label_keeps_its_own_text_colour_outside_the_body()
    {
        // The control for the row above. The switch is a property of a delay event,
        // so a planned label placed outside its body must NOT silently adopt the
        // outside style: that would recolour every label in the chart.
        LabelPlanResult result = Plan(Request(
            Shape(100, 20, 10),
            "abcde",
            GanttLabelPosition.Auto,
            GanttEntityType.AsPlannedActivity,
            textStyle: new SceneStyle("AsPlannedActivity", textColour: ColourHex.Parse("#123456")),
            outsideStyle: new SceneStyle("DefaultText", textColour: ColourHex.Parse("#000000"))));

        Assert.NotEqual(GanttLabelPosition.Inside, result.Position);
        Assert.Equal("#123456", result.Primitive!.Style.TextColour!.ToString());
    }

    [Fact]
    public void A_delay_label_with_no_outside_style_keeps_its_own_text_colour()
    {
        // The documented fallback when OutsideTextStyle is absent. It degrades to the
        // inside style rather than substituting a colour, which is the same
        // no-substitution rule the fill and stroke members follow.
        LabelPlanResult result = Plan(Request(
            Shape(100, 20, 10),
            "abcde",
            GanttLabelPosition.Auto,
            GanttEntityType.DelayEvent,
            styleDefaultPosition: GanttLabelPosition.Inside,
            textStyle: new SceneStyle("DelayText", textColour: ColourHex.Parse("#FFFFFF"))));

        Assert.NotEqual(GanttLabelPosition.Inside, result.Position);
        Assert.Equal("#FFFFFF", result.Primitive!.Style.TextColour!.ToString());
    }

    private static LabelRequest Request(
        RectD shape,
        string text,
        GanttLabelPosition position = GanttLabelPosition.Auto,
        GanttEntityType type = GanttEntityType.AsPlannedActivity,
        GanttLabelPosition? styleDefaultPosition = null,
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
            textStyle ?? new SceneStyle("Label", fontFamily: "Aptos", fontSizePt: 8),
            _metrics,
            outsideStyle,
            StyleDefaultPosition: styleDefaultPosition ?? GanttLabelPosition.Inside
        );
}
