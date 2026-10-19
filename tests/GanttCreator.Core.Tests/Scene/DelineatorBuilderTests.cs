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

    // The predetermined description-label style every other label path
    // is gated on, carried here so a delineator label is asserted
    // against the same DefaultText typography the live factory resolves.
    private static readonly SceneStyle DefaultLabelStyle = new(
        "DefaultText",
        fontFamily: "Aptos",
        fontSizePt: 8.0);

    private static DelineatorRequest Req(
        GanttEvent e,
        ITextMetrics? metrics = null,
        GanttLabelPosition position = GanttLabelPosition.Auto,
        double? width = null,
        ColourHex? textColour = null,
        SceneStyle? labelStyle = null) =>
        new(
            e,
            new SceneStyle(
                "DefaultDelineator",
                strokeColour: ColourHex.Parse("#404040"),
                textColour: textColour),
            width ?? 0.75,
            Plot,
            Chart,
            2.0,
            metrics ?? Metrics,
            position,
            LabelStyle: labelStyle ?? DefaultLabelStyle);

    /// <summary>
    /// R4.11: replacing the delineator's line width must not discard the named
    /// style's text colour.
    /// </summary>
    /// <remarks>
    /// §24 fixes the line width at the <c>DelineatorLinePt</c> token, so
    /// <c>WithWidth</c> rebuilds the style rather than copying it. That rebuild
    /// takes <see cref="SceneStyle"/>'s members positionally, and
    /// <c>TextColour</c> is its tenth optional parameter — so an earlier version
    /// dropped it silently. The golden snapshot could not catch it: the reference
    /// fixture's delineator style declares no text colour, so there was nothing to
    /// lose. This test supplies one.
    /// </remarks>
    [Fact]
    public void Applying_the_delineator_line_width_preserves_the_named_styles_text_colour()
    {
        ColourHex white = ColourHex.Parse("#FFFFFF");

        DelineatorResult result = Build(Req(Event(), textColour: white));

        Assert.NotNull(result.Primitive);
        Assert.Equal(white, result.Primitive!.Style.TextColour);
    }

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
    public void A_label_no_corner_can_hold_is_suppressed_with_a_warning_per_label()
    {
        // Positive test for the suppression path: four blockers leave no
        // corner, and the builder warns once per label rather than
        // drawing over an obstruction. The row's date label is
        // suppressed with the description label: its corner box is
        // wider than the description's, so the same blockers hold
        // neither.
        List<RectD> blockers = [.. Corners.Select(c => Build(Event(), position: c).LabelBounds!.Value)];
        DelineatorResult r = Build(Req(Event()), blockers);

        Assert.NotNull(r.Primitive);
        Assert.Null(r.Label);
        Assert.Null(r.DateLabel);
        Assert.Equal(2, r.Warnings.Count);
        Assert.Equal(
            new[] { DelineatorBuilder.LabelSuppressedCode, DelineatorBuilder.DateLabelSuppressedCode },
            r.Warnings.Select(warning => warning.Code));
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
    public void The_label_carries_the_predetermined_label_style_not_the_line_style()
    {
        // The gap this pins: the label previously carried the LINE
        // style, so it rendered with a stroke-only style that names no
        // font - the host drew its own face and size inside a box the
        // 8pt seam had measured. The line keeps the line style; the
        // label takes the resolved DefaultText label style, exactly as
        // a span's, a milestone's, and a splitter's description label
        // does.
        DelineatorResult r = Build(Req(Event()));

        Assert.Equal("DefaultDelineator", r.Primitive!.Style.StyleKey);
        Assert.Equal("DefaultText", r.Label!.Style.StyleKey);
        Assert.Equal("Aptos", r.Label.Style.FontFamily);
        Assert.Equal(8.0, r.Label.Style.FontSizePt);
    }

    // ------------------------------------------------------------------
    // The positionable date label (owner ruling 2026-10-14)
    // ------------------------------------------------------------------

    [Fact]
    public void The_date_label_renders_the_invariant_start_date_at_the_first_free_corner()
    {
        // The date label the positionable vocabulary renders: the
        // row's Start in the invariant dd/mm/yyyy format (ADR-0016),
        // at the first corner the description label does not already
        // hold. The description label holds the Auto order's first
        // corner (TopRight), so the date label takes the second
        // (TopLeft), right-aligned so it ends against the line — the
        // §24 alignment rule the description label follows.
        DelineatorResult r = Build(Event());

        Assert.Equal("12/09/2026", r.DateLabel!.Text);
        Assert.EndsWith($":{DateLabelBuilder.StartRole}", r.DateLabel.PrimitiveId, StringComparison.Ordinal);
        Assert.Equal(98.0, r.DateLabel.TextBounds.Left);
        Assert.Equal(Plot.Top, r.DateLabel.TextBounds.Top);
        Assert.Equal(40.0, r.DateLabel.TextBounds.Width);
        Assert.Equal(GanttTextAlignment.Right, r.DateLabel.Alignment);
        Assert.Equal(ZLayer.DelineatorLabel, r.DateLabel.ZLayer);
        Assert.Equal("DefaultText", r.DateLabel.Style.StyleKey);
    }

    [Fact]
    public void The_date_label_avoids_the_rows_description_label()
    {
        // §23's rule, applied to the delineator: the date label is
        // planned against the description label just placed for the
        // row, so the two labels of one row never overlap — they sit
        // at opposite corners of the same stack slot.
        DelineatorResult r = Build(Event());

        Assert.NotNull(r.Label);
        Assert.NotNull(r.DateLabel);
        Assert.False(r.Label.TextBounds.IntersectsWith(r.DateLabel.TextBounds));
        Assert.Equal(GanttLabelPosition.TopRight, r.LabelPosition);
        Assert.True(r.DateLabel.TextBounds.Right < r.Primitive!.From.X);
    }

    [Fact]
    public void An_explicit_date_label_corner_is_honoured_exactly()
    {
        // The explicit-position discipline: a requested corner is
        // rendered exactly, never re-placed by the Auto order.
        // BottomRight puts the box below the line's end, reading
        // away from the line (left-aligned).
        GanttEvent e = Event() with { DateLabelPlacement = GanttLabelPosition.BottomRight };
        DelineatorResult r = Build(Req(e));

        Assert.Equal("12/09/2026", r.DateLabel!.Text);
        Assert.Equal(Left + (2 * Day) + 2.0, r.DateLabel.TextBounds.Left);
        Assert.Equal(Plot.Bottom - 10.0, r.DateLabel.TextBounds.Top);
        Assert.Equal(GanttTextAlignment.Left, r.DateLabel.Alignment);
        Assert.Empty(r.Warnings);
    }

    [Fact]
    public void A_none_date_label_placement_emits_no_date_label_and_no_warning()
    {
        // None is the explicit no-date-label choice in the
        // positionable vocabulary, so the gate skips the date label
        // entirely — no box, no warning.
        GanttEvent e = Event() with { DateLabelPlacement = GanttLabelPosition.None };
        DelineatorResult r = Build(Req(e));

        Assert.NotNull(r.Primitive);
        Assert.Null(r.DateLabel);
        Assert.Null(r.DateLabelBounds);
        Assert.Empty(r.Warnings);
    }

    [Fact]
    public void An_explicit_date_label_corner_that_fits_nowhere_is_suppressed_with_a_warning()
    {
        // Positive test for the explicit-position discipline's
        // failure path: a requested corner that fits nowhere is
        // suppressed and reported, never silently moved. The blocker
        // is the bottom-right box a ten-character date occupies.
        RectD blocker = new(Left + (2 * Day) + 2.0, Plot.Bottom - 10.0, 40.0, 10.0);
        GanttEvent e = Event() with { DateLabelPlacement = GanttLabelPosition.BottomRight };
        DelineatorResult r = Build(Req(e), [blocker]);

        Assert.NotNull(r.Primitive);
        Assert.Null(r.DateLabel);
        Assert.Equal(DelineatorBuilder.DateLabelSuppressedCode, Assert.Single(r.Warnings).Code);
    }

    [Fact]
    public void A_hidden_date_label_toggle_emits_no_date_label()
    {
        // The date-labels toggle gates every date-label path, the
        // delineator's included.
        DelineatorResult r = Build(Req(Event()) with { ShowDateLabels = false });

        Assert.NotNull(r.Primitive);
        Assert.Null(r.DateLabel);
        Assert.Empty(r.Warnings);
    }

    [Fact]
    public void A_request_without_a_label_style_emits_the_line_but_no_label()
    {
        // Positive test for the gate: a null LabelStyle is the same
        // no-label decision BuildLabels and SplitterBuilder make, so
        // the line still draws and nothing warns - a label that cannot
        // name its own typography is not emitted with a host default.
        DelineatorResult r = Build(Req(Event()) with { LabelStyle = null });

        Assert.NotNull(r.Primitive);
        Assert.Null(r.Label);
        Assert.Null(r.LabelBounds);
        Assert.Null(r.LabelPosition);
        Assert.Empty(r.Warnings);
    }

    [Fact]
    public void A_label_wider_than_its_corner_is_clipped_with_an_ellipsis_to_fit()
    {
        // The clipping flow every other label follows (§22's overflow
        // rule): the right corner's available width is the chart's
        // right edge minus the line and the gap - 400 - 140 - 2 =
        // 258pt - and a 100-character description measures 400pt at
        // the fake seam's 4pt advance. The label is ellipsised to the
        // widest prefix that fits (63 characters plus the marker,
        // 256pt) with exactly one LabelTruncatedToFit warning, rather
        // than being suppressed for having no corner that holds it
        // whole.
        DelineatorResult r = Build(Req(Event(description: new string('x', 100))));

        Assert.Equal(GanttLabelPosition.TopRight, r.LabelPosition);
        Assert.EndsWith("…", r.Label!.Text, StringComparison.Ordinal);
        Assert.Equal(256.0, r.Label.TextBounds.Width);
        Assert.Equal(142.0, r.Label.TextBounds.Left, precision: 9);
        Assert.True(r.Label.TextBounds.Right <= Chart.Right);
        SceneWarning warning = Assert.Single(r.Warnings);
        Assert.Equal(LabelPlanner.TruncatedToFitCode, warning.Code);
    }

    [Fact]
    public void A_left_corner_label_is_clipped_to_the_chart_s_left_edge()
    {
        // The left side's available width is the line minus the gap
        // minus the chart's left edge: 140 - 2 - 0 = 138pt, so the
        // same 400pt description clips to 33 characters plus the
        // marker (136pt) and the box still ends against the line,
        // right-aligned, exactly as an unclipped left label would.
        DelineatorResult r = Build(
            Req(Event(description: new string('x', 100)), position: GanttLabelPosition.TopLeft));

        Assert.Equal(GanttLabelPosition.TopLeft, r.LabelPosition);
        Assert.EndsWith("…", r.Label!.Text, StringComparison.Ordinal);
        Assert.Equal(136.0, r.Label.TextBounds.Width);
        Assert.Equal(138.0, r.Label.TextBounds.Right, precision: 9);
        Assert.Equal(LabelPlanner.TruncatedToFitCode, Assert.Single(r.Warnings).Code);
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

        // Each row contributes two labels — its description and its
        // date label — which share the row's stack slot at opposite
        // corners of the line (the date label avoids the description
        // label's box). The primitives are emitted description-first,
        // date-second, row by row.
        Assert.Equal(4, labels.Count);

        // §24: distinct labels are "stacked deterministically with StackGapPt", so
        // the gap is the space *between* two labels, not the step from the plot
        // edge to each one. The second row must therefore clear the first row's
        // whole height and then the gap. Advancing by one gap per label put the
        // second label inside the first whenever the gap was smaller than the
        // label height, and it only avoided an intersection by being pushed onto
        // the opposite side of the line -- a different corner, not a stack.
        Assert.Equal(labels[0].TextBounds.Top, labels[1].TextBounds.Top);
        Assert.Equal(labels[2].TextBounds.Top, labels[3].TextBounds.Top);
        Assert.Equal(labels[0].TextBounds.Bottom + 6.0, labels[2].TextBounds.Top);
        Assert.Equal(labels[1].TextBounds.Bottom + 6.0, labels[3].TextBounds.Top);
        Assert.All(labels, label => Assert.Equal(10.0, label.TextBounds.Height));

        // The stack is a stack: no two boxes of the group intersect,
        // however small the gap is relative to the label height.
        for (var i = 1; i < labels.Count; i++)
        {
            for (var j = 0; j < i; j++)
            {
                Assert.False(labels[i].TextBounds.IntersectsWith(labels[j].TextBounds));
            }
        }
    }

    [Fact]
    public void Stacked_labels_do_not_overlap_when_the_gap_is_smaller_than_the_label()
    {
        // The discriminating case for the cumulative offset. FakeTextMetrics
        // measures every label 10pt tall, and the gap here is 2pt, so the old
        // "one gap per label" step put the second label's top 2pt below the
        // first's -- eight points inside it. The only reason the overlap never
        // showed was that the colliding corner was rejected and the label moved
        // to the opposite side of the line, which is a flip, not a stack.
        const double gap = 2.0;
        DelineatorGroupResult r = DelineatorLayout.TryBuildGroup(
            new DelineatorGroupRequest(
                [
                    Req(Event(description: "Alpha", sortOrder: 1)),
                    Req(Event(description: "Bravo", sortOrder: 2)),
                    Req(Event(description: "Charlie", sortOrder: 3)),
                ],
                gap),
            Scale()).Result!;

        List<SceneText> labels = [.. r.Primitives.OfType<SceneText>()];

        // Three rows, each contributing its description and its date
        // label (the date label takes the corner the description
        // label does not), so six labels in three stacked pairs.
        Assert.Equal(6, labels.Count);
        Assert.All(labels, label => Assert.Equal(10.0, label.TextBounds.Height));

        // Each row clears the previous row's labels by exactly the gap, so no
        // two boxes can intersect however small the gap is relative to the
        // label height.
        for (var i = 1; i < labels.Count; i++)
        {
            for (var j = 0; j < i; j++)
            {
                Assert.False(labels[i].TextBounds.IntersectsWith(labels[j].TextBounds));
            }
        }

        // The pairs stack: the row's two labels share a top, and the
        // next pair starts one label height plus the gap below.
        Assert.Equal(labels[0].TextBounds.Top, labels[1].TextBounds.Top);
        Assert.Equal(labels[0].TextBounds.Bottom + gap, labels[2].TextBounds.Top);
        Assert.Equal(labels[2].TextBounds.Top, labels[3].TextBounds.Top);
        Assert.Equal(labels[2].TextBounds.Bottom + gap, labels[4].TextBounds.Top);
        Assert.Equal(labels[4].TextBounds.Top, labels[5].TextBounds.Top);
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

    [Fact]
    public void A_group_whose_members_disagree_on_line_style_is_refused()
    {
        // Positive test for the InconsistentLineStyle guard: §24 draws the line
        // once per resolved line style, so a group with two styles has no single
        // line to emit. Both halves are covered, because either alone would still
        // restyle the mixed member: a different stroke colour, and the same colour
        // at a different width.
        GanttEvent a = Event(description: "Alpha");
        GanttEvent b = Event(description: "Bravo");
        DelineatorRequest odd = Req(b) with
        {
            LineStyle = new SceneStyle("DefaultDelineator", strokeColour: ColourHex.Parse("#802020")),
        };

        Assert.Equal(
            DelineatorGroupRefusal.InconsistentLineStyle,
            DelineatorLayout.TryBuildGroup(new DelineatorGroupRequest([Req(a), odd], 2.0), Scale()).Refusal);
        Assert.Equal(
            DelineatorGroupRefusal.InconsistentLineStyle,
            DelineatorLayout
                .TryBuildGroup(new DelineatorGroupRequest([Req(a), Req(b) with { LineWidthPt = 1.5 }], 2.0), Scale())
                .Refusal);
    }

    [Fact]
    public void A_group_of_equal_but_separately_resolved_line_styles_still_builds()
    {
        // The counterweight to the guard above, and the reason it compares by
        // value: Req builds a fresh SceneStyle per call, so every existing group
        // test already exercises equal-but-distinct instances. Comparing by
        // reference would refuse all of them.
        DelineatorGroupResult r = Group(Event(description: "Alpha"), Event(description: "Bravo"));

        Assert.Single(r.Primitives.OfType<SceneLine>());
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
    public void A_group_refuses_when_a_member_would_be_silently_dropped()
    {
        // A member whose own TryBuild refuses must refuse the group, not vanish.
        // The old code ran the member's TryBuild and discarded a refusal without
        // a warning, so the contributing row disappeared from the scene with no
        // diagnostic at all.
        GanttEvent notADelineator = Event() with { Type = GanttEntityType.AsPlannedActivity };

        Assert.Equal(
            DelineatorGroupRefusal.MemberRefused,
            DelineatorLayout.TryBuildGroup(
                new DelineatorGroupRequest([Req(Event()), Req(notADelineator)], 2.0),
                Scale()).Refusal);
    }

    [Fact]
    public void A_group_collects_the_label_warning_of_every_member()
    {
        // Warnings come from every member's outcome, not only the one that built
        // the line. A member whose label no corner can hold emits a suppression
        // warning that the old code dropped with the discarded label outcome.
        // Chart bounds far smaller than the plot are shared by both members, so
        // the group passes its consistency check while no corner can hold a box.
        var narrow = new RectD(0.0, 0.0, 10.0, 10.0);
        DelineatorGroupResult r = DelineatorLayout.TryBuildGroup(
            new DelineatorGroupRequest(
                [
                    Req(Event(description: "Alpha")) with { ChartBounds = narrow },
                    Req(Event(description: "Bravo")) with { ChartBounds = narrow },
                ],
                2.0),
            Scale()).Result!;

        Assert.Equal(2, r.Warnings.Count(warning => warning.Code == DelineatorBuilder.LabelSuppressedCode));
        Assert.Empty(r.Primitives.OfType<SceneText>());
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
