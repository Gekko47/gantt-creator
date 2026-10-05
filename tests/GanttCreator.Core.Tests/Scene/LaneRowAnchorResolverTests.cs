using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// ADR-0034 D1: every lane resolves to the measured worksheet row it must coincide
/// with, and every refusal is a typed refusal rather than a silent fall back to
/// stacking.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why this suite exists, stated as an observed measurement.</b> Setting
/// <c>AnchorLanesToRows = false</c> in <c>ExcelSceneBuildRequestFactory</c> — which
/// disables row anchoring for every LIVE refresh, restoring the pre-ADR-0034 stacking
/// that this resolver exists to replace — leaves the whole suite green. Before this
/// file, nothing observed anchoring: the golden snapshot is built on the
/// <c>Raster</c> profile with anchoring off, and no test set the flag or called the
/// resolver. So a total regression of the feature passed every gate. These tests are
/// what make that mutation fail.
/// </para>
/// <para>
/// <b>Why the refusal tests are the point, not the padding.</b> Each refusal is a
/// branch in <see cref="LaneRowAnchorResolver.TryResolve"/> that exists precisely so
/// an unrepresentable layout is reported. A refusal that no test constructs is a
/// refusal nobody has run.
/// </para>
/// </remarks>
public sealed class LaneRowAnchorResolverTests
{
    private static readonly LaneLayoutMetrics _metrics = new(18, 3, 3, 2, 18, 9);

    // ---------------------------------------------------------------- happy path

    /// <summary>
    /// A single lane is anchored to its own measured row's top and height.
    /// </summary>
    [Fact]
    public void A_lane_is_anchored_to_its_own_measured_row()
    {
        GanttEvent first = Event(1, "first");

        LaneRowAnchorResolution resolution = Resolve([Input(first)], [first], Grid([18], originTopPt: 40));

        Assert.True(resolution.Succeeded);
        LaneRowAnchor anchor = Assert.Single(resolution.Anchors);

        // The anchor is relative to the plot's top, which ADR-0030 D1 makes the first
        // body row's measured top, so the offset subtracts the same origin the
        // renderer adds back once.
        Assert.Equal(0, anchor.TopOffsetPt);
        Assert.Equal(18, anchor.HeightPt);
        Assert.Equal(LaneOrdering.LaneKey(Input(first)), anchor.LaneKey);
    }

    /// <summary>
    /// A later lane is anchored to its own row, not stacked at the running sum.
    /// </summary>
    /// <remarks>
    /// Rows of 20 and 30 measured from origin 40 put the second row's absolute top at
    /// 60, so its offset from the plot top is 20. Stacking would also give 18 here if
    /// the lane height were 18, so the row heights are deliberately NOT the lane
    /// height: 20pt and 30pt rows make stacking and anchoring disagree, which is what
    /// makes this test able to fail.
    /// </remarks>
    [Fact]
    public void A_later_lane_takes_its_own_rows_top_and_height()
    {
        GanttEvent first = Event(1, "first");
        GanttEvent second = Event(2, "second");

        LaneRowAnchorResolution resolution = Resolve([Input(first), Input(second)], [first, second], Grid([20, 30], originTopPt: 40));

        Assert.True(resolution.Succeeded);
        Assert.Equal(2, resolution.Anchors.Count);
        Assert.Equal(20, resolution.Anchors[1].TopOffsetPt);
        Assert.Equal(30, resolution.Anchors[1].HeightPt);
    }

    /// <summary>
    /// A mixed-height body anchors every lane to its own row, exactly.
    /// </summary>
    /// <remarks>
    /// R4.7I's stated acceptance test, and the one the whole feature exists for. The
    /// heights are deliberately irregular so that no lane coincides with the stacked
    /// position by accident.
    /// </remarks>
    [Fact]
    public void A_mixed_height_body_anchors_every_lane_to_its_own_row()
    {
        GanttEvent[] events = [Event(1, "a"), Event(2, "b"), Event(3, "c"), Event(4, "d")];
        double[] heights = [24, 12, 31, 17];

        LaneRowAnchorResolution resolution = Resolve([.. events.Select(e => Input(e))], events, Grid(heights, originTopPt: 100));

        Assert.True(resolution.Succeeded);

        double running = 100;
        for (var index = 0; index < events.Length; index++)
        {
            LaneRowAnchor anchor = resolution.Anchors[index];
            Assert.Equal(running - 100, anchor.TopOffsetPt);
            Assert.Equal(heights[index], anchor.HeightPt);
            running += heights[index];
        }
    }

    /// <summary>
    /// A row that owns no lane leaves its own band empty: the lane below it stays on
    /// its own row instead of being pulled up.
    /// </summary>
    /// <remarks>
    /// This is the reported defect ADR-0034 D1 exists for. Rows 1 and 3 own lanes;
    /// row 2 is a projected child and owns none. Row 3's measured top is
    /// 18 + 20 = 38 from the origin, so its lane sits at 38. Under the old stacking
    /// it would have sat at 18 — one row too high — with the gap stranded at the
    /// bottom of the plot.
    /// </remarks>
    [Fact]
    public void A_row_that_owns_no_lane_leaves_its_own_band_empty()
    {
        GanttEvent owner = Event(1, "owner");
        GanttEvent child = Event(2, "child", parentId: owner.Id);
        GanttEvent later = Event(3, "later");

        LaneRowAnchorResolution resolution = Resolve(
            [Input(owner), Input(child, projectedOnto: owner), Input(later)],
            [owner, child, later],
            Grid([18, 20, 18], originTopPt: 0)
        );

        Assert.True(resolution.Succeeded);

        // Two lanes only: the projected child creates none of its own.
        Assert.Equal(2, resolution.Anchors.Count);

        LaneRowAnchor laterAnchor = resolution.Anchors[1];
        Assert.Equal(LaneOrdering.LaneKey(Input(later)), laterAnchor.LaneKey);

        // 18 + 20, not 18. The 20pt row's band is skipped, not consumed.
        Assert.Equal(38, laterAnchor.TopOffsetPt);
        Assert.Equal(18, laterAnchor.HeightPt);
    }

    /// <summary>
    /// A projected child does not anchor to its own row but its lane is anchored once,
    /// on the owner.
    /// </summary>
    [Fact]
    public void A_projected_child_anchors_onto_its_owner_lane_once()
    {
        GanttEvent owner = Event(1, "owner");
        GanttEvent child = Event(2, "child", parentId: owner.Id);
        LaneEventInput projected = Input(child, projectedOnto: owner);

        LaneRowAnchorResolution resolution = Resolve([Input(owner), projected], [owner, child], Grid([18, 20], originTopPt: 0));

        Assert.True(resolution.Succeeded);

        // Both members resolve to the SAME lane key, so one anchor is produced and it
        // is the owner's row — never the child's.
        LaneRowAnchor anchor = Assert.Single(resolution.Anchors);
        Assert.Equal(LaneOrdering.LaneKey(Input(owner)), anchor.LaneKey);
        Assert.Equal(0, anchor.TopOffsetPt);
        Assert.Equal(18, anchor.HeightPt);
    }

    /// <summary>
    /// An anchor is found by its lane key, so it can never be applied to a different
    /// lane than the one it was resolved for.
    /// </summary>
    [Fact]
    public void An_anchor_is_found_only_by_its_own_lane_key()
    {
        GanttEvent first = Event(1, "first");

        LaneRowAnchorResolution resolution = Resolve([Input(first)], [first], Grid([18]));

        Assert.True(resolution.TryGet(LaneOrdering.LaneKey(Input(first)), out LaneRowAnchor? found));
        Assert.NotNull(found);
        Assert.Equal(18, found!.HeightPt);

        Assert.False(resolution.TryGet("lane:nothing-resolved-this", out LaneRowAnchor? missing));
        Assert.Null(missing);
    }

    /// <summary>
    /// A hidden row measures zero, so the rows below it share its top — and a lane on
    /// a later row lands where the worksheet puts it.
    /// </summary>
    /// <remarks>
    /// <para>
    /// A collapsed outline group hides its child rows and Excel reports their height as
    /// zero. Refusing zero as a missing measurement is what previously made
    /// collapse-then-Refresh fail outright, so the grid accepts it and consumes no
    /// space.
    /// </para>
    /// <para>
    /// The hidden row here is PROJECTED onto the first row, so it owns no lane of its
    /// own. That matters: a zero-height row that <em>did</em> own a lane is refused
    /// (<see cref="A_lane_owning_row_measuring_zero_height_is_refused"/>), because a
    /// row that owns a lane is by definition a visible one. Only the lane-free hidden
    /// row is legitimate, and this is the case that must not refuse.
    /// </para>
    /// </remarks>
    [Fact]
    public void A_zero_height_row_consumes_no_space_and_the_lane_below_shares_its_top()
    {
        GanttEvent first = Event(1, "first");
        GanttEvent hidden = Event(2, "hidden", parentId: first.Id);
        GanttEvent later = Event(3, "later");

        PanelCellGrid grid = Grid([18, 0, 18], originTopPt: 0);

        // The grid is where "consumes no space" is decided: a zero row contributes
        // nothing to the running top, so the row below inherits its predecessor's.
        Assert.Equal([0, 18, 18], grid.RowTopsPt);

        LaneRowAnchorResolution resolution = Resolve(
            [Input(first), Input(hidden, projectedOnto: first), Input(later)],
            [first, hidden, later],
            grid
        );

        Assert.True(resolution.Succeeded);
        Assert.Equal(18, resolution.Anchors[1].TopOffsetPt);
    }

    /// <summary>
    /// Input order does not change the resolution: the ordinal comes from row order.
    /// </summary>
    /// <remarks>
    /// The measured heights are positional, so a caller who supplies its rows in a
    /// different order must get the same anchors. An ordinal taken from the input
    /// order rather than from row order would anchor each lane to whichever row
    /// happened to arrive first.
    /// </remarks>
    [Fact]
    public void The_resolution_is_independent_of_the_supplied_input_order()
    {
        GanttEvent[] events = [Event(1, "a"), Event(2, "b"), Event(3, "c")];

        LaneRowAnchorResolution forwards = Resolve([.. events.Select(e => Input(e))], events, Grid([15, 25, 35], originTopPt: 0));

        LaneRowAnchorResolution reversed = Resolve(
            [.. events.Reverse().Select(e => Input(e))],
            [.. events.Reverse()],
            Grid([15, 25, 35], originTopPt: 0)
        );

        Assert.True(forwards.Succeeded);
        Assert.True(reversed.Succeeded);
        Assert.Equal(
            forwards.Anchors.Select(anchor => (anchor.LaneKey, anchor.TopOffsetPt, anchor.HeightPt)),
            reversed.Anchors.Select(anchor => (anchor.LaneKey, anchor.TopOffsetPt, anchor.HeightPt))
        );
    }

    /// <summary>
    /// A row sharing one <c>LaneId</c> with another is one lane, anchored once on the
    /// first row that owns it.
    /// </summary>
    [Fact]
    public void Rows_sharing_a_lane_id_produce_one_anchor_for_that_lane()
    {
        var laneId = GanttRowId.New();
        GanttEvent first = Event(1, "first", laneId: laneId);
        GanttEvent second = Event(2, "second", laneId: laneId);

        LaneRowAnchorResolution resolution = Resolve([Input(first), Input(second)], [first, second], Grid([18, 18], originTopPt: 0));

        Assert.True(resolution.Succeeded);
        LaneRowAnchor anchor = Assert.Single(resolution.Anchors);
        Assert.Equal(laneId.Value, anchor.LaneKey);
        Assert.Equal(0, anchor.TopOffsetPt);
    }

    // ---------------------------------------------------------------- refusals

    /// <summary>A null lane-input collection is refused, not defaulted.</summary>
    [Fact]
    public void A_null_lane_input_collection_is_refused()
    {
        LaneRowAnchorResolution resolution = LaneRowAnchorResolver.TryResolve(null, [Event(1, "a")], Grid([18]));

        Assert.False(resolution.Succeeded);
        Assert.Equal(LaneRowAnchorRefusal.NullInput, resolution.Refusal);
        Assert.Empty(resolution.Anchors);
    }

    /// <summary>A null event collection is refused.</summary>
    [Fact]
    public void A_null_event_collection_is_refused()
    {
        LaneRowAnchorResolution resolution = LaneRowAnchorResolver.TryResolve([Input(Event(1, "a"))], null!, Grid([18]));

        Assert.False(resolution.Succeeded);
        Assert.Equal(LaneRowAnchorRefusal.NullEvents, resolution.Refusal);
        Assert.Empty(resolution.Anchors);
    }

    /// <summary>A null measured grid is refused.</summary>
    [Fact]
    public void A_null_measured_grid_is_refused()
    {
        LaneRowAnchorResolution resolution = LaneRowAnchorResolver.TryResolve([Input(Event(1, "a"))], [Event(1, "a")], null);

        Assert.False(resolution.Succeeded);
        Assert.Equal(LaneRowAnchorRefusal.NullGrid, resolution.Refusal);
        Assert.Empty(resolution.Anchors);
    }

    /// <summary>
    /// A lane whose owning row is not among the supplied events is refused.
    /// </summary>
    /// <remarks>
    /// An invented anchor here would place the lane at a plausible but wrong height,
    /// which is the exact silent misalignment the resolver exists to remove.
    /// </remarks>
    [Fact]
    public void A_lane_whose_owning_row_is_absent_from_the_events_is_refused()
    {
        // The lane input names a row the event set does not contain.
        LaneEventInput orphan = Input(Event(9, "orphan"));

        LaneRowAnchorResolution resolution = LaneRowAnchorResolver.TryResolve([orphan], [Event(1, "present")], Grid([18, 18]));

        Assert.False(resolution.Succeeded);
        Assert.Equal(LaneRowAnchorRefusal.UnknownLaneOwner, resolution.Refusal);
    }

    /// <summary>
    /// An owning row beyond the measured body is refused rather than read off the end.
    /// </summary>
    /// <remarks>
    /// Without this guard the lookup would index past the measured list and the lanes
    /// below would anchor to whatever row came last — a measurement that returned
    /// fewer rows than the events reference, which is a real adapter disagreement.
    /// </remarks>
    [Fact]
    public void A_row_beyond_the_measured_body_is_refused()
    {
        GanttEvent first = Event(1, "first");
        GanttEvent third = Event(3, "third");

        // Two events, but only ONE measured row: the second row's height is missing.
        LaneRowAnchorResolution resolution = LaneRowAnchorResolver.TryResolve([Input(first), Input(third)], [first, third], Grid([18]));

        Assert.False(resolution.Succeeded);
        Assert.Equal(LaneRowAnchorRefusal.RowOutsideMeasuredRows, resolution.Refusal);
    }

    /// <summary>
    /// A lane-owning row measuring zero height is refused.
    /// </summary>
    /// <remarks>
    /// A hidden row measures zero and is legitimately consumed by nothing, but a row
    /// that OWNS a lane is a visible row, so a zero here is a real disagreement
    /// between the hierarchy and the sheet rather than a row to skip.
    /// </remarks>
    [Fact]
    public void A_lane_owning_row_measuring_zero_height_is_refused()
    {
        GanttEvent first = Event(1, "first");

        LaneRowAnchorResolution resolution = LaneRowAnchorResolver.TryResolve([Input(first)], [first], Grid([0]));

        Assert.False(resolution.Succeeded);
        Assert.Equal(LaneRowAnchorRefusal.DegenerateRowHeight, resolution.Refusal);
    }

    /// <summary>
    /// A zero-height row that owns NO lane is not this resolver's problem: the lane
    /// below it still anchors.
    /// </summary>
    /// <remarks>
    /// The counterweight to the refusal above. Without it, a collapsed outline group
    /// would refuse the whole refresh, which is the failure ADR-0034 relaxed
    /// <c>PanelCellGrid</c>'s zero-height rule to avoid.
    /// </remarks>
    [Fact]
    public void A_zero_height_row_that_owns_no_lane_does_not_refuse()
    {
        GanttEvent first = Event(1, "first");
        GanttEvent collapsed = Event(2, "collapsed");
        GanttEvent later = Event(3, "later");

        // The middle row is projected, so it owns no lane and its zero height is not
        // consulted for any anchor.
        LaneRowAnchorResolution resolution = Resolve(
            [Input(first), Input(collapsed, projectedOnto: first), Input(later)],
            [first, collapsed, later],
            Grid([18, 0, 18], originTopPt: 0)
        );

        Assert.True(resolution.Succeeded);
        Assert.Equal(18, resolution.Anchors[1].TopOffsetPt);
    }

    // ------------------------------------------- the contract the resolver serves

    /// <summary>
    /// The anchored lanes reach <see cref="LaneLayoutBuilder"/> and are placed on
    /// their rows, where stacking would have put them elsewhere.
    /// </summary>
    /// <remarks>
    /// The resolver alone does not move a bar; this is the link that makes the
    /// resolver's output reach the geometry. Without it a correct resolver would be
    /// called and ignored, which is precisely the shape of the gap this suite closes.
    /// </remarks>
    [Fact]
    public void The_resolved_anchors_place_the_lanes_on_their_measured_rows()
    {
        GanttEvent first = Event(1, "first");
        GanttEvent child = Event(2, "child", parentId: first.Id);
        GanttEvent later = Event(3, "later");

        IReadOnlyList<LaneEventInput> inputs = [Input(first), Input(child, projectedOnto: first), Input(later)];

        LaneRowAnchorResolution anchors = Resolve(inputs, [first, child, later], Grid([18, 20, 18]));
        Assert.True(anchors.Succeeded);

        LaneLayoutCreationOutcome anchored = LaneLayoutBuilder.TryBuild(inputs, _metrics, anchors);
        LaneLayoutCreationOutcome stacked = LaneLayoutBuilder.TryBuild(inputs, _metrics);

        Assert.True(anchored.Succeeded);
        Assert.True(stacked.Succeeded);

        // The later lane is the second in both layouts, so the comparison is direct.
        Assert.Equal(38, anchored.Layout!.Lanes[1].Top);
        Assert.Equal(18, stacked.Layout!.Lanes[1].Top);
        Assert.Equal(18, anchored.Layout.Lanes[1].Height);
    }

    /// <summary>
    /// A refused resolution stops the layout rather than falling back to stacking.
    /// </summary>
    [Fact]
    public void A_refused_resolution_stops_the_lane_layout()
    {
        var refused = new LaneRowAnchorResolution([], LaneRowAnchorRefusal.UnknownLaneOwner);

        LaneLayoutCreationOutcome outcome = LaneLayoutBuilder.TryBuild([Input(Event(1, "only"))], _metrics, refused);

        Assert.False(outcome.Succeeded);
        Assert.Equal(LaneLayoutRefusal.InvalidRowAnchors, outcome.Refusal);
        Assert.Null(outcome.Layout);
    }

    // ------------------------------------------------------------------- helpers

    private static LaneRowAnchorResolution Resolve(
        IReadOnlyList<LaneEventInput> inputs,
        IReadOnlyList<GanttEvent> events,
        PanelCellGrid grid
    ) => LaneRowAnchorResolver.TryResolve(inputs, events, grid);

    /// <summary>
    /// A measured grid over the given body-row heights.
    /// </summary>
    /// <remarks>
    /// The five columns are the ones the panel primitives are emitted from; they are
    /// named rather than generated because <see cref="PanelCellGrid"/> refuses a blank
    /// or duplicated schema name, so a synthesised list would fail validation for a
    /// reason unrelated to anchoring.
    /// </remarks>
    private static PanelCellGrid Grid(IReadOnlyList<double> rowHeightsPt, double originTopPt = 0) =>
        PanelCellGrid
            .TryCreate(
                [
                    new PanelColumn("Id", 80),
                    new PanelColumn("Type", 120),
                    new PanelColumn("Description", 180),
                    new PanelColumn("Start", 70),
                    new PanelColumn("Finish", 70),
                ],
                rowHeightsPt,
                headerHeightPt: 10,
                requiredColumns: ["Id", "Type", "Description", "Start", "Finish"],
                originTopPt: originTopPt
            )
            .Grid
        ?? throw new InvalidOperationException("The test grid failed to validate.");

    private static LaneEventInput Input(GanttEvent atEvent, GanttEvent? projectedOnto = null) =>
        new(atEvent, 8, EffectiveStackIndex: null, RenderLaneOwner: projectedOnto);

    private static GanttEvent Event(int row, string suffix, GanttRowId? laneId = null, GanttRowId? parentId = null) =>
        new(
            row,
            GanttRowId.New(),
            laneId,
            StackIndex: null,
            Type: GanttEntityType.AsPlannedActivity,
            Description: suffix,
            Start: new DateOnly(2026, 1, 5),
            Finish: new DateOnly(2026, 1, 16),
            ParentId: parentId,
            StyleKey: null,
            LabelPosition: null,
            FillColour: null,
            StrokeColour: null,
            Visible: true,
            SortOrder: null
        );
}
