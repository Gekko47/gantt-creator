using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// The audit's representative acceptance fixture: one workbook exercising the
/// first-class domain cases together rather than in isolation — multiple events
/// on one lane, overlapping planned/actual spans, a critical interval with valid
/// parentage, milestones, a delineator, a per-row colour override, and a hidden
/// row.
/// </summary>
/// <remarks>
/// <para>
/// The fixture pins observable behaviour, not a golden image: which rows validate,
/// which become events, which carry resolved values, and that the resulting scene
/// is canonical. The row set is deliberately a realistic mix, so a change that
/// breaks one case in combination with the others is caught here even when the
/// per-feature tests still pass.
/// </para>
/// <para>
/// Per AGENTS.md, each scenario asserts the exact validation codes so a silent
/// change in a validator's outcome fails this fixture rather than hiding behind a
/// still-valid workbook.
/// </para>
/// </remarks>
public sealed class RepresentativeAcceptanceFixtureTests
{
    private static int s_nextId;

    private static string NewId() => $"G-{Interlocked.Increment(ref s_nextId):x32}";

    private static string NewLane() => NewId();

    private static DateOnly Day(int day) => new(2026, 9, day);

    /// <summary>
    /// A style definition for the vertical scene test. Only the fields the scene
    /// builder reads are supplied; the rest fall to their defaults.
    /// </summary>
    private static GanttStyleDefinition Style(string key, double height) =>
        new(
            key,
            new HashSet<GanttLabelPosition>
            {
                GanttLabelPosition.None,
                GanttLabelPosition.Auto,
                GanttLabelPosition.Left,
                GanttLabelPosition.Right,
                GanttLabelPosition.Inside,
            },
            EntityColourCapability.Fill | EntityColourCapability.Stroke,
            GanttLabelPosition.Inside,
            FillColour: "#92D050",
            StrokeColour: "#404040",
            TextColour: "#000000",
            HatchPattern: GanttHatchPattern.None,
            HatchPitchPt: 0,
            HatchLinePt: 0,
            StandardOutlinePt: 0.5,
            ActivityHeightPt: height,
            MilestoneSizePt: 8);

    /// <summary>
    /// The styles the representative rows actually reference, built with
    /// <c>ActivityHeightPt</c> 8 so the critical overlay's half-height is the
    /// predetermined 4 rather than a derived parent height.
    /// </summary>
    private static readonly GanttStyleRegistry SceneRegistry = new(
    [
        Style("AsPlannedActivity", 8),
        Style("AsBuiltActivity", 8),
        Style("AsPlannedMilestone", 8),
        Style("AsBuiltMilestone", 8),
        Style("CriticalInterval", 8),
        Style("DefaultDelineator", 8),
    ]);

    /// <summary>
    /// One measured panel height per body row, matching the panel's positional
    /// requirement (see <c>SceneBuilderTests</c>). The representative set has nine
    /// body rows (worksheet rows 2-10).
    /// </summary>
    private static PanelCellGrid Grid(int rowCount) =>
        PanelCellGrid.TryCreate(
            [new PanelColumn("Id", 40), new PanelColumn("Description", 160)],
            [.. Enumerable.Repeat(10.0, rowCount)],
            10,
            ["Id", "Description"]).Grid!;

    private static FrameBandsTheme FrameTheme() =>
        new(
            new SceneStyle("Background"),
            new SceneStyle("AlternateBand"),
            new SceneStyle("MinorGrid"),
            new SceneStyle("MajorGrid"),
            new SceneStyle("YearHeader"),
            new SceneStyle("PeriodHeader"),
            new SceneStyle("Title"));

    /// <summary>
    /// Builds a real scene from the representative fixture's parent and its critical
    /// child, driven through the real <c>SceneBuilder</c>. This is the vertical
    /// integration seam the R4.7E work item requires: the same rows the fixture
    /// validates must also project into a scene, so the critical child is proven
    /// present in a BUILT scene rather than only in a hand-assembled primitive list.
    /// </summary>
    /// <remarks>
    /// The pair is selected from the full fixture rather than re-declared, so this
    /// cannot drift from the rows the validation tests assert on. The slice is
    /// deliberate: the fixture's nine rows share one lane and deliberately overlap
    /// at the same stack (that is what the stacking assertion must be able to see
    /// fail), which the lane layout refuses as a single build. The critical-child
    /// question does not depend on those cases, and <c>SceneBuilderTests</c> already
    /// drives the full multi-stack lane layouts through the real builder.
    /// </remarks>
    private static SceneBuildOutcome BuildRepresentativeCriticalScene()
    {
        GanttValidationOutcome validated = GanttRowValidator.Validate(BuildRepresentativeRows());
        GanttEvent child = Assert.Single(
            validated.Events,
            e => e.Type == GanttEntityType.CriticalInterval);
        GanttEvent parent = Assert.Single(validated.Events, e => e.Id == child.ParentId);
        GanttEvent[] pair = [parent, child];
        return SceneBuilder.TryBuild(
            new SceneBuildRequest
            {
                Events = pair,
                Registry = SceneRegistry,
                Grid = Grid(pair.Length),
                PlotBounds = new RectD(200, 60, 300, 140),
                Metrics = new FakeTextMetrics(_ => 8.0, 10.0),
                LaneMetrics = new LaneLayoutMetrics(40, 3, 3, 2, 18, 9),
                FrameTheme = FrameTheme(),
                PlotStart = Day(1),
                PlotFinish = Day(30),
                GridLinePt = 0.5,
                MajorBoundaryPt = 1,
                MilestoneSizePt = 8,
                TitleBandHeightPt = 14,
                YearBandHeightPt = 16,
                PeriodBandHeightPt = 20,
                DelineatorLinePt = 1,
                DelineatorStackGapPt = 10,
                LabelGapPt = 2,
                LabelHeightPt = 8,
                ChartOuterPaddingPt = 0,
                MinimumHeaderLabelWidthPt = 0,
            });
    }

    private static GanttRowDto Row(
        int rowNumber,
        string id,
        string? lane,
        int? stack,
        string type,
        string? description,
        DateOnly? start,
        DateOnly? finish = null,
        string? parentId = null,
        string? styleKey = null,
        string? fillColour = null,
        bool visible = true) =>
        new(
            rowNumber,
            id,
            lane,
            stack,
            type,
            description,
            start,
            finish,
            parentId,
            styleKey,
            null,
            fillColour,
            null,
            visible,
            null);

    /// <summary>
    /// Builds the representative row set: an overlapping planned/actual pair in
    /// one lane, a second stacked event, a critical interval under the first span,
    /// two milestones, a delineator, a colour-overridden row, and a hidden row.
    /// </summary>
    /// <returns>The representative body rows in authoring order.</returns>
    private static List<GanttRowDto> BuildRepresentativeRows()
    {
        var lane = NewLane();
        var parentId = NewId();

        return
        [
            Row(2, parentId, lane, 0, "As-Planned Activity", "Site preparation", Day(1), Day(10), styleKey: "AsPlannedActivity"),
            // Overlaps row 2, so it takes the next stack rather than sharing one:
            // an overlapping pair stacked identically is what the stacking
            // assertion below must be able to see fail.
            Row(3, NewId(), lane, 1, "As-Built Activity", "Site preparation (actual)", Day(4), Day(12), styleKey: "AsBuiltActivity"),
            // A separate span that must not be able to stand in for the pair.
            Row(4, NewId(), lane, 2, "As-Planned Activity", "Foundation pour", Day(8), Day(18), styleKey: "AsPlannedActivity"),
            Row(5, NewId(), lane, 3, "Critical Interval", "Critical handover", Day(6), Day(9), parentId, "CriticalInterval"),
            Row(6, NewId(), lane, 0, "As-Planned Milestone", "Design freeze", Day(1), styleKey: "AsPlannedMilestone"),
            Row(7, NewId(), lane, 0, "As-Built Milestone", "Practical completion", Day(18), styleKey: "AsBuiltMilestone"),
            Row(8, NewId(), null, null, "Delineator", null, Day(13), styleKey: "DefaultDelineator"),
            Row(9, NewId(), lane, 0, "As-Planned Activity", "Snagging", Day(14), Day(16), styleKey: "AsPlannedActivity", fillColour: "#112233"),
            Row(10, NewId(), lane, 0, "As-Planned Activity", "Superseded activity", Day(2), Day(3), styleKey: "AsPlannedActivity", visible: false),
        ];
    }

    [Fact]
    public void The_representative_workbook_validates_with_no_errors()
    {
        List<GanttRowDto> rows = BuildRepresentativeRows();

        GanttValidationOutcome outcome = GanttRowValidator.Validate(rows);

        Assert.True(
            outcome.IsValid,
            "Unexpected errors: " + string.Join(
                " | ",
                outcome.Issues
                    .Where(issue => issue.Severity == GanttValidationSeverity.Error)
                    .Select(issue => $"row {issue.RowNumber} {issue.Field}/{issue.Code}: {issue.Message}")));
    }

    [Fact]
    public void The_representative_workbook_keeps_every_row_as_an_event()
    {
        List<GanttRowDto> rows = BuildRepresentativeRows();

        GanttValidationOutcome outcome = GanttRowValidator.Validate(rows);

        Assert.Equal(rows.Count, outcome.Events.Count);
        Assert.Equal(
            rows.Select(row => row.RowNumber).Order(),
            outcome.Events.Select(@event => @event.RowNumber).Order());
    }

    [Fact]
    public void Overlapping_spans_in_one_lane_share_a_lane_with_distinct_stacks()
    {
        List<GanttRowDto> rows = BuildRepresentativeRows();

        GanttValidationOutcome outcome = GanttRowValidator.Validate(rows);

        List<GanttEvent> laneRows = [.. outcome.Events.Where(@event => @event.LaneId is not null)];
        Assert.True(laneRows.Count > 1, "Expected several lane-bound events.");

        string firstLane = laneRows[0].LaneId!.Value;
        List<GanttEvent> sameLane = [.. laneRows.Where(@event => @event.LaneId!.Value == firstLane)];
        Assert.Equal(laneRows.Count, sameLane.Count);

        // Assert the overlapping pair's own stacks, not merely that the lane
        // contains some stack indices: a separate non-overlapping span would
        // otherwise satisfy a "distinct stacks exist" check while the pair
        // itself stayed stacked identically.
        GanttEvent planned = Assert.Single(sameLane, @event => @event.Description == "Site preparation");
        GanttEvent asBuilt = Assert.Single(
            sameLane,
            @event => @event.Description == "Site preparation (actual)");
        Assert.True(
            planned.Start!.Value <= asBuilt.Start!.Value && asBuilt.Start <= planned.Finish!.Value,
            "Fixture precondition: the planned and as-built spans must overlap.");
        Assert.Equal(0, planned.StackIndex);
        Assert.Equal(1, asBuilt.StackIndex);
        Assert.NotEqual(planned.StackIndex, asBuilt.StackIndex);

        Assert.Contains(sameLane, @event => @event.StackIndex == 0);
        Assert.Contains(sameLane, @event => @event.StackIndex == 1);
    }

    [Fact]
    public void The_critical_interval_retains_its_resolved_parent()
    {
        List<GanttRowDto> rows = BuildRepresentativeRows();

        GanttValidationOutcome outcome = GanttRowValidator.Validate(rows);

        GanttEvent critical = Assert.Single(outcome.Events, @event =>
            @event.Type == GanttEntityType.CriticalInterval);
        Assert.NotNull(critical.ParentId);
        GanttEvent parent = Assert.Single(outcome.Events, @event =>
            @event.Id.Value == critical.ParentId!.Value);
        Assert.Equal(GanttEntityType.AsPlannedActivity, parent.Type);
        Assert.DoesNotContain(
            outcome.Issues,
            issue => issue.Code is GanttValidationCodes.ParentUnknown
                or GanttValidationCodes.ParentInvalid
                or GanttValidationCodes.ParentNotSpan);
    }

    [Fact]
    public void Milestones_read_only_their_start_date()
    {
        // Per the entity contract a milestone has no Finish, so its null Finish
        // must not be reported as a missing required value.
        List<GanttRowDto> rows = BuildRepresentativeRows();

        GanttValidationOutcome outcome = GanttRowValidator.Validate(rows);

        GanttEvent[] milestones =
        [
            .. outcome.Events.Where(@event => @event.Type is
                GanttEntityType.AsPlannedMilestone or GanttEntityType.AsBuiltMilestone),
        ];
        Assert.Equal(2, milestones.Length);
        Assert.All(milestones, milestone =>
        {
            Assert.NotNull(milestone.Start);
            Assert.Null(milestone.Finish);
        });
        Assert.Equal(Day(1), milestones[0].Start);
        Assert.Equal(Day(18), milestones[1].Start);
    }

    [Fact]
    public void The_delineator_ignores_lane_and_stack()
    {
        List<GanttRowDto> rows = BuildRepresentativeRows();

        GanttValidationOutcome outcome = GanttRowValidator.Validate(rows);

        GanttEvent delineator = Assert.Single(
            outcome.Events,
            @event => @event.Type == GanttEntityType.Delineator);
        Assert.Null(delineator.LaneId);
        Assert.Null(delineator.StackIndex);
    }

    [Fact]
    public void The_colour_override_is_carried_onto_the_event()
    {
        List<GanttRowDto> rows = BuildRepresentativeRows();

        GanttValidationOutcome outcome = GanttRowValidator.Validate(rows);

        GanttEvent overridden = Assert.Single(outcome.Events, @event => @event.FillColour == "#112233");
        Assert.Equal("AsPlannedActivity", overridden.StyleKey);
    }

    [Fact]
    public void The_hidden_row_validates_but_is_marked_not_visible()
    {
        List<GanttRowDto> rows = BuildRepresentativeRows();

        GanttValidationOutcome outcome = GanttRowValidator.Validate(rows);

        GanttEvent hidden = Assert.Single(outcome.Events, @event => !@event.Visible);
        Assert.Equal(10, hidden.RowNumber);
        Assert.DoesNotContain(
            outcome.Issues,
            issue => issue.RowNumber == 10 && issue.Severity == GanttValidationSeverity.Error);
    }

    [Fact]
    public void Validation_is_deterministic_across_repeated_runs()
    {
        // The same rows must produce identical issue order every time; the
        // Severity -> Row -> Field -> Code sort is the contract.
        List<GanttRowDto> rows = BuildRepresentativeRows();

        GanttValidationOutcome first = GanttRowValidator.Validate(rows);
        GanttValidationOutcome second = GanttRowValidator.Validate(rows);

        Assert.Equal(
            first.Issues.Select(issue => (issue.RowNumber, issue.Field, issue.Code, issue.Severity)),
            second.Issues.Select(issue => (issue.RowNumber, issue.Field, issue.Code, issue.Severity)));
        Assert.Equal(
            first.Events.Select(@event => @event.Id.Value),
            second.Events.Select(@event => @event.Id.Value));
    }

    /// <summary>
    /// The vertical integration proof the R4.7E work item requires: the fixture's
    /// critical child — which sits INSIDE its parent's span (row 2 is Sep 1-10, the
    /// child is Sep 6-9) — reaches a scene built by the REAL <c>SceneBuilder</c>, is
    /// present as a filled critical overlay, and is centred on its own slot at half
    /// the predetermined height. This is deliberately not satisfied by the
    /// hand-assembled primitive list above, which never consults the overlay builder
    /// and so cannot regress if the overlay stops being emitted at all.
    /// </summary>
    [Fact]
    public void The_critical_child_is_present_in_the_built_scene_as_a_filled_half_height_overlay()
    {
        SceneBuildOutcome outcome = BuildRepresentativeCriticalScene();

        Assert.True(outcome.Succeeded, "Scene build refused: " + outcome.Refusal);

        SceneRect overlay = Assert.Single(
            outcome.Result!.Scene.Primitives.OfType<SceneRect>(),
            rect => rect.ZLayer == ZLayer.CriticalOverlay);

        // ADR-0027 D2: the overlay carries a FILL. An outline-only or null-fill rect
        // is the defect the ADR removes, so it is asserted, not assumed.
        Assert.NotNull(overlay.Style.FillColour);

        // ADR-0027 D3 as amended: half the PREDETERMINED ActivityHeightPt (8 in this
        // fixture), not half the parent's resolved height and not the parent's own
        // bounds. It consumes no lane height, so its height is style-driven.
        Assert.Equal(4, overlay.Bounds.Height, precision: 6);

        // It is drawn, not deleted: the interval's own Sep 6-9 lies inside the
        // parent's Sep 1-10 and inside the Sep 1-30 plot, so a regression that
        // silently dropped the primitive would leave no overlay to assert on.
        Assert.True(overlay.Bounds.Width > 0, "A critical overlay inside its parent must still be drawn.");
    }

    /// <summary>
    /// The critical child consumes no lane height and does not move its parent (ADR-0027
    /// D7 / R4.7B projection). The fixture places the child at stack 3 while its parent
    /// is at stack 0, so they share the parent's LANE but sit in different stack bands —
    /// being in the same lane does not mean being in the same band, and this test must not
    /// assert that it does. The meaningful contract is that adding the child leaves the
    /// parent's geometry untouched.
    /// </summary>
    [Fact]
    public void The_critical_child_consumes_no_lane_height_and_does_not_move_its_parent()
    {
        GanttValidationOutcome validated = GanttRowValidator.Validate(BuildRepresentativeRows());
        GanttEvent child = Assert.Single(validated.Events, e => e.Type == GanttEntityType.CriticalInterval);
        GanttEvent parent = Assert.Single(validated.Events, e => e.Id == child.ParentId);

        SceneBuildOutcome withChild = BuildRepresentativeCriticalScene();
        Assert.True(withChild.Succeeded, "Scene build refused: " + withChild.Refusal);

        // The parent alone, built through the same seam.
        SceneBuildOutcome parentOnly = SceneBuilder.TryBuild(
            new SceneBuildRequest
            {
                Events = [parent],
                Registry = SceneRegistry,
                Grid = Grid(1),
                PlotBounds = new RectD(200, 60, 300, 140),
                Metrics = new FakeTextMetrics(_ => 8.0, 10.0),
                LaneMetrics = new LaneLayoutMetrics(40, 3, 3, 2, 18, 9),
                FrameTheme = FrameTheme(),
                PlotStart = Day(1),
                PlotFinish = Day(30),
                GridLinePt = 0.5,
                MajorBoundaryPt = 1,
                MilestoneSizePt = 8,
                TitleBandHeightPt = 14,
                YearBandHeightPt = 16,
                PeriodBandHeightPt = 20,
                DelineatorLinePt = 1,
                DelineatorStackGapPt = 10,
                LabelGapPt = 2,
                LabelHeightPt = 8,
                ChartOuterPaddingPt = 0,
                MinimumHeaderLabelWidthPt = 0,
            });
        Assert.True(parentOnly.Succeeded, "Parent-only build refused: " + parentOnly.Refusal);

        // The parent's own bar is byte-for-byte identical with and without the child:
        // the overlay is additive and consumes no lane height.
        SceneRect parentWithChild = Assert.Single(
            withChild.Result!.Scene.Primitives.OfType<SceneRect>(),
            rect => rect.ZLayer == ZLayer.ActivityBody);
        SceneRect parentAlone = Assert.Single(
            parentOnly.Result!.Scene.Primitives.OfType<SceneRect>(),
            rect => rect.ZLayer == ZLayer.ActivityBody);

        Assert.Equal(parentAlone.Bounds, parentWithChild.Bounds);

        // And the chart bounds are unchanged too, proving the child added no height.
        Assert.Equal(parentOnly.Result!.Scene.ChartBounds, withChild.Result!.Scene.ChartBounds);

        // The child itself is present as an overlay.
        Assert.Single(
            withChild.Result!.Scene.Primitives.OfType<SceneRect>(),
            rect => rect.ZLayer == ZLayer.CriticalOverlay);
    }

    [Fact]
    public void The_fixture_builds_a_canonical_scene_with_unique_primitive_ids()
    {
        // Proves the acceptance path end to end: representative rows -> events ->
        // scene primitives -> a validated scene, with no duplicate or unresolved
        // group children.
        List<GanttRowDto> rows = BuildRepresentativeRows();
        GanttValidationOutcome validated = GanttRowValidator.Validate(rows);

        List<ScenePrimitive> primitives = [];
        foreach (GanttEvent @event in validated.Events)
        {
            SceneOwnerId owner = SceneOwnerId.ForRow(@event.Id);
            var bounds = new RectD(0, 0, 10, 10);
            primitives.Add(new SceneRect(
                ScenePrimitive.CreateId(owner, "bar"),
                owner,
                ZLayer.ActivityBody,
                bounds,
                new SceneStyle(@event.StyleKey ?? "Default"),
                @event.Type,
                @event.LaneId is null ? null : 0,
                @event.StackIndex,
                @event.SortOrder));
            primitives.Add(new SceneGroup(
                ScenePrimitive.CreateId(owner, "group"),
                owner,
                ZLayer.Label,
                [ScenePrimitive.CreateId(owner, "bar")],
                @event.Type));
        }

        SceneCreationOutcome scene = GanttScene.TryCreate(
            new RectD(0, 0, 500, 300),
            new RectD(40, 20, 460, 260),
            primitives,
            []);

        Assert.True(scene.Succeeded, "Scene refused: " + scene.Refusal);
        Assert.Equal(primitives.Count, scene.Scene!.Primitives.Count);
        Assert.Equal(
            scene.Scene.Primitives.Select(primitive => primitive.PrimitiveId).Distinct().Count(),
            scene.Scene.Primitives.Count);
    }
}
