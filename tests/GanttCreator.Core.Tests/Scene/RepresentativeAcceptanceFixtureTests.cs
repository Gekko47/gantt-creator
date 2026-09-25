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
