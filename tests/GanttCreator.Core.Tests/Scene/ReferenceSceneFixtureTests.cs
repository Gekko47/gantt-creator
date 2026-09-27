using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// Tests for the R3.12 canonical neutral fixture. The fixture is the single
/// representative input shared by the Core scene tests and the Office
/// integration tests, so its contents and its determinism are pinned here.
/// </summary>
public sealed class ReferenceSceneFixtureTests
{
    /// <summary>The lane holding the planned, actual, and baseline overlap trio.</summary>
    private const string FirstActivityLaneId = "G-00000000000000000000000000000001";
    [Fact]
    public void The_committed_fixture_file_is_present_and_parses()
    {
        Assert.NotEmpty(ReferenceSceneFixture.LoadRows());
    }

    [Fact]
    public void Every_fixture_row_validates_without_a_blocking_error()
    {
        GanttValidationOutcome outcome = ReferenceSceneFixture.LoadValidated();

        Assert.DoesNotContain(outcome.Issues, issue => issue.Severity == GanttValidationSeverity.Error);
        Assert.NotEmpty(outcome.Events);
    }

    [Fact]
    public void The_fixture_carries_the_first_class_domain_cases_together()
    {
        // The point of the fixture is that these cases coexist, so a change that
        // breaks one only in combination with the others is caught here.
        GanttValidationOutcome outcome = ReferenceSceneFixture.LoadValidated();
        HashSet<GanttEntityType> types = [.. outcome.Events.Select(@event => @event.Type)];

        Assert.Contains(GanttEntityType.AsPlannedActivity, types);
        Assert.Contains(GanttEntityType.AsBuiltActivity, types);
        Assert.Contains(GanttEntityType.BaselineActivity, types);
        Assert.Contains(GanttEntityType.CriticalInterval, types);
        Assert.Contains(GanttEntityType.DelayEvent, types);
        Assert.Contains(GanttEntityType.AsPlannedProcurement, types);
        Assert.Contains(GanttEntityType.AsBuiltProcurement, types);
        Assert.Contains(GanttEntityType.AsPlannedMilestone, types);
        Assert.Contains(GanttEntityType.AsBuiltMilestone, types);
        Assert.Contains(GanttEntityType.CriticalMilestone, types);
        Assert.Contains(GanttEntityType.Delineator, types);
    }

    [Fact]
    public void Three_activity_types_share_one_lane()
    {
        GanttValidationOutcome outcome = ReferenceSceneFixture.LoadValidated();

        // Lane L-01 carries the planned/actual/baseline trio, so the three overlap
        // cases are exercised in combination rather than in isolation. The same lane
        // also holds two critical intervals parented to the planned span.
        GanttEvent[] lane =
        [
            .. outcome.Events.Where(@event => @event.LaneId?.Value == FirstActivityLaneId),
        ];
        Assert.Equal(
            [GanttEntityType.AsPlannedActivity, GanttEntityType.AsBuiltActivity, GanttEntityType.BaselineActivity],
            lane.Where(@event => @event.Type is not GanttEntityType.CriticalInterval).Select(@event => @event.Type));
        Assert.Equal(2, lane.Count(@event => @event.Type == GanttEntityType.CriticalInterval));
    }

    [Fact]
    public void Every_critical_interval_names_a_valid_mapped_parent()
    {
        GanttValidationOutcome outcome = ReferenceSceneFixture.LoadValidated();
        HashSet<string> ids = [.. outcome.Events.Select(@event => @event.Id.Value)];

        GanttEvent[] critical = [.. outcome.Events.Where(@event => @event.Type == GanttEntityType.CriticalInterval)];
        Assert.Equal(2, critical.Length);
        foreach (GanttEvent @event in critical)
        {
            Assert.NotNull(@event.ParentId);
            Assert.Contains(@event.ParentId!.Value, ids);
        }
    }

    [Fact]
    public void Two_delineators_share_one_date_and_a_third_is_later()
    {
        GanttValidationOutcome outcome = ReferenceSceneFixture.LoadValidated();
        GanttEvent[] delineators = [.. outcome.Events.Where(@event => @event.Type == GanttEntityType.Delineator)];

        Assert.Equal(3, delineators.Length);
        Assert.Equal(2, delineators.Count(@event => @event.Start == new DateOnly(2026, 2, 2)));
    }

    [Fact]
    public void Two_milestones_share_one_date()
    {
        GanttValidationOutcome outcome = ReferenceSceneFixture.LoadValidated();
        GanttEvent[] milestones =
        [
            .. outcome.Events.Where(@event => @event.Type is GanttEntityType.AsPlannedMilestone
                or GanttEntityType.AsBuiltMilestone
                or GanttEntityType.CriticalMilestone),
        ];

        Assert.True(
            milestones.Length > 1
                && milestones.Select(@event => @event.Start).Distinct().Count() < milestones.Length,
            "The fixture must place more than one milestone on the same date.");
    }

    [Fact]
    public void Milestones_and_delineators_read_only_the_start_date()
    {
        GanttValidationOutcome outcome = ReferenceSceneFixture.LoadValidated();

        Assert.All(
            outcome.Events.Where(@event => @event.Type is GanttEntityType.AsPlannedMilestone
                or GanttEntityType.AsBuiltMilestone
                or GanttEntityType.CriticalMilestone
                or GanttEntityType.Delineator),
            @event => Assert.Null(@event.Finish));
    }

    [Fact]
    public void Loading_the_fixture_twice_produces_identical_rows()
    {
        IReadOnlyList<GanttRowDto> first = ReferenceSceneFixture.LoadRows();
        IReadOnlyList<GanttRowDto> second = ReferenceSceneFixture.LoadRows();

        Assert.Equal(
            first.Select(row => (row.RowNumber, row.IdCell.Value, row.TypeCell.Value, row.StartCell.Value, row.FinishCell.Value)),
            second.Select(row => (row.RowNumber, row.IdCell.Value, row.TypeCell.Value, row.StartCell.Value, row.FinishCell.Value)));
    }

    [Fact]
    public void Validation_is_deterministic_across_repeated_loads()
    {
        GanttValidationOutcome first = ReferenceSceneFixture.LoadValidated();
        GanttValidationOutcome second = ReferenceSceneFixture.LoadValidated();

        Assert.Equal(
            first.Issues.Select(issue => (issue.RowNumber, issue.Field, issue.Code, issue.Severity)),
            second.Issues.Select(issue => (issue.RowNumber, issue.Field, issue.Code, issue.Severity)));
        Assert.Equal(
            first.Events.Select(@event => @event.Id.Value),
            second.Events.Select(@event => @event.Id.Value));
    }

    [Fact]
    public void No_fixture_row_identifier_is_generated_at_load_time()
    {
        // GanttRowId.New() would mint a different id per call, so a fixture built
        // from generated ids could never produce a byte-stable scene. The ids are
        // fixed in the committed file in the code-owned G- plus 32 lowercase hex
        // form, which is what this pins.
        IReadOnlyList<GanttRowDto> rows = ReferenceSceneFixture.LoadRows();

        Assert.Equal("G-000000000000000000000000000000a1", rows[0].IdCell.Value);
        Assert.All(rows, row => Assert.Matches("^G-[0-9a-f]{32}$", row.IdCell.Value!));
    }
}
