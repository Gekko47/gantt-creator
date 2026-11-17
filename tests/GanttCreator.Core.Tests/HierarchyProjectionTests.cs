namespace GanttCreator.Core.Tests;

/// <summary>
/// The outline is presentation of the <c>ParentId</c> column, so it must be planned from the rows
/// as read rather than from the rows that validate. Measured live 2026-11-09: an inserted activity
/// has no dates, so it was absent from the validated events, the plan held no group for it and the
/// refresh's reset pass ungrouped the row the insert had just grouped.
/// </summary>
public sealed class HierarchyProjectionTests
{
    private static readonly GanttRowId _parent = GanttRowId.Parse("G-00000000000000000000000000000001");
    private static readonly GanttRowId _child = GanttRowId.Parse("G-00000000000000000000000000000002");

    private static GanttRowDto Row(int rowNumber, string? id, string? parentId, string? type = "As-Planned Activity") =>
        new(
            rowNumber,
            id,
            null,
            null,
            type,
            null,
            null,
            null,
            parentId,
            null,
            null,
            null,
            null,
            true,
            null);

    /// <summary>
    /// A row with a parent and NO dates still carries its <c>ParentId</c>, which is the whole point:
    /// this is exactly the newly inserted activity that the validated plan could not see.
    /// </summary>
    [Fact]
    public void A_dateless_child_keeps_its_parent()
    {
        IReadOnlyList<GanttEvent> events = HierarchyProjection.FromRows(
        [
            Row(1, _parent.Value, null),
            Row(2, _child.Value, _parent.Value),
        ]);

        Assert.Equal(2, events.Count);
        Assert.Null(events[0].ParentId);
        Assert.Equal(_parent, events[1].ParentId);
        Assert.Null(events[1].Start);
        Assert.Null(events[1].Finish);
    }

    /// <summary>
    /// The planner turns that projection into a group, which is what the refresh must apply and
    /// must therefore not subsequently reset.
    /// </summary>
    [Fact]
    public void The_plan_groups_a_dateless_child_under_its_parent()
    {
        OutlineGroupPlan plan = OutlineGroupPlanner.Plan(HierarchyProjection.FromRows(
        [
            Row(1, _parent.Value, null),
            Row(2, _child.Value, _parent.Value),
        ]));

        OutlineGroup group = Assert.Single(plan.Groups);
        Assert.Equal(1, group.ParentRowNumber);
        Assert.Equal([2], group.ChildRowNumbers);
        Assert.Empty(plan.RowsToUngroup);
    }

    /// <summary>
    /// A row whose <c>Id</c> cannot be parsed is skipped rather than becoming an unaddressable
    /// member, and row numbers are the read row's own so the group spans real rows.
    /// </summary>
    [Fact]
    public void An_unparseable_id_is_skipped_and_row_numbers_are_preserved()
    {
        IReadOnlyList<GanttEvent> events = HierarchyProjection.FromRows(
        [
            Row(4, _parent.Value, null),
            Row(5, "not-an-id", _parent.Value),
            Row(6, _child.Value, _parent.Value),
        ]);

        Assert.Equal(2, events.Count);
        Assert.Equal(4, events[0].RowNumber);
        Assert.Equal(6, events[1].RowNumber);
    }

    /// <summary>
    /// A malformed <c>ParentId</c> is treated as top-level, matching the shared projection the
    /// insert path and the mutators use, so a typo cannot invent a parent.
    /// </summary>
    [Fact]
    public void A_malformed_parent_is_treated_as_top_level()
    {
        IReadOnlyList<GanttEvent> events = HierarchyProjection.FromRows(
        [
            Row(1, _parent.Value, null),
            Row(2, _child.Value, "rubbish"),
        ]);

        Assert.Null(events[1].ParentId);
    }
}
