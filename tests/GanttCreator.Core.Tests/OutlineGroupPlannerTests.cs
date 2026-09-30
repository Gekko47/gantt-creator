namespace GanttCreator.Core.Tests;

/// <summary>
/// Tests for <see cref="OutlineGroupPlanner"/> (R4.7D, ADR-0026 D5): outline
/// groups derived from <c>ParentId</c> alone, never read back from Excel.
/// </summary>
public sealed class OutlineGroupPlannerTests
{
    private static GanttRowId NewId() => GanttRowId.New();

    private static GanttEvent Event(
        int rowNumber,
        GanttRowId id,
        GanttRowId? parentId = null,
        GanttEntityType type = GanttEntityType.AsPlannedActivity) =>
        new(
            rowNumber,
            id,
            NewId(),
            0,
            type,
            "Build frame",
            new DateOnly(2026, 1, 1),
            new DateOnly(2026, 12, 31),
            parentId,
            "AsPlannedActivity",
            null,
            null,
            null,
            true,
            null);

    /// <summary>
    /// A parent with two contiguous children is one group, and the group is
    /// reconstructed purely from the ParentId links.
    /// </summary>
    [Fact]
    public void A_parent_with_contiguous_children_is_one_group()
    {
        GanttRowId parent = NewId();
        GanttRowId first = NewId();
        GanttRowId second = NewId();

        OutlineGroupPlan plan = OutlineGroupPlanner.Plan(
            [Event(2, parent), Event(3, first, parent), Event(4, second, parent)]);

        OutlineGroup group = Assert.Single(plan.Groups);
        Assert.Equal(2, group.ParentRowNumber);
        Assert.Equal([3, 4], group.ChildRowNumbers);
        Assert.Empty(plan.RowsToUngroup);
    }

    /// <summary>
    /// Two sibling parents each get their own group, in worksheet order. Excel
    /// groups are per-parent, so collapsing one must not affect the other.
    /// </summary>
    [Fact]
    public void Two_parents_get_two_groups_in_row_order()
    {
        GanttRowId firstParent = NewId();
        GanttRowId secondParent = NewId();

        OutlineGroupPlan plan = OutlineGroupPlanner.Plan(
            [
                Event(2, firstParent),
                Event(3, NewId(), firstParent),
                Event(4, secondParent),
                Event(5, NewId(), secondParent),
            ]);

        Assert.Equal([2, 4], plan.Groups.Select(static group => group.ParentRowNumber));
        Assert.Equal(2, plan.GroupedRowCount);
    }

    /// <summary>
    /// Input order does not change the result. A re-sorted worksheet must produce
    /// the same grouping, or a user sort would change the outline.
    /// </summary>
    [Fact]
    public void The_plan_is_independent_of_input_order()
    {
        GanttRowId parent = NewId();
        GanttRowId first = NewId();
        GanttRowId second = NewId();
        List<GanttEvent> forward =
            [Event(2, parent), Event(3, first, parent), Event(4, second, parent)];

        OutlineGroupPlan forwardPlan = OutlineGroupPlanner.Plan(forward);
        OutlineGroupPlan reversedPlan = OutlineGroupPlanner.Plan([.. Enumerable.Reverse(forward)]);

        Assert.Equal(forwardPlan.Groups, reversedPlan.Groups);
    }

    /// <summary>
    /// A child whose parent is not a row of this table cannot be grouped, and is not
    /// reported for ungrouping either: there is no stale group to clear, because the
    /// validator owns that case. The planner is silent rather than noisy.
    /// </summary>
    [Fact]
    public void A_child_of_an_absent_parent_is_not_planned()
    {
        OutlineGroupPlan plan = OutlineGroupPlanner.Plan([Event(2, NewId(), NewId())]);

        Assert.Empty(plan.Groups);
        Assert.Empty(plan.RowsToUngroup);
    }

    /// <summary>
    /// A child ABOVE its parent is reported for ungrouping rather than having its
    /// rows moved. An Excel outline group is a contiguous run beneath its parent, so
    /// this cannot be drawn; silently reordering a user's rows is their decision.
    /// </summary>
    [Fact]
    public void A_child_above_its_parent_is_reported_for_ungrouping()
    {
        GanttRowId parent = NewId();
        GanttRowId child = NewId();

        OutlineGroupPlan plan = OutlineGroupPlanner.Plan([Event(5, child, parent), Event(9, parent)]);

        Assert.Empty(plan.Groups);
        Assert.Equal([5], plan.RowsToUngroup);
    }

    /// <summary>
    /// An unrelated row between two children breaks the contiguity the group needs.
    /// The first run stays a group and the broken child is reported, rather than the
    /// planner moving rows to make the group drawable.
    /// </summary>
    [Fact]
    public void A_non_contiguous_child_is_reported_and_the_run_is_kept()
    {
        GanttRowId parent = NewId();
        GanttRowId first = NewId();
        GanttRowId straggler = NewId();

        OutlineGroupPlan plan = OutlineGroupPlanner.Plan(
            [
                Event(2, parent),
                Event(3, first, parent),
                Event(4, NewId()),
                Event(5, straggler, parent),
            ]);

        OutlineGroup group = Assert.Single(plan.Groups);
        Assert.Equal([3], group.ChildRowNumbers);
        Assert.Equal([5], plan.RowsToUngroup);
    }

    /// <summary>
    /// No row is ever both grouped and ungrouped by one plan. A stray child used to
    /// be reported for ungrouping and simultaneously given a second group of its own,
    /// so the Office adapter would both create and clear a group on the same row and
    /// the resulting outline would depend on write order. This is the invariant that
    /// makes the two lists safe to apply in sequence.
    /// </summary>
    [Fact]
    public void No_row_is_both_grouped_and_listed_for_ungrouping()
    {
        GanttRowId parent = NewId();
        GanttRowId first = NewId();
        GanttRowId straggler = NewId();

        OutlineGroupPlan plan = OutlineGroupPlanner.Plan(
            [
                Event(2, parent),
                Event(3, first, parent),
                Event(4, NewId()),
                Event(5, straggler, parent),
            ]);

        int[] grouped = [.. plan.Groups.SelectMany(static group => group.ChildRowNumbers)];
        Assert.Empty(plan.RowsToUngroup.Intersect(grouped));
    }

    /// <summary>
    /// A Critical Interval is an ordinary child, so it forms a group exactly as any
    /// other child does: one group under its parent, holding the child's row.
    /// </summary>
    [Fact]
    public void A_critical_interval_child_forms_a_group()
    {
        GanttRowId activity = NewId();
        GanttRowId interval = NewId();

        OutlineGroupPlan plan = OutlineGroupPlanner.Plan(
            [
                Event(2, activity),
                Event(3, interval, activity, GanttEntityType.CriticalInterval),
            ]);

        OutlineGroup group = Assert.Single(plan.Groups);
        Assert.Equal([3], group.ChildRowNumbers);
    }

    /// <summary>
    /// A top-level Critical Interval owning a child is also a two-level hierarchy and
    /// groups normally. This is the case the owner's depth rule leaves intact.
    /// </summary>
    [Fact]
    public void A_top_level_critical_interval_owning_a_child_groups_normally()
    {
        GanttRowId interval = NewId();
        GanttRowId child = NewId();

        OutlineGroupPlan plan = OutlineGroupPlanner.Plan(
            [
                Event(2, interval, type: GanttEntityType.CriticalInterval),
                Event(3, child, interval),
            ]);

        OutlineGroup group = Assert.Single(plan.Groups);
        Assert.Equal(2, group.ParentRowNumber);
        Assert.Equal([3], group.ChildRowNumbers);
    }

    /// <summary>
    /// A top-level row that is nobody's parent produces no groups at all. The common
    /// case -- a flat table -- must not draw outline controls.
    /// </summary>
    [Fact]
    public void A_flat_table_plans_no_groups()
    {
        OutlineGroupPlan plan = OutlineGroupPlanner.Plan(
            [Event(2, NewId()), Event(3, NewId()), Event(4, NewId())]);

        Assert.Empty(plan.Groups);
        Assert.Empty(plan.RowsToUngroup);
    }

    /// <summary>
    /// An empty table plans nothing successfully rather than refusing.
    /// </summary>
    [Fact]
    public void An_empty_table_plans_no_groups()
    {
        OutlineGroupPlan plan = OutlineGroupPlanner.Plan([]);

        Assert.Empty(plan.Groups);
        Assert.Equal(0, plan.GroupedRowCount);
    }

    /// <summary>
    /// A null batch throws rather than returning a plan, so a caller bug cannot be
    /// mistaken for "nothing to group".
    /// </summary>
    [Fact]
    public void A_null_batch_throws()
    {
        Assert.Throws<ArgumentNullException>(() => OutlineGroupPlanner.Plan(null!));
    }
}
