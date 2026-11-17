namespace GanttCreator.Core.Tests;

/// <summary>
/// Tests for <see cref="OutlineCollapsePlanner"/> (R5.4): the rows a
/// collapse-all or expand-all request writes, derived from the outline the
/// sheet carries.
/// </summary>
public sealed class OutlineCollapsePlannerTests
{
    /// <summary>
    /// A grouped child row is written: a row deeper than the top level belongs
    /// to a group, and a group has no child to hide without it.
    /// </summary>
    [Fact]
    public void A_grouped_child_row_is_written()
    {
        OutlineCollapsePlan plan = OutlineCollapsePlanner.Plan(
            [
                new OutlineRowLevel(1, OutlineCollapsePlanner.TopLevelOutlineLevel),
                new OutlineRowLevel(2, OutlineCollapsePlanner.TopLevelOutlineLevel + 1),
            ],
            OutlineCollapseDirection.Collapse);

        Assert.Equal([2], plan.ChildRowNumbers);
        Assert.Equal(1, plan.RowCount);
        Assert.False(plan.IsEmpty);
    }

    /// <summary>
    /// The parent row is never in the plan: it stays visible and keeps its
    /// native collapse control.
    /// </summary>
    [Fact]
    public void The_parent_row_is_never_written()
    {
        OutlineCollapsePlan plan = OutlineCollapsePlanner.Plan(
            [
                new OutlineRowLevel(1, OutlineCollapsePlanner.TopLevelOutlineLevel),
                new OutlineRowLevel(2, OutlineCollapsePlanner.TopLevelOutlineLevel + 1),
            ],
            OutlineCollapseDirection.Collapse);

        Assert.DoesNotContain(1, plan.ChildRowNumbers);
        Assert.Contains(2, plan.ChildRowNumbers);
    }

    /// <summary>
    /// Every group's children are written, across more than one group, so
    /// "collapse all" reaches every relationship and not just the first.
    /// </summary>
    [Fact]
    public void Every_groups_children_are_written()
    {
        OutlineCollapsePlan plan = OutlineCollapsePlanner.Plan(
            [
                new OutlineRowLevel(1, OutlineCollapsePlanner.TopLevelOutlineLevel),
                new OutlineRowLevel(2, OutlineCollapsePlanner.TopLevelOutlineLevel + 1),
                new OutlineRowLevel(3, OutlineCollapsePlanner.TopLevelOutlineLevel),
                new OutlineRowLevel(4, OutlineCollapsePlanner.TopLevelOutlineLevel + 1),
            ],
            OutlineCollapseDirection.Collapse);

        Assert.Equal([2, 4], plan.ChildRowNumbers);
    }

    /// <summary>
    /// The rows come back in ascending order whatever order they arrive in, so
    /// the request writes a deterministic sequence.
    /// </summary>
    [Fact]
    public void The_rows_come_back_in_ascending_order()
    {
        OutlineCollapsePlan plan = OutlineCollapsePlanner.Plan(
            [
                new OutlineRowLevel(4, OutlineCollapsePlanner.TopLevelOutlineLevel + 1),
                new OutlineRowLevel(2, OutlineCollapsePlanner.TopLevelOutlineLevel + 1),
                new OutlineRowLevel(3, OutlineCollapsePlanner.TopLevelOutlineLevel),
            ],
            OutlineCollapseDirection.Collapse);

        Assert.Equal([2, 4], plan.ChildRowNumbers);
    }

    /// <summary>
    /// A duplicated row is written once: the request names a set of rows, not
    /// a multiset.
    /// </summary>
    [Fact]
    public void A_duplicated_row_is_written_once()
    {
        OutlineCollapsePlan plan = OutlineCollapsePlanner.Plan(
            [
                new OutlineRowLevel(2, OutlineCollapsePlanner.TopLevelOutlineLevel + 1),
                new OutlineRowLevel(2, OutlineCollapsePlanner.TopLevelOutlineLevel + 1),
            ],
            OutlineCollapseDirection.Collapse);

        Assert.Equal([2], plan.ChildRowNumbers);
    }

    /// <summary>
    /// A sheet with no group writes nothing: there is no relationship to
    /// collapse or expand, and the request is a no-op rather than an error.
    /// </summary>
    [Fact]
    public void A_sheet_with_no_group_writes_nothing()
    {
        OutlineCollapsePlan plan = OutlineCollapsePlanner.Plan(
            [
                new OutlineRowLevel(1, OutlineCollapsePlanner.TopLevelOutlineLevel),
                new OutlineRowLevel(2, OutlineCollapsePlanner.TopLevelOutlineLevel),
            ],
            OutlineCollapseDirection.Collapse);

        Assert.Empty(plan.ChildRowNumbers);
        Assert.Equal(0, plan.RowCount);
        Assert.True(plan.IsEmpty);
    }

    /// <summary>
    /// The expand request names the same rows as the collapse request: only
    /// the direction differs, so the two commands can never drift apart in
    /// which rows they touch.
    /// </summary>
    [Fact]
    public void Expand_and_collapse_name_the_same_rows()
    {
        IReadOnlyList<OutlineRowLevel> rows =
        [
            new OutlineRowLevel(1, OutlineCollapsePlanner.TopLevelOutlineLevel),
            new OutlineRowLevel(2, OutlineCollapsePlanner.TopLevelOutlineLevel + 1),
        ];

        OutlineCollapsePlan collapse = OutlineCollapsePlanner.Plan(rows, OutlineCollapseDirection.Collapse);
        OutlineCollapsePlan expand = OutlineCollapsePlanner.Plan(rows, OutlineCollapseDirection.Expand);

        Assert.Equal(collapse.ChildRowNumbers, expand.ChildRowNumbers);
        Assert.Equal(OutlineCollapseDirection.Collapse, collapse.Direction);
        Assert.Equal(OutlineCollapseDirection.Expand, expand.Direction);
    }

    /// <summary>
    /// A row deeper than the top level is grouped whatever its exact level: a
    /// deeper nesting is still a child of some group.
    /// </summary>
    [Fact]
    public void A_deeper_level_is_still_a_grouped_child()
    {
        OutlineCollapsePlan plan = OutlineCollapsePlanner.Plan(
            [new OutlineRowLevel(2, OutlineCollapsePlanner.TopLevelOutlineLevel + 2)],
            OutlineCollapseDirection.Collapse);

        Assert.Equal([2], plan.ChildRowNumbers);
    }

    /// <summary>
    /// A null row list is a programming error, not a sheet with no group: it
    /// is refused rather than silently planned as empty.
    /// </summary>
    [Fact]
    public void A_null_row_list_is_refused()
    {
        Assert.Throws<ArgumentNullException>(
            () => OutlineCollapsePlanner.Plan(null!, OutlineCollapseDirection.Collapse));
    }
}
