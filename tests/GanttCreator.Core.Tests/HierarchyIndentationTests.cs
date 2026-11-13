namespace GanttCreator.Core.Tests;

/// <summary>
/// Contracts for <see cref="HierarchyIndentation"/> (R5.18i): the indent a body row's
/// description cell carries is the outline group the same pass writes, and nothing else.
/// </summary>
public sealed class HierarchyIndentationTests
{
    private static OutlineGroupPlan PlanOf(params OutlineGroup[] groups) => new(groups, []);

    [Fact]
    public void A_row_the_plan_does_not_group_carries_no_indent()
    {
        IReadOnlyDictionary<int, int> indents = HierarchyIndentation.Plan(PlanOf());

        Assert.Equal(0, HierarchyIndentation.IndentFor(indents, 1));
        Assert.Equal(0, HierarchyIndentation.IndentFor(indents, 7));
    }

    [Fact]
    public void Every_grouped_child_carries_one_indent_step()
    {
        IReadOnlyDictionary<int, int> indents =
            HierarchyIndentation.Plan(PlanOf(new OutlineGroup(1, [2, 3])));

        Assert.Equal(HierarchyIndentation.ChildIndentLevel, HierarchyIndentation.IndentFor(indents, 2));
        Assert.Equal(HierarchyIndentation.ChildIndentLevel, HierarchyIndentation.IndentFor(indents, 3));

        // The parent keeps the bracket and carries no indent of its own; indenting it
        // would put the +/- control one step inside its own children.
        Assert.Equal(0, HierarchyIndentation.IndentFor(indents, 1));
    }

    [Fact]
    public void A_promoted_child_is_reset_to_zero()
    {
        // The child of the first plan is the top-level row of the second: the same row
        // number, no group, and therefore no indent. This is the promotion the outline
        // reset pass already covers, and the indent has to follow it or a promoted row
        // keeps a step that its now-missing bracket no longer explains.
        IReadOnlyDictionary<int, int> whileGrouped =
            HierarchyIndentation.Plan(PlanOf(new OutlineGroup(1, [2])));
        IReadOnlyDictionary<int, int> afterPromotion = HierarchyIndentation.Plan(PlanOf());

        Assert.Equal(1, HierarchyIndentation.IndentFor(whileGrouped, 2));
        Assert.Equal(0, HierarchyIndentation.IndentFor(afterPromotion, 2));
    }

    [Fact]
    public void A_child_the_planner_could_not_group_is_not_indented()
    {
        // The planner reports a child that cannot sit under its parent -- above it, or in
        // a non-contiguous run -- in RowsToUngroup. That row has no bracket, so it must
        // have no indent either; the plan carries both facts, so one read is enough.
        OutlineGroupPlan plan = new([], [2]);

        IReadOnlyDictionary<int, int> indents = HierarchyIndentation.Plan(plan);

        Assert.Equal(0, HierarchyIndentation.IndentFor(indents, 2));
    }

    [Fact]
    public void The_step_is_one_less_than_the_supported_hierarchy_depth()
    {
        // "Indent by depth - 1" is one step today because the validator permits exactly two
        // levels. If MaxDepth grows without the outline and this map growing with it, the
        // indent would understate the nesting; failing here makes that a decision rather
        // than a silent regression.
        Assert.Equal(EntityHierarchyCatalog.MaxDepth - 1, HierarchyIndentation.ChildIndentLevel);
        Assert.Equal(1, HierarchyIndentation.ChildIndentLevel);
    }

    [Fact]
    public void A_null_plan_is_refused_rather_than_treated_as_empty() =>
        _ = Assert.Throws<ArgumentNullException>(() => HierarchyIndentation.Plan(null!));

    [Fact]
    public void A_null_indent_map_is_refused_rather_than_read_as_zero() =>
        _ = Assert.Throws<ArgumentNullException>(() => HierarchyIndentation.IndentFor(null!, 1));
}
