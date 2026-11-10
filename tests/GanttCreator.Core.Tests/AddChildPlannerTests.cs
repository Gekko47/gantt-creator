using GanttCreator.Core;

namespace GanttCreator.Core.Tests;

/// <summary>
/// Tests for the R5.3 WI-E add-child planner: a new row can be planned as a
/// child of a selected parent, and every rule that forbids it produces a
/// TYPED refusal rather than a silent success.
/// </summary>
public sealed class AddChildPlannerTests
{
    [Fact]
    public void Evaluate_accepts_a_child_capable_type_under_a_top_level_activity()
    {
        Assert.Null(AddChildPlanner.Evaluate(
            GanttEntityType.AsPlannedActivity,
            GanttRowId.Parse("G-0000000000000000000000000000000a"),
            parentOfParentId: null,
            existingChildCount: 0,
            GanttEntityType.AsPlannedMilestone));
    }

    [Fact]
    public void Evaluate_refuses_a_leaf_parent() =>
        Assert.Equal(
            AddChildRefusalReason.ParentCannotOwnChildren,
            AddChildPlanner.Evaluate(
                GanttEntityType.AsPlannedMilestone,
                GanttRowId.Parse("G-0000000000000000000000000000000a"),
                parentOfParentId: null,
                existingChildCount: 0,
                GanttEntityType.AsPlannedMilestone));

    [Fact]
    public void Evaluate_refuses_a_child_type_that_cannot_be_a_child() =>
        Assert.Equal(
            AddChildRefusalReason.TypeCannotBeChild,
            AddChildPlanner.Evaluate(
                GanttEntityType.AsPlannedActivity,
                GanttRowId.Parse("G-0000000000000000000000000000000a"),
                parentOfParentId: null,
                existingChildCount: 0,
                GanttEntityType.Splitter));

    [Fact]
    public void Evaluate_refuses_a_parent_that_is_itself_a_child() =>
        Assert.Equal(
            AddChildRefusalReason.TooDeep,
            AddChildPlanner.Evaluate(
                GanttEntityType.AsPlannedActivity,
                GanttRowId.Parse("G-0000000000000000000000000000000a"),
                GanttRowId.Parse("G-0000000000000000000000000000000b"),
                existingChildCount: 0,
                GanttEntityType.AsPlannedMilestone));

    [Fact]
    public void Evaluate_refuses_a_parent_at_capacity() =>
        Assert.Equal(
            AddChildRefusalReason.ParentAtCapacity,
            AddChildPlanner.Evaluate(
                GanttEntityType.AsPlannedActivity,
                GanttRowId.Parse("G-0000000000000000000000000000000a"),
                parentOfParentId: null,
                existingChildCount: EntityHierarchyCatalog.MaxChildrenPerParent,
                GanttEntityType.AsPlannedMilestone));

    [Fact]
    public void Evaluate_refuses_an_unresolved_parent() =>
        Assert.Equal(
            AddChildRefusalReason.SelectionUnknown,
            AddChildPlanner.Evaluate(
                parentType: null,
                parentId: null,
                parentOfParentId: null,
                existingChildCount: 0,
                GanttEntityType.AsPlannedMilestone));

    [Fact]
    public void An_activity_with_room_is_an_eligible_parent()
    {
        Assert.True(AddChildPlanner.IsEligibleParent(
            GanttEntityType.AsPlannedActivity, parentId: null, existingChildCount: 6));
    }

    [Fact]
    public void A_leaf_type_is_not_an_eligible_parent()
    {
        // Positive test for the leaf branch: a milestone may be a child but
        // never owns one.
        Assert.False(AddChildPlanner.IsEligibleParent(
            GanttEntityType.AsPlannedMilestone, parentId: null, existingChildCount: 0));
    }

    [Fact]
    public void A_parent_that_is_itself_a_child_is_not_eligible()
    {
        // Positive test for the depth branch: a child activity would make the
        // new row depth 3.
        Assert.False(AddChildPlanner.IsEligibleParent(
            GanttEntityType.AsPlannedActivity,
            GanttRowId.Parse("G-0000000000000000000000000000000a"),
            existingChildCount: 0));
    }

    [Fact]
    public void A_parent_at_seven_children_is_not_eligible()
    {
        // Positive test for the capacity branch.
        Assert.False(AddChildPlanner.IsEligibleParent(
            GanttEntityType.AsPlannedActivity,
            parentId: null,
            existingChildCount: EntityHierarchyCatalog.MaxChildrenPerParent));
    }
    private static GanttRowId NewId() => GanttRowId.New();

    private static GanttEvent Row(
        int rowNumber,
        GanttRowId id,
        GanttEntityType type = GanttEntityType.AsPlannedActivity,
        GanttRowId? parentId = null) =>
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

    [Fact]
    public void A_child_is_planned_under_its_parent()
    {
        GanttRowId parentId = NewId();
        GanttRowId childId = NewId();
        List<GanttEvent> rows = [Row(1, parentId)];

        AddChildPlan plan = AddChildPlanner.Plan(
            rows,
            parentId,
            GanttEntityType.AsBuiltActivity,
            childId);

        Assert.True(plan.Succeeded);
        AddChildWrite write = Assert.Single(plan.Writes!);
        Assert.Equal(childId, write.RowId);
        Assert.Equal(parentId, write.ParentId);
        Assert.Equal(0, write.SiblingOrder);
    }

    [Fact]
    public void A_second_child_takes_the_next_sibling_order()
    {
        GanttRowId parentId = NewId();
        GanttRowId firstChildId = NewId();
        GanttRowId childId = NewId();
        List<GanttEvent> rows = [Row(1, parentId), Row(2, firstChildId, parentId: parentId)];

        AddChildPlan plan = AddChildPlanner.Plan(
            rows,
            parentId,
            GanttEntityType.AsBuiltActivity,
            childId);

        Assert.True(plan.Succeeded);
        Assert.Equal(1, Assert.Single(plan.Writes!).SiblingOrder);
    }

    [Fact]
    public void An_unknown_selection_is_refused()
    {
        List<GanttEvent> rows = [Row(1, NewId())];

        AddChildPlan plan = AddChildPlanner.Plan(
            rows,
            NewId(),
            GanttEntityType.AsBuiltActivity,
            NewId());

        Assert.False(plan.Succeeded);
        Assert.Equal(AddChildRefusalReason.SelectionUnknown, plan.Refusal);
    }

    [Fact]
    public void A_parent_that_may_not_own_children_is_refused()
    {
        // A milestone is a leaf: it may be a child, but it never owns one.
        GanttRowId parentId = NewId();
        List<GanttEvent> rows = [Row(1, parentId, GanttEntityType.AsPlannedMilestone)];

        AddChildPlan plan = AddChildPlanner.Plan(
            rows,
            parentId,
            GanttEntityType.AsBuiltActivity,
            NewId());

        Assert.False(plan.Succeeded);
        Assert.Equal(AddChildRefusalReason.ParentCannotOwnChildren, plan.Refusal);
    }

    [Fact]
    public void A_type_that_may_not_be_a_child_is_refused()
    {
        // A Spacer is standalone: it is never a child of anything.
        GanttRowId parentId = NewId();
        List<GanttEvent> rows = [Row(1, parentId)];

        AddChildPlan plan = AddChildPlanner.Plan(
            rows,
            parentId,
            GanttEntityType.Spacer,
            NewId());

        Assert.False(plan.Succeeded);
        Assert.Equal(AddChildRefusalReason.TypeCannotBeChild, plan.Refusal);
    }

    [Fact]
    public void A_parent_at_seven_children_is_refused()
    {
        GanttRowId parentId = NewId();
        List<GanttEvent> rows = [Row(1, parentId)];
        for (int index = 0; index < EntityHierarchyCatalog.MaxChildrenPerParent; index++)
        {
            rows.Add(Row(2 + index, NewId(), parentId: parentId));
        }

        AddChildPlan plan = AddChildPlanner.Plan(
            rows,
            parentId,
            GanttEntityType.AsBuiltActivity,
            NewId());

        Assert.False(plan.Succeeded);
        Assert.Equal(AddChildRefusalReason.ParentAtCapacity, plan.Refusal);
    }

    [Fact]
    public void A_parent_that_is_itself_a_child_is_refused()
    {
        // The new row would be depth 3, which the catalogue refuses rather
        // than supports.
        GanttRowId grandParentId = NewId();
        GanttRowId parentId = NewId();
        List<GanttEvent> rows =
        [
            Row(1, grandParentId),
            Row(2, parentId, parentId: grandParentId),
        ];

        AddChildPlan plan = AddChildPlanner.Plan(
            rows,
            parentId,
            GanttEntityType.AsBuiltActivity,
            NewId());

        Assert.False(plan.Succeeded);
        Assert.Equal(AddChildRefusalReason.TooDeep, plan.Refusal);
    }

    [Fact]
    public void A_critical_interval_is_never_a_parent()
    {
        // The R5.3 owner ruling: Critical Interval is child-only and never
        // owns. It is therefore refused as a PARENT, not given a narrower
        // list of child types.
        GanttRowId parentId = NewId();
        List<GanttEvent> rows = [Row(1, parentId, GanttEntityType.CriticalInterval)];

        AddChildPlan plan = AddChildPlanner.Plan(
            rows,
            parentId,
            GanttEntityType.AsBuiltActivity,
            NewId());

        Assert.False(plan.Succeeded);
        Assert.Equal(AddChildRefusalReason.ParentCannotOwnChildren, plan.Refusal);
        Assert.Empty(AddChildPlanner.PermittedChildTypes(GanttEntityType.CriticalInterval));
    }

    [Fact]
    public void A_parent_that_may_not_own_children_permits_no_child_types()
    {
        Assert.Empty(AddChildPlanner.PermittedChildTypes(GanttEntityType.AsPlannedMilestone));
        Assert.Empty(AddChildPlanner.PermittedChildTypes(GanttEntityType.Splitter));
    }

    [Fact]
    public void Every_permitted_type_is_accepted_by_the_planner()
    {
        // The Ribbon's list and the planner's acceptance are the SAME
        // projection, so a type the UI offers can never be refused by the
        // planner. This is the test that keeps them agreeing.
        GanttRowId parentId = NewId();
        List<GanttEvent> rows = [Row(1, parentId, GanttEntityType.AsPlannedActivity)];

        foreach (GanttEntityType childType in AddChildPlanner.PermittedChildTypes(GanttEntityType.AsPlannedActivity))
        {
            AddChildPlan plan = AddChildPlanner.Plan(rows, parentId, childType, NewId());

            Assert.True(plan.Succeeded, $"The planner refused a permitted type: {childType}.");
        }
    }

    [Fact]
    public void The_permitted_list_is_the_catalogue_s_child_capable_types()
    {
        // The list is PROJECTED from the catalogue, never authored beside it,
        // so a change to the hierarchy rules reaches the Ribbon.
        Assert.Equal(
            EntityHierarchyCatalog.TypesThatMayBeChild().Order(),
            AddChildPlanner.PermittedChildTypes(GanttEntityType.AsPlannedActivity).Order());
    }
}
