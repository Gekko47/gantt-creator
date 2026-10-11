namespace GanttCreator.Core.Tests;

/// <summary>
/// Tests for <see cref="HierarchyMutations.PlanPromotion"/>, the promote-on-delete
/// rule (R4.7A D6).
/// </summary>
public sealed class HierarchyMutationsTests
{
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

    /// <summary>
    /// Deleting a top-level parent promotes its children to top-level: the
    /// inherited parent is null, so each child's ParentId must be cleared.
    /// </summary>
    [Fact]
    public void Deleting_a_top_level_parent_promotes_children_to_top_level()
    {
        GanttRowId parent = NewId();
        GanttRowId firstChild = NewId();
        GanttRowId secondChild = NewId();
        List<GanttEvent> rows =
        [
            Row(2, parent),
            Row(3, firstChild, parentId: parent),
            Row(4, secondChild, parentId: parent),
        ];

        HierarchyPromotionPlan plan = HierarchyMutations.PlanPromotion(rows, parent);

        Assert.True(plan.PromotedAnything);
        Assert.Equal([firstChild, secondChild], plan.PromotedChildren);
        Assert.Null(plan.InheritedParentId);
    }

    /// <summary>
    /// Deleting a nested parent promotes its children to the deleted row's own
    /// parent -- they become siblings of the row being deleted, not top-level.
    /// </summary>
    [Fact]
    public void Deleting_a_nested_parent_promotes_children_to_the_grandparent()
    {
        GanttRowId grandparent = NewId();
        GanttRowId parent = NewId();
        GanttRowId child = NewId();
        List<GanttEvent> rows =
        [
            Row(2, grandparent),
            Row(3, parent, parentId: grandparent),
            Row(4, child, parentId: parent),
        ];

        HierarchyPromotionPlan plan = HierarchyMutations.PlanPromotion(rows, parent);

        Assert.Equal([child], plan.PromotedChildren);
        Assert.Equal(grandparent, plan.InheritedParentId);
    }

    /// <summary>
    /// Deleting a leaf promotes nothing. The plan is still well formed, so a
    /// caller need not special-case a row with no children.
    /// </summary>
    [Fact]
    public void Deleting_a_leaf_promotes_nothing()
    {
        GanttRowId parent = NewId();
        GanttRowId leaf = NewId();
        List<GanttEvent> rows = [Row(2, parent), Row(3, leaf, parentId: parent)];

        HierarchyPromotionPlan plan = HierarchyMutations.PlanPromotion(rows, leaf);

        Assert.False(plan.PromotedAnything);
        Assert.Empty(plan.PromotedChildren);
        Assert.Equal(parent, plan.InheritedParentId);
    }

    /// <summary>
    /// Promoted children keep their own identities. The plan reports ids only; it
    /// never proposes replacing them, because a promoted child is the same entity
    /// that simply lost its parent.
    /// </summary>
    [Fact]
    public void Promoted_children_keep_their_own_identities()
    {
        GanttRowId parent = NewId();
        GanttRowId child = NewId();
        List<GanttEvent> rows = [Row(2, parent), Row(3, child, parentId: parent)];

        HierarchyPromotionPlan plan = HierarchyMutations.PlanPromotion(rows, parent);

        Assert.Equal([child], plan.PromotedChildren);
        Assert.DoesNotContain(parent, plan.PromotedChildren);
    }

    /// <summary>
    /// Determinism: children are reported in worksheet row order regardless of the
    /// order the rows are supplied in, so the same table always produces the same
    /// writes in the same sequence.
    /// </summary>
    [Fact]
    public void Promoted_children_are_ordered_by_row_not_by_input_order()
    {
        GanttRowId parent = NewId();
        GanttRowId first = NewId();
        GanttRowId second = NewId();
        GanttRowId third = NewId();
        List<GanttEvent> rows =
        [
            Row(2, parent),
            Row(5, third, parentId: parent),
            Row(3, first, parentId: parent),
            Row(4, second, parentId: parent),
        ];

        HierarchyPromotionPlan plan = HierarchyMutations.PlanPromotion(rows, parent);

        Assert.Equal([first, second, third], plan.PromotedChildren);
    }

    /// <summary>
    /// A grandchild is unaffected by deleting its grandparent -- the child's own
    /// parent survives, so nothing about that child changes. This is the case a
    /// naive "re-parent everything under the deleted id" implementation gets
    /// wrong.
    /// </summary>
    [Fact]
    public void Deleting_a_grandparent_does_not_touch_its_grandchildren()
    {
        GanttRowId grandparent = NewId();
        GanttRowId parent = NewId();
        GanttRowId child = NewId();
        List<GanttEvent> rows =
        [
            Row(2, grandparent),
            Row(3, parent, parentId: grandparent),
            Row(4, child, parentId: parent),
        ];

        HierarchyPromotionPlan plan = HierarchyMutations.PlanPromotion(rows, grandparent);

        Assert.Equal([parent], plan.PromotedChildren);
        Assert.DoesNotContain(child, plan.PromotedChildren);
    }

    /// <summary>
    /// Deleting an id that is not present yields an empty plan rather than
    /// throwing. The row may already have been removed by a prior step, and a
    /// delete command must be idempotent with respect to that.
    /// </summary>
    [Fact]
    public void Deleting_an_absent_id_yields_an_empty_plan()
    {
        GanttRowId present = NewId();
        GanttRowId absent = NewId();
        List<GanttEvent> rows = [Row(2, present), Row(3, NewId(), parentId: present)];

        HierarchyPromotionPlan plan = HierarchyMutations.PlanPromotion(rows, absent);

        Assert.False(plan.PromotedAnything);
        Assert.Null(plan.InheritedParentId);
    }

    /// <summary>
    /// The plan does not mutate its input. The caller still holds the original
    /// events, which is what lets it apply the plan and, on failure, apply nothing.
    /// </summary>
    [Fact]
    public void Planning_does_not_mutate_the_input_rows()
    {
        GanttRowId parent = NewId();
        GanttRowId child = NewId();
        List<GanttEvent> rows = [Row(2, parent), Row(3, child, parentId: parent)];
        List<GanttEvent> snapshot = [.. rows];

        _ = HierarchyMutations.PlanPromotion(rows, parent);

        Assert.Equal(snapshot, rows);
    }

    [Fact]
    public void Null_arguments_are_rejected()
    {
        Assert.Throws<ArgumentNullException>(() => HierarchyMutations.PlanPromotion(null!, NewId()));
        Assert.Throws<ArgumentNullException>(() => HierarchyMutations.PlanPromotion([], null!));
    }

    /// <summary>
    /// A rewrite is required whenever anything was promoted, including the
    /// top-level case where the new value is blank rather than a different id --
    /// the cell must still be cleared or the child would keep pointing at a row
    /// that no longer exists.
    /// </summary>
    [Fact]
    public void A_rewrite_is_required_even_when_the_new_parent_is_blank()
    {
        GanttRowId parent = NewId();
        GanttRowId child = NewId();

        HierarchyPromotionPlan plan = HierarchyMutations.PlanPromotion(
            [Row(2, parent), Row(3, child, parentId: parent)],
            parent);

        Assert.Null(plan.InheritedParentId);
        Assert.True(HierarchyMutations.RequiresParentRewrite(plan));
    }

    /// <summary>
    /// R5.3 D2: promoting the first of two children moves the promoted row's
    /// content to the bottom of the run and resequences the remaining child
    /// densely from zero.
    /// </summary>
    [Fact]
    public void Promoting_the_first_child_moves_it_to_the_bottom_and_resequences()
    {
        GanttRowId parent = NewId();
        GanttRowId first = NewId();
        GanttRowId second = NewId();
        List<GanttEvent> rows =
        [
            Row(2, parent),
            Row(3, first, parentId: parent),
            Row(4, second, parentId: parent),
        ];

        PromotePlan plan = HierarchyMutations.PlanPromoteToParent(rows, first);

        Assert.True(plan.Succeeded);
        Assert.Equal(first, plan.TargetId);
        Assert.Equal(parent, plan.ClearedParentId);
        Assert.Equal(4, plan.MoveToBodyIndex);
        // The remaining run resequences densely, and the
        // promoted row joins the top-level run at its new
        // position (the run bottom): the parent keeps 0 and
        // the promoted row takes 1, so its stale child-run
        // value is never carried over (R5.3 D2, amended
        // 2026-10-11).
        Assert.Equal(
            [
                new SiblingOrderWrite(second, 0),
                new SiblingOrderWrite(parent, 0),
                new SiblingOrderWrite(first, 1),
            ],
            plan.SiblingOrderWrites);
    }

    /// <summary>
    /// Promoting the last child is a clear plus a one-row resequence.
    /// </summary>
    [Fact]
    public void Promoting_the_last_child_clears_with_no_gap_to_close()
    {
        GanttRowId parent = NewId();
        GanttRowId first = NewId();
        GanttRowId second = NewId();
        List<GanttEvent> rows =
        [
            Row(2, parent),
            Row(3, first, parentId: parent),
            Row(4, second, parentId: parent),
        ];

        PromotePlan plan = HierarchyMutations.PlanPromoteToParent(rows, second);

        Assert.True(plan.Succeeded);
        Assert.Equal(4, plan.MoveToBodyIndex);
        // The remaining run resequences densely, and the
        // promoted row joins the top-level run at its new
        // position (the run bottom, unchanged here): the
        // parent keeps 0 and the promoted row takes 1
        // (R5.3 D2, amended 2026-10-11).
        Assert.Equal(
            [
                new SiblingOrderWrite(first, 0),
                new SiblingOrderWrite(parent, 0),
                new SiblingOrderWrite(second, 1),
            ],
            plan.SiblingOrderWrites);
    }

    /// <summary>
    /// A promote between top-level rows inserts the promoted
    /// row into the top-level run at its new position -- the
    /// bottom of its former sibling run -- and the run
    /// resequences densely around it: the rows below it shift
    /// to make room, and the promoted row's stale child-run
    /// value is never carried over (R5.3 D2, amended
    /// 2026-10-11).
    /// </summary>
    [Fact]
    public void Promoting_a_child_between_top_level_rows_resequences_the_run_it_joins()
    {
        GanttRowId top = NewId();
        GanttRowId parent = NewId();
        GanttRowId first = NewId();
        GanttRowId second = NewId();
        GanttRowId below = NewId();
        List<GanttEvent> rows =
        [
            Row(2, top),
            Row(3, parent),
            Row(4, first, parentId: parent),
            Row(5, second, parentId: parent),
            Row(6, below),
        ];

        PromotePlan plan = HierarchyMutations.PlanPromoteToParent(rows, first);

        Assert.True(plan.Succeeded);
        Assert.Equal(5, plan.MoveToBodyIndex);
        // The remaining run resequences densely; the
        // top-level run resequences around the joined
        // row: both top-level rows above it keep their
        // order, the promoted row takes its position,
        // and the row below shifts to make room.
        Assert.Equal(
            [
                new SiblingOrderWrite(second, 0),
                new SiblingOrderWrite(top, 0),
                new SiblingOrderWrite(parent, 1),
                new SiblingOrderWrite(first, 2),
                new SiblingOrderWrite(below, 3),
            ],
            plan.SiblingOrderWrites);
    }

    /// <summary>
    /// A top-level row is not a child, so promote refuses with
    /// <c>NotAChild</c> and mutates nothing.
    /// </summary>
    [Fact]
    public void Promoting_a_top_level_row_is_refused_as_not_a_child()
    {
        GanttRowId parent = NewId();
        List<GanttEvent> rows = [Row(2, parent), Row(3, NewId(), parentId: parent)];

        PromotePlan plan = HierarchyMutations.PlanPromoteToParent(rows, parent);

        Assert.False(plan.Succeeded);
        Assert.Equal(ParentChildRefusalReason.NotAChild, plan.Refusal);
        Assert.Empty(plan.SiblingOrderWrites);
    }

    /// <summary>
    /// R5.3 D1: a Critical Interval can never be promoted, even though it is
    /// always a child.
    /// </summary>
    [Fact]
    public void Promoting_a_critical_interval_is_forbidden()
    {
        GanttRowId parent = NewId();
        GanttRowId interval = NewId();
        List<GanttEvent> rows =
        [
            Row(2, parent),
            Row(3, interval, GanttEntityType.CriticalInterval, parent),
        ];

        PromotePlan plan = HierarchyMutations.PlanPromoteToParent(rows, interval);

        Assert.False(plan.Succeeded);
        Assert.Equal(ParentChildRefusalReason.PromotionForbiddenForType, plan.Refusal);
    }

    /// <summary>
    /// A child that owns children of its own cannot be promoted: the plan
    /// refuses rather than orphaning a third level.
    /// </summary>
    [Fact]
    public void Promoting_a_row_that_owns_children_is_refused()
    {
        GanttRowId parent = NewId();
        GanttRowId middle = NewId();
        GanttRowId leaf = NewId();
        List<GanttEvent> rows =
        [
            Row(2, parent),
            Row(3, middle, parentId: parent),
            Row(4, leaf, parentId: middle),
        ];

        PromotePlan plan = HierarchyMutations.PlanPromoteToParent(rows, middle);

        Assert.False(plan.Succeeded);
        Assert.Equal(ParentChildRefusalReason.HasChildren, plan.Refusal);
    }

    /// <summary>
    /// R5.3 D3: demoting a top-level row sets its parent to the nearest
    /// eligible row above and assigns it the tail of the new run.
    /// </summary>
    [Fact]
    public void Demoting_assigns_the_nearest_eligible_parent_above()
    {
        GanttRowId parent = NewId();
        GanttRowId other = NewId();
        GanttRowId target = NewId();
        List<GanttEvent> rows =
        [
            Row(2, parent),
            Row(3, other),
            Row(4, target),
        ];

        DemotePlan plan = HierarchyMutations.PlanDemoteToChild(rows, target);

        Assert.True(plan.Succeeded);
        Assert.Equal(target, plan.TargetId);
        Assert.Equal(other, plan.NewParentId);
        Assert.Contains(new SiblingOrderWrite(target, 0), plan.SiblingOrderWrites);
    }

    /// <summary>
    /// A selection directly beneath a parent's existing child run
    /// extends the run: the demotion lands on that parent, the tail
    /// assignment follows the run's new length, and the run stays one
    /// contiguous outline group (R5.3 D3).
    /// </summary>
    [Fact]
    public void Demoting_directly_beneath_a_child_run_extends_the_run()
    {
        GanttRowId parent = NewId();
        GanttRowId child = NewId();
        GanttRowId target = NewId();
        List<GanttEvent> rows =
        [
            Row(2, parent),
            Row(3, child, parentId: parent),
            Row(4, target),
        ];

        DemotePlan plan = HierarchyMutations.PlanDemoteToChild(rows, target);

        Assert.True(plan.Succeeded);
        Assert.Equal(parent, plan.NewParentId);
        Assert.Contains(new SiblingOrderWrite(target, 1), plan.SiblingOrderWrites);
    }

    /// <summary>
    /// A selection separated from the target parent's child run by an
    /// unrelated row MOVES to sit directly beneath the run: the demoted row
    /// takes the tail of the new run, its content moves to the run bottom,
    /// and the demotion stays one contiguous outline group (R5.3 D3, amended
    /// 2026-10-11).
    /// </summary>
    [Fact]
    public void Demoting_across_an_unrelated_row_moves_beneath_the_target_run()
    {
        GanttRowId parent = NewId();
        GanttRowId child = NewId();
        GanttRowId milestone = NewId();
        GanttRowId target = NewId();
        List<GanttEvent> rows =
        [
            Row(2, parent),
            Row(3, child, parentId: parent),
            Row(4, milestone, GanttEntityType.AsPlannedMilestone),
            Row(5, target),
        ];

        DemotePlan plan = HierarchyMutations.PlanDemoteToChild(rows, target);

        Assert.True(plan.Succeeded);
        Assert.Equal(target, plan.TargetId);
        Assert.Equal(parent, plan.NewParentId);
        Assert.Equal(4, plan.MoveToBodyIndex);
        Assert.Empty(plan.RehomedChildIds);
        Assert.Contains(new SiblingOrderWrite(target, 1), plan.SiblingOrderWrites);
    }

    /// <summary>
    /// Demoting a row that owns children re-homes them: the demoted row and
    /// its children all join the new parent's run in row order, so nobody
    /// nests three deep (R5.3 D3, amended 2026-10-11).
    /// </summary>
    [Fact]
    public void Demoting_a_parent_rehomes_its_children_to_the_new_parent()
    {
        GanttRowId parent = NewId();
        GanttRowId target = NewId();
        GanttRowId first = NewId();
        GanttRowId second = NewId();
        List<GanttEvent> rows =
        [
            Row(2, parent),
            Row(3, target),
            Row(4, first, parentId: target),
            Row(5, second, parentId: target),
        ];

        DemotePlan plan = HierarchyMutations.PlanDemoteToChild(rows, target);

        Assert.True(plan.Succeeded);
        Assert.Equal(target, plan.TargetId);
        Assert.Equal(parent, plan.NewParentId);
        Assert.Equal(3, plan.MoveToBodyIndex);
        Assert.Equal([first, second], plan.RehomedChildIds);
        Assert.Contains(new SiblingOrderWrite(target, 0), plan.SiblingOrderWrites);
        Assert.Contains(new SiblingOrderWrite(first, 1), plan.SiblingOrderWrites);
        Assert.Contains(new SiblingOrderWrite(second, 2), plan.SiblingOrderWrites);
    }

    /// <summary>
    /// Capacity counts the demoted row TOGETHER with its re-homed children:
    /// a run that would exceed seven by taking them all refuses with
    /// <c>ParentUnavailable</c> (R5.3 D3, amended 2026-10-11).
    /// </summary>
    [Fact]
    public void Demoting_a_parent_counts_rehomed_children_against_capacity()
    {
        GanttRowId full = NewId();
        GanttRowId target = NewId();
        GanttRowId first = NewId();
        GanttRowId second = NewId();
        List<GanttEvent> rows = [Row(2, full)];
        for (var i = 0; i < EntityHierarchyCatalog.MaxChildrenPerParent - 1; i++)
        {
            rows.Add(Row(3 + i, NewId(), parentId: full));
        }

        rows.Add(Row(3 + EntityHierarchyCatalog.MaxChildrenPerParent, target));
        rows.Add(Row(4 + EntityHierarchyCatalog.MaxChildrenPerParent, first, parentId: target));
        rows.Add(Row(5 + EntityHierarchyCatalog.MaxChildrenPerParent, second, parentId: target));

        DemotePlan plan = HierarchyMutations.PlanDemoteToChild(rows, target);

        Assert.False(plan.Succeeded);
        Assert.Equal(ParentChildRefusalReason.ParentUnavailable, plan.Refusal);
        Assert.Empty(plan.SiblingOrderWrites);
    }

    /// <summary>
    /// A delineator between the parent and the selection is skipped by the
    /// scan -- and the demotion MOVES to sit directly beneath the parent's
    /// run, so it stays one contiguous outline group (R5.3 D3, amended
    /// 2026-10-11).
    /// </summary>
    [Fact]
    public void Demoting_beneath_a_delineator_moves_beneath_the_parent_run()
    {
        GanttRowId parent = NewId();
        GanttRowId line = NewId();
        GanttRowId target = NewId();
        List<GanttEvent> rows =
        [
            Row(2, parent),
            Row(3, line, GanttEntityType.Delineator),
            Row(4, target),
        ];

        DemotePlan plan = HierarchyMutations.PlanDemoteToChild(rows, target);

        Assert.True(plan.Succeeded);
        Assert.Equal(target, plan.TargetId);
        Assert.Equal(parent, plan.NewParentId);
        Assert.Equal(3, plan.MoveToBodyIndex);
        Assert.Empty(plan.RehomedChildIds);
        Assert.Contains(new SiblingOrderWrite(target, 0), plan.SiblingOrderWrites);
    }

    /// <summary>
    /// A selection adjacent to a full group's run extends that run,
    /// and no parent further up can be adjacent either -- its run
    /// would be broken by the full group's children. The capacity
    /// rule therefore refuses with <c>ParentUnavailable</c> instead
    /// of skipping to a landing the outline cannot represent
    /// (R5.3 D3).
    /// </summary>
    [Fact]
    public void Demoting_adjacent_to_a_full_group_is_refused_as_parent_unavailable()
    {
        GanttRowId upper = NewId();
        GanttRowId full = NewId();
        GanttRowId target = NewId();
        List<GanttEvent> rows = [Row(2, upper), Row(3, full)];
        for (var i = 0; i < EntityHierarchyCatalog.MaxChildrenPerParent; i++)
        {
            rows.Add(Row(4 + i, NewId(), parentId: full));
        }

        rows.Add(Row(4 + EntityHierarchyCatalog.MaxChildrenPerParent, target));

        DemotePlan plan = HierarchyMutations.PlanDemoteToChild(rows, target);

        Assert.False(plan.Succeeded);
        Assert.Equal(ParentChildRefusalReason.ParentUnavailable, plan.Refusal);
        Assert.Empty(plan.SiblingOrderWrites);
    }

    /// <summary>
    /// Demoting a row that is already a child refuses with
    /// <c>AlreadyAChild</c>.
    /// </summary>
    [Fact]
    public void Demoting_a_child_is_refused_as_already_a_child()
    {
        GanttRowId parent = NewId();
        GanttRowId child = NewId();
        List<GanttEvent> rows = [Row(2, parent), Row(3, child, parentId: parent)];

        DemotePlan plan = HierarchyMutations.PlanDemoteToChild(rows, child);

        Assert.False(plan.Succeeded);
        Assert.Equal(ParentChildRefusalReason.AlreadyAChild, plan.Refusal);
    }

    /// <summary>
    /// A standalone row (Splitter) can never be demoted.
    /// </summary>
    [Fact]
    public void Demoting_a_splitter_is_refused_as_wrong_type()
    {
        GanttRowId parent = NewId();
        GanttRowId splitter = NewId();
        List<GanttEvent> rows =
        [
            Row(2, parent),
            Row(3, splitter, GanttEntityType.Splitter),
        ];

        DemotePlan plan = HierarchyMutations.PlanDemoteToChild(rows, splitter);

        Assert.False(plan.Succeeded);
        Assert.Equal(ParentChildRefusalReason.TypeCannotBeChild, plan.Refusal);
    }

    /// <summary>
    /// With no eligible row above, demote refuses with
    /// <c>NoEligibleParentAbove</c> and moves nothing.
    /// </summary>
    [Fact]
    public void Demoting_with_nothing_eligible_above_is_refused()
    {
        GanttRowId line = NewId();
        GanttRowId target = NewId();
        List<GanttEvent> rows =
        [
            Row(2, line, GanttEntityType.Delineator),
            Row(3, target),
        ];

        DemotePlan plan = HierarchyMutations.PlanDemoteToChild(rows, target);

        Assert.False(plan.Succeeded);
        Assert.Equal(ParentChildRefusalReason.NoEligibleParentAbove, plan.Refusal);
        Assert.Empty(plan.SiblingOrderWrites);
    }

    /// <summary>
    /// A milestone between the parent and the selection is passed over by
    /// the scan -- and the demotion MOVES to sit directly beneath the
    /// parent's run, so it stays one contiguous outline group (R5.3 D3,
    /// amended 2026-10-11).
    /// </summary>
    [Fact]
    public void Demoting_beneath_a_milestone_moves_beneath_the_parent_run()
    {
        GanttRowId parent = NewId();
        GanttRowId milestone = NewId();
        GanttRowId target = NewId();
        List<GanttEvent> rows =
        [
            Row(2, parent),
            Row(3, milestone, GanttEntityType.AsPlannedMilestone),
            Row(4, target),
        ];

        DemotePlan plan = HierarchyMutations.PlanDemoteToChild(rows, target);

        Assert.True(plan.Succeeded);
        Assert.Equal(target, plan.TargetId);
        Assert.Equal(parent, plan.NewParentId);
        Assert.Equal(3, plan.MoveToBodyIndex);
        Assert.Empty(plan.RehomedChildIds);
        Assert.Contains(new SiblingOrderWrite(target, 0), plan.SiblingOrderWrites);
    }

    /// <summary>
    /// Planner input is never mutated by either new plan.
    /// </summary>
    [Fact]
    public void Parent_child_planning_does_not_mutate_its_input()
    {
        GanttRowId parent = NewId();
        GanttRowId child = NewId();
        List<GanttEvent> rows = [Row(2, parent), Row(3, child, parentId: parent)];
        List<GanttEvent> snapshot = [.. rows];

        _ = HierarchyMutations.PlanPromoteToParent(rows, child);
        _ = HierarchyMutations.PlanDemoteToChild(rows, parent);

        Assert.Equal(snapshot, rows);
    }
}
