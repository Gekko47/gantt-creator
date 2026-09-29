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
}
