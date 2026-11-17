namespace GanttCreator.Core.Tests;

/// <summary>
/// Behavioural pins for <see cref="HierarchyMutations"/> plans.
/// </summary>
/// <remarks>
/// <para>
/// Mutation testing left 74 mutants open in this file: 20 boolean, 16 statement,
/// 12 equality, 6 post-increment-to-decrement, 5 block-removal, 3
/// <c>OrderBy</c>-to-<c>OrderByDescending</c>, plus conditional and arithmetic
/// mutants. Every one of them is observable in the returned plan, so these tests
/// assert the plan's <em>exact</em> fields and write sequence rather than
/// "something happened". A plan that swaps the wrong window, resequences with a
/// decrementing counter, or picks the wrong neighbour changes the user's row
/// order, which is the whole product surface of these commands.
/// </para>
/// <para>
/// The resequencing assertions pin the literal order values
/// (<c>0, 1, 2 …</c>) because "dense ascending" is the contract: the planner is
/// the authority and existing cell values are ignored (R5.3 D4).
/// </para>
/// </remarks>
public class HierarchyMutationsBehaviourTests
{
    private static GanttRowId Rid(char digit)
    {
        Assert.True(
            GanttRowId.TryParse($"G-{new string(digit, 32)}", out GanttRowId? parsed),
            "test fixture id must be a well-formed row id");
        return parsed!;
    }

    private static GanttEvent Row(int rowNumber, char id, GanttEntityType type, char? parentId = null) =>
        new(
            rowNumber,
            Rid(id),
            null,
            null,
            type,
            null,
            new DateOnly(2026, 9, 1),
            new DateOnly(2026, 9, 5),
            parentId is { } parent ? Rid(parent) : null,
            null,
            null,
            null,
            null,
            true,
            null);

    private static GanttEvent Span(int rowNumber, char id, char? parentId = null) =>
        Row(rowNumber, id, GanttEntityType.AsPlannedActivity, parentId);

    private static (GanttRowId RowId, int SiblingOrder)[] Writes(MovePlan plan) =>
        [.. plan.SiblingOrderWrites.Select(write => (write.RowId, write.SiblingOrder))];

    private static (GanttRowId RowId, int SiblingOrder)[] Writes(PromotePlan plan) =>
        [.. plan.SiblingOrderWrites.Select(write => (write.RowId, write.SiblingOrder))];

    // ---------------------------------------------------------------- moves

    [Fact]
    public void A_child_move_up_swaps_with_the_row_above_in_its_own_run()
    {
        GanttRowId a = Rid('1');
        GanttRowId b = Rid('2');
        GanttRowId c = Rid('3');
        GanttEvent[] rows = [Span(1, '1'), Span(2, '2', '1'), Span(3, '3', '1')];

        MovePlan plan = HierarchyMutations.PlanMove(rows, c, moveUp: true);

        Assert.True(plan.Succeeded);
        Assert.True(plan.MovesAnything);
        Assert.Equal(2, plan.WindowStart);
        Assert.Equal(1, plan.WindowLength);
        Assert.Equal(1, plan.SwapStart);
        Assert.Equal(1, plan.SwapLength);
        Assert.Equal([(c, 0), (b, 1)], Writes(plan));
        _ = a;
    }

    [Fact]
    public void A_child_move_down_swaps_with_the_row_below_in_its_own_run()
    {
        GanttRowId b = Rid('2');
        GanttRowId c = Rid('3');
        GanttEvent[] rows = [Span(1, '1'), Span(2, '2', '1'), Span(3, '3', '1')];

        MovePlan plan = HierarchyMutations.PlanMove(rows, b, moveUp: false);

        Assert.True(plan.MovesAnything);
        Assert.Equal(1, plan.WindowStart);
        Assert.Equal(1, plan.WindowLength);
        Assert.Equal(2, plan.SwapStart);
        Assert.Equal(1, plan.SwapLength);
        Assert.Equal([(c, 0), (b, 1)], Writes(plan));
    }

    [Fact]
    public void A_child_at_the_top_of_its_run_is_a_silent_no_op()
    {
        // Its parent is the neighbour, which is a run boundary, not a sibling.
        GanttRowId b = Rid('2');
        GanttEvent[] rows = [Span(1, '1'), Span(2, '2', '1'), Span(3, '3', '1')];

        MovePlan plan = HierarchyMutations.PlanMove(rows, b, moveUp: true);

        Assert.True(plan.Succeeded);
        Assert.False(plan.MovesAnything);
        Assert.Equal(0, plan.WindowLength);
        Assert.Empty(plan.SiblingOrderWrites);
    }

    [Fact]
    public void A_child_at_the_bottom_of_its_run_is_a_silent_no_op()
    {
        GanttRowId c = Rid('3');
        GanttEvent[] rows = [Span(1, '1'), Span(2, '2', '1'), Span(3, '3', '1')];

        MovePlan plan = HierarchyMutations.PlanMove(rows, c, moveUp: false);

        Assert.True(plan.Succeeded);
        Assert.False(plan.MovesAnything);
        Assert.Empty(plan.SiblingOrderWrites);
    }

    [Fact]
    public void A_top_level_move_down_swaps_whole_groups_and_resequences_the_run()
    {
        // A(1) with child A1(2), then B(3) with child B1(4). Moving A down
        // swaps both two-row windows, not just the header rows.
        GanttRowId a = Rid('1');
        GanttRowId b = Rid('3');
        GanttEvent[] rows = [Span(1, '1'), Span(2, '2', '1'), Span(3, '3'), Span(4, '4', '3')];

        MovePlan plan = HierarchyMutations.PlanMove(rows, a, moveUp: false);

        Assert.True(plan.MovesAnything);
        Assert.Equal(0, plan.WindowStart);
        Assert.Equal(2, plan.WindowLength);
        Assert.Equal(2, plan.SwapStart);
        Assert.Equal(2, plan.SwapLength);
        Assert.Equal([(b, 0), (a, 1)], Writes(plan));
    }

    [Fact]
    public void A_top_level_move_up_swaps_whole_groups_and_resequences_the_run()
    {
        GanttRowId a = Rid('1');
        GanttRowId b = Rid('3');
        GanttEvent[] rows = [Span(1, '1'), Span(2, '2', '1'), Span(3, '3'), Span(4, '4', '3')];

        MovePlan plan = HierarchyMutations.PlanMove(rows, b, moveUp: true);

        Assert.True(plan.MovesAnything);
        Assert.Equal(2, plan.WindowStart);
        Assert.Equal(2, plan.WindowLength);
        Assert.Equal(0, plan.SwapStart);
        Assert.Equal(2, plan.SwapLength);
        Assert.Equal([(b, 0), (a, 1)], Writes(plan));
    }

    [Fact]
    public void A_top_level_row_at_the_top_of_the_body_is_a_silent_no_op()
    {
        GanttEvent[] rows = [Span(1, '1'), Span(2, '2', '1'), Span(3, '3'), Span(4, '4', '3')];

        MovePlan plan = HierarchyMutations.PlanMove(rows, Rid('1'), moveUp: true);

        Assert.True(plan.Succeeded);
        Assert.False(plan.MovesAnything);
        Assert.Empty(plan.SiblingOrderWrites);
    }

    [Fact]
    public void A_top_level_group_at_the_bottom_of_the_body_is_a_silent_no_op()
    {
        GanttEvent[] rows = [Span(1, '1'), Span(2, '2', '1'), Span(3, '3'), Span(4, '4', '3')];

        MovePlan plan = HierarchyMutations.PlanMove(rows, Rid('3'), moveUp: false);

        Assert.True(plan.Succeeded);
        Assert.False(plan.MovesAnything);
        Assert.Empty(plan.SiblingOrderWrites);
    }

    [Fact]
    public void An_unknown_move_selection_is_a_typed_refusal()
    {
        GanttEvent[] rows = [Span(1, '1')];

        MovePlan plan = HierarchyMutations.PlanMove(rows, Rid('9'), moveUp: true);

        Assert.False(plan.Succeeded);
        Assert.Equal(ParentChildRefusalReason.SelectionUnknown, plan.Refusal);
        Assert.Empty(plan.SiblingOrderWrites);
    }

    // -------------------------------------------------------------- promote

    [Fact]
    public void Promoting_a_middle_child_resequences_the_remaining_run_densely()
    {
        // The remaining siblings take 0 then 1, and the promoted row joins the
        // top-level run at the bottom of its former run. A decrementing counter
        // would emit 0 then -1, which is the mutant class being pinned here.
        GanttRowId c1 = Rid('2');
        GanttRowId c2 = Rid('3');
        GanttRowId c3 = Rid('4');
        GanttRowId parent = Rid('1');
        GanttEvent[] rows = [Span(1, '1'), Span(2, '2', '1'), Span(3, '3', '1'), Span(4, '4', '1')];

        PromotePlan plan = HierarchyMutations.PlanPromoteToParent(rows, c2);

        Assert.True(plan.Succeeded);
        Assert.Equal(c2, plan.TargetId);
        Assert.Equal(parent, plan.ClearedParentId);
        Assert.Equal(4, plan.MoveToBodyIndex);
        Assert.Equal([(c1, 0), (c3, 1), (parent, 0), (c2, 1)], Writes(plan));
    }

    [Fact]
    public void Promotion_inherits_the_parent_of_the_first_row_carrying_the_deleted_id()
    {
        // A duplicated id is already a validation error, so this only pins the
        // defensive first-wins policy: the inherited parent must come from the
        // first match, never the last.
        GanttRowId deleted = Rid('5');
        GanttRowId firstParent = Rid('1');
        GanttRowId secondParent = Rid('2');
        GanttEvent[] rows =
        [
            Span(1, '1'),
            Span(2, '2'),
            Span(3, '5', '1'),
            Span(4, '6', '5'),
            Span(5, '5', '2'),
        ];

        HierarchyPromotionPlan plan = HierarchyMutations.PlanPromotion(rows, deleted);

        Assert.Equal(firstParent, plan.InheritedParentId);
        Assert.Equal([Rid('6')], plan.PromotedChildren);
        Assert.NotEqual(secondParent, plan.InheritedParentId);
        Assert.True(HierarchyMutations.RequiresParentRewrite(plan));
    }

    // --------------------------------------------------------------- demote

    [Fact]
    public void A_demotion_that_lands_exactly_on_the_child_cap_is_allowed()
    {
        // Six existing children plus the joining row is exactly the maximum, so
        // the boundary must succeed. A ">=" comparison would refuse it.
        int existing = EntityHierarchyCatalog.MaxChildrenPerParent - 1;
        GanttRowId parent = Rid('1');
        GanttRowId selection = Rid('9');

        List<GanttEvent> rows = [Span(1, '1')];
        for (int i = 0; i < existing; i++)
        {
            rows.Add(Span(2 + i, (char)('2' + i), '1'));
        }

        rows.Add(Span(2 + existing, '9'));

        DemotePlan plan = HierarchyMutations.PlanDemoteToChild(rows, selection);

        Assert.True(plan.Succeeded);
        Assert.Equal(parent, plan.NewParentId);
        Assert.Equal(2 + existing, plan.MoveToBodyIndex);
        Assert.Equal((selection, existing), (plan.SiblingOrderWrites[0].RowId, plan.SiblingOrderWrites[0].SiblingOrder));
    }

    [Fact]
    public void A_demotion_that_would_exceed_the_child_cap_is_refused()
    {
        int existing = EntityHierarchyCatalog.MaxChildrenPerParent;
        GanttRowId selection = Rid('9');

        List<GanttEvent> rows = [Span(1, '1')];
        for (int i = 0; i < existing; i++)
        {
            rows.Add(Span(2 + i, (char)('2' + i), '1'));
        }

        rows.Add(Span(2 + existing, '9'));

        DemotePlan plan = HierarchyMutations.PlanDemoteToChild(rows, selection);

        Assert.False(plan.Succeeded);
        Assert.Equal(ParentChildRefusalReason.ParentUnavailable, plan.Refusal);
        Assert.Empty(plan.SiblingOrderWrites);
        Assert.Empty(plan.RehomedChildIds);
    }

    [Fact]
    public void A_demotion_that_moves_the_selection_and_its_children_resequences_both_runs()
    {
        // P(1) owns C(2). S(3) is top-level and owns S1(4). Demoting S lands it
        // beneath P's run at row 3, so S1 re-homes to P and both runs resequence.
        GanttRowId parent = Rid('1');
        GanttRowId firstChild = Rid('2');
        GanttRowId selection = Rid('3');
        GanttRowId rehomed = Rid('4');
        GanttEvent[] rows = [Span(1, '1'), Span(2, '2', '1'), Span(3, '3'), Span(4, '4', '3')];

        DemotePlan plan = HierarchyMutations.PlanDemoteToChild(rows, selection);

        Assert.True(plan.Succeeded);
        Assert.Equal(parent, plan.NewParentId);
        Assert.Equal(3, plan.MoveToBodyIndex);
        Assert.Equal([rehomed], plan.RehomedChildIds);

        // The selection takes the tail slot after the parent's existing child
        // (order 0), the re-homed child follows it, and the top-level run it
        // leaves resequences densely. The existing child keeps order 0, so it
        // is deliberately not written.
        Assert.Equal((selection, 1), (plan.SiblingOrderWrites[0].RowId, plan.SiblingOrderWrites[0].SiblingOrder));
        Assert.Equal((rehomed, 2), (plan.SiblingOrderWrites[1].RowId, plan.SiblingOrderWrites[1].SiblingOrder));
        Assert.Equal((parent, 0), (plan.SiblingOrderWrites[2].RowId, plan.SiblingOrderWrites[2].SiblingOrder));
        Assert.DoesNotContain(plan.SiblingOrderWrites, write => write.RowId == firstChild);
    }

    // -------------------------------------------------------- argument guards

    [Fact]
    public void Every_planner_rejects_a_null_row_list()
    {
        _ = Assert.Throws<ArgumentNullException>(() => HierarchyMutations.PlanPromotion(null!, Rid('1')));
        _ = Assert.Throws<ArgumentNullException>(() => HierarchyMutations.PlanPromoteToParent(null!, Rid('1')));
        _ = Assert.Throws<ArgumentNullException>(() => HierarchyMutations.PlanDemoteToChild(null!, Rid('1')));
        _ = Assert.Throws<ArgumentNullException>(() => HierarchyMutations.PlanMove(null!, Rid('1'), moveUp: true));
    }

    [Fact]
    public void Every_planner_rejects_a_null_selection()
    {
        GanttEvent[] rows = [Span(1, '1')];

        _ = Assert.Throws<ArgumentNullException>(() => HierarchyMutations.PlanPromotion(rows, null!));
        _ = Assert.Throws<ArgumentNullException>(() => HierarchyMutations.PlanPromoteToParent(rows, null!));
        _ = Assert.Throws<ArgumentNullException>(() => HierarchyMutations.PlanDemoteToChild(rows, null!));
        _ = Assert.Throws<ArgumentNullException>(() => HierarchyMutations.PlanMove(rows, null!, moveUp: true));
    }

    [Fact]
    public void Requires_parent_rewrite_rejects_a_null_plan() =>
        _ = Assert.Throws<ArgumentNullException>(() => HierarchyMutations.RequiresParentRewrite(null!));
}
