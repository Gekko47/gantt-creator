namespace GanttCreator.Core.Tests;

/// <summary>
/// Tests for <see cref="CloneReseedPlanner"/>, the atomic identity reseed of a
/// duplicated managed structure (R4.7A D7).
/// </summary>
public sealed class CloneReseedPlannerTests
{
    private static GanttRowId NewId() => GanttRowId.New();

    /// <summary>
    /// A deterministic generator, so a plan is assertable. Real callers pass
    /// <see cref="GanttRowId.New"/>.
    /// </summary>
    private static Func<GanttRowId> Sequential(params GanttRowId[] ids)
    {
        var index = 0;
        return () => index < ids.Length ? ids[index++] : NewId();
    }

    /// <summary>
    /// A duplicated worksheet is the ordinary case: the same identifier appears
    /// twice, the first occurrence keeps it, and the clone is given a new one.
    /// </summary>
    [Fact]
    public void A_duplicated_id_is_reseeded_on_the_clone_only()
    {
        GanttRowId original = NewId();
        GanttRowId replacement = NewId();

        CloneReseedOutcome outcome = CloneReseedPlanner.Plan([original, original], Sequential(replacement));

        Assert.True(outcome.Succeeded);
        CloneReseedPlan plan = outcome.Plan!;
        Assert.Equal(1, plan.DuplicatedRowCount);
        Assert.Equal(replacement, plan.NewIdsByOldId[original]);
    }

    /// <summary>
    /// A cloned parent/child pair keeps its relationship under the new identities:
    /// the child's rewritten ParentId is the parent's NEW id, not the old one.
    /// This is the property that makes a clone usable rather than merely unique.
    /// </summary>
    [Fact]
    public void A_cloned_parent_child_pair_keeps_its_relationship()
    {
        GanttRowId parent = NewId();
        GanttRowId child = NewId();
        GanttRowId newParent = NewId();
        GanttRowId newChild = NewId();

        // Worksheet order: original parent, original child, cloned parent, cloned child.
        CloneReseedOutcome outcome = CloneReseedPlanner.Plan(
            [parent, child, parent, child],
            Sequential(newParent, newChild));

        Assert.True(outcome.Succeeded);
        CloneReseedPlan plan = outcome.Plan!;

        Assert.Equal(newParent, plan.NewIdsByOldId[parent]);
        Assert.Equal(newChild, plan.NewIdsByOldId[child]);

        // The cloned child's parent reference is translated to the cloned parent's
        // new id, so the clone is internally consistent.
        Assert.Equal(newParent, plan.TranslateParentId(parent));
    }

    /// <summary>
    /// A first-canonical row that is not duplicated keeps its identity: only the
    /// clone is rewritten. A reseed that replaced every id would break references
    /// from outside the structure.
    /// </summary>
    [Fact]
    public void A_unique_id_is_not_reseeded()
    {
        GanttRowId duplicated = NewId();
        GanttRowId unique = NewId();
        GanttRowId replacement = NewId();

        CloneReseedOutcome outcome = CloneReseedPlanner.Plan([unique, duplicated, duplicated], Sequential(replacement));

        Assert.True(outcome.Succeeded);
        Assert.Equal([duplicated], outcome.Plan!.NewIdsByOldId.Keys);
    }

    /// <summary>
    /// Nothing to do is an explicit, typed refusal rather than an empty plan, so a
    /// caller cannot mistake "no clone" for "clone reseeded successfully".
    /// </summary>
    [Fact]
    public void No_duplication_refuses_rather_than_returning_an_empty_plan()
    {
        CloneReseedOutcome outcome = CloneReseedPlanner.Plan([NewId(), NewId()], NewId);

        Assert.False(outcome.Succeeded);
        Assert.Equal(CloneReseedRefusal.NoDuplicatedIds, outcome.Refusal);
    }

    /// <summary>
    /// An identifier on three rows cannot be attributed to a single clone, so the
    /// whole reseed is refused. Rewriting a guess would corrupt a relationship the
    /// user never created -- and a partial reseed is worse than none.
    /// </summary>
    [Fact]
    public void An_identifier_on_three_rows_refuses_the_whole_reseed()
    {
        GanttRowId id = NewId();

        CloneReseedOutcome outcome = CloneReseedPlanner.Plan([id, id, id], NewId);

        Assert.False(outcome.Succeeded);
        Assert.Equal(CloneReseedRefusal.TooManyDuplicates, outcome.Refusal);
    }

    /// <summary>
    /// Refusal is all-or-nothing even when other identifiers in the same table are
    /// cleanly reseedable. The atomicity D7 requires is that one unresolvable id
    /// stops the entire plan.
    /// </summary>
    [Fact]
    public void One_unresolvable_id_refuses_a_table_with_other_clean_pairs()
    {
        GanttRowId clean = NewId();
        GanttRowId overDuplicated = NewId();

        CloneReseedOutcome outcome = CloneReseedPlanner.Plan(
            [clean, clean, overDuplicated, overDuplicated, overDuplicated],
            NewId);

        Assert.False(outcome.Succeeded);
        Assert.Equal(CloneReseedRefusal.TooManyDuplicates, outcome.Refusal);
        Assert.Null(outcome.Plan);
    }

    /// <summary>
    /// A parent reference to a row that was not reseeded is left alone rather than
    /// pointed at an arbitrary clone.
    /// </summary>
    [Fact]
    public void A_parent_reference_outside_the_reseed_is_left_alone()
    {
        GanttRowId outside = NewId();
        GanttRowId duplicated = NewId();
        GanttRowId replacement = NewId();

        CloneReseedOutcome outcome = CloneReseedPlanner.Plan([outside, duplicated, duplicated], Sequential(replacement));

        Assert.True(outcome.Succeeded);
        Assert.Null(outcome.Plan!.TranslateParentId(outside));
    }

    /// <summary>
    /// A blank parent reference translates to blank, which is what makes a
    /// top-level row stay top-level through a reseed.
    /// </summary>
    [Fact]
    public void A_blank_parent_reference_stays_blank()
    {
        GanttRowId duplicated = NewId();

        CloneReseedOutcome outcome = CloneReseedPlanner.Plan([duplicated, duplicated], Sequential(NewId()));

        Assert.True(outcome.Succeeded);
        Assert.Null(outcome.Plan!.TranslateParentId(null));
    }

    /// <summary>
    /// A generator that repeats an id already in the table is retried, so a
    /// reseed can never introduce the very collision it is removing.
    /// </summary>
    [Fact]
    public void A_generator_repeating_an_existing_id_is_retried()
    {
        GanttRowId existing = NewId();
        GanttRowId duplicated = NewId();
        GanttRowId fresh = NewId();

        var sequence = new Queue<GanttRowId>([existing, fresh]);
        CloneReseedOutcome outcome = CloneReseedPlanner.Plan([existing, duplicated, duplicated], () => sequence.Dequeue());

        Assert.True(outcome.Succeeded);
        Assert.Equal(fresh, outcome.Plan!.NewIdsByOldId[duplicated]);
    }

    /// <summary>
    /// Determinism: the plan is a pure function of the rows and the generator, so
    /// the same input always yields the same map.
    /// </summary>
    [Fact]
    public void Planning_is_deterministic_for_the_same_input()
    {
        GanttRowId a = NewId();
        GanttRowId b = NewId();
        GanttRowId newA = NewId();
        GanttRowId newB = NewId();

        CloneReseedOutcome first = CloneReseedPlanner.Plan([a, b, a, b], Sequential(newA, newB));
        CloneReseedOutcome second = CloneReseedPlanner.Plan([a, b, a, b], Sequential(newA, newB));

        Assert.Equal(first.Plan!.NewIdsByOldId, second.Plan!.NewIdsByOldId);
    }

    [Fact]
    public void Null_arguments_are_rejected()
    {
        Assert.Throws<ArgumentNullException>(() => CloneReseedPlanner.Plan(null!, NewId));
        Assert.Throws<ArgumentNullException>(() => CloneReseedPlanner.Plan([], null!));
    }
}
