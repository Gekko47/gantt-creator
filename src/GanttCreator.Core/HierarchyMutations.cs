namespace GanttCreator.Core;

/// <summary>
/// The result of planning a promote-on-delete: which children's <c>ParentId</c>
/// values change, and to what.
/// </summary>
/// <param name="DeletedId">The parent that was deleted.</param>
/// <param name="PromotedChildren">
/// The promoted children in ascending worksheet row order. Empty when the deleted
/// row owned no children.
/// </param>
/// <param name="InheritedParentId">
/// The parent the promoted children inherit, or <see langword="null"/> when they
/// become top-level. This is the deleted row's own parent, so the whole plan is
/// one value rather than a per-child decision.
/// </param>
public sealed record HierarchyPromotionPlan(
    GanttRowId DeletedId,
    IReadOnlyList<GanttRowId> PromotedChildren,
    GanttRowId? InheritedParentId)
{
    /// <summary>Gets whether the delete promoted anything at all.</summary>
    public bool PromotedAnything => PromotedChildren.Count > 0;
}

/// <summary>
/// One <c>SiblingOrder</c> cell rewrite within a single sibling run.
/// </summary>
/// <param name="RowId">The row whose <c>SiblingOrder</c> cell is rewritten.</param>
/// <param name="SiblingOrder">The dense order value (<c>0..n</c>) the cell takes.</param>
public sealed record SiblingOrderWrite(GanttRowId RowId, int SiblingOrder);

/// <summary>
/// Why a promote or demote plan was refused. Every refusal leaves the
/// workbook untouched.
/// </summary>
public enum ParentChildRefusalReason
{
    /// <summary>The selected identifier names no row of the table.</summary>
    SelectionUnknown = 0,

    /// <summary>Promote was asked of a row that is not a child (blank <c>ParentId</c>).</summary>
    NotAChild = 1,

    /// <summary>Demote was asked of a row that is already a child.</summary>
    AlreadyAChild = 2,

    /// <summary>The row's type may never be a child (standalone or leaf-parent).</summary>
    TypeCannotBeChild = 3,

    /// <summary>A <c>CriticalInterval</c> was asked to promote; it is child-only by contract.</summary>
    PromotionForbiddenForType = 4,

    /// <summary>The promoted row owns children of its own; promotion would orphan them.</summary>
    HasChildren = 5,

    /// <summary>No eligible parent-capable row exists above the selection.</summary>
    NoEligibleParentAbove = 6,

    /// <summary>The only eligible parent above refused the child (capacity exhausted up the table).</summary>
    ParentUnavailable = 7,

    /// <summary>The child would nest three deep (the candidate parent is itself a child).</summary>
    TooDeep = 8,

    /// <summary>The move would create a parent cycle.</summary>
    Cycle = 9,

    /// <summary>A <c>CriticalInterval</c> child would land on a non-span parent.</summary>
    ParentNotSpan = 10,

    /// <summary>
    /// The selection is separated from the nearest eligible parent's
    /// child run by rows that are not its children, so the demotion
    /// could not be drawn as one contiguous outline group.
    /// </summary>
    LandingNotAdjacent = 11,
}

/// <summary>
/// The result of planning one child-to-parent promotion (R5.3 D2): which row
/// moves to the bottom of its sibling run, which <c>ParentId</c> is cleared,
/// and which <c>SiblingOrder</c> cells close the gap.
/// </summary>
/// <param name="TargetId">The promoted row.</param>
/// <param name="ClearedParentId">The parent the row leaves.</param>
/// <param name="MoveToBodyIndex">
/// The one-based body index the row's content moves to (the bottom of its
/// former sibling run; equal to its current index when already last).
/// </param>
/// <param name="SiblingOrderWrites">
/// Dense resequencing of the remaining run, ascending row order. Empty when
/// the promoted row was its parent's only child.
/// </param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/>.</param>
public sealed record PromotePlan(
    GanttRowId? TargetId,
    GanttRowId? ClearedParentId,
    int? MoveToBodyIndex,
    IReadOnlyList<SiblingOrderWrite> SiblingOrderWrites,
    ParentChildRefusalReason? Refusal)
{
    /// <summary>Gets whether planning succeeded.</summary>
    public bool Succeeded => Refusal is null;

    /// <summary>Creates a refusal plan. Nothing is mutated.</summary>
    /// <param name="refusal">Why no plan was produced.</param>
    /// <returns>The refusal plan.</returns>
    public static PromotePlan Refused(ParentChildRefusalReason refusal) =>
        new(null, null, null, [], refusal);
}

/// <summary>
/// The result of planning one parent-to-child demotion (R5.3 D3): which row
/// gains which <c>ParentId</c>, and which <c>SiblingOrder</c> cells resequence.
/// </summary>
/// <param name="TargetId">The demoted row.</param>
/// <param name="NewParentId">The nearest eligible parent above the selection.</param>
/// <param name="SiblingOrderWrites">
/// The demoted row's tail assignment in its new run, plus the dense
/// resequencing of the top-level run it leaves.
/// </param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/>.</param>
public sealed record DemotePlan(
    GanttRowId? TargetId,
    GanttRowId? NewParentId,
    IReadOnlyList<SiblingOrderWrite> SiblingOrderWrites,
    ParentChildRefusalReason? Refusal)
{
    /// <summary>Gets whether planning succeeded.</summary>
    public bool Succeeded => Refusal is null;

    /// <summary>Creates a refusal plan. Nothing is mutated.</summary>
    /// <param name="refusal">Why no plan was produced.</param>
    /// <returns>The refusal plan.</returns>
    public static DemotePlan Refused(ParentChildRefusalReason refusal) =>
        new(null, null, [], refusal);
}

/// <summary>
/// Pure, deterministic hierarchy mutations that do not touch Excel. Every result
/// is a <em>plan</em>: the caller applies it. Nothing here writes a cell, reads a
/// workbook, or mutates its input.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why a plan and not an action.</b> Promote-on-delete rewrites
/// <c>ParentId</c> for several rows at once, and a failure part-way through would
/// leave a half-promoted hierarchy that is worse than an unpromoted one. A plan is
/// computed whole and can therefore be refused whole, and applying it is the
/// Office layer's single responsibility.
/// </para>
/// <para>
/// <b>Determinism.</b> Children are returned in ascending worksheet row order, so
/// the same table always produces the same plan and the same writes.
/// </para>
/// </remarks>
public static class HierarchyMutations
{
    /// <summary>
    /// Plans the promotion of a deleted row's children. A top-level parent is
    /// deleted and its children become top-level; a nested parent is deleted and
    /// its children inherit the deleted row's own parent. Promoted children keep
    /// their own IDs and schedule data -- only <c>ParentId</c> changes, and only
    /// when the inherited parent differs from what the child already had.
    /// </summary>
    /// <param name="rows">The validated rows, in any order. Not mutated.</param>
    /// <param name="deletedId">The identifier of the row being deleted.</param>
    /// <returns>The promotion plan, whose <c>PromotedChildren</c> is empty when the row owned none.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="rows"/> is null.</exception>
    public static HierarchyPromotionPlan PlanPromotion(
        IReadOnlyList<GanttEvent> rows,
        GanttRowId deletedId)
    {
        ArgumentNullException.ThrowIfNull(rows);
        ArgumentNullException.ThrowIfNull(deletedId);

        GanttRowId? inherited = null;
        foreach (GanttEvent row in rows)
        {
            if (row.Id == deletedId)
            {
                inherited = row.ParentId;
                break;
            }
        }

        List<GanttRowId> promoted = [];
        foreach (GanttEvent row in rows.OrderBy(static r => r.RowNumber))
        {
            if (row.Id != deletedId && row.ParentId == deletedId)
            {
                promoted.Add(row.Id);
            }
        }

        return new HierarchyPromotionPlan(deletedId, [.. promoted], inherited);
    }

    /// <summary>
    /// Whether a child's <c>ParentId</c> must be rewritten after its parent was
    /// deleted. A child that would inherit the same parent it already has -- the
    /// case where the deleted row was itself top-level, so its children become
    /// top-level with a blank <c>ParentId</c> -- still needs its cell cleared, so
    /// this returns <see langword="true"/> for any surviving child of the deleted
    /// row.
    /// </summary>
    /// <param name="plan">The promotion plan.</param>
    /// <returns><see langword="true"/> when at least one child's ParentId changes.</returns>
    public static bool RequiresParentRewrite(HierarchyPromotionPlan plan)
    {
        ArgumentNullException.ThrowIfNull(plan);
        return plan.PromotedAnything;
    }

    /// <summary>
    /// Plans one child-to-parent promotion (R5.3 D2). The target must be a
    /// child of a promotable type that owns no children of its own; its table
    /// content moves to the bottom of its former sibling run, its
    /// <c>ParentId</c> is cleared, the remaining run is resequenced
    /// densely, and the promoted row joins the top-level run at its new
    /// position with the run resequenced around it (R5.3 D4). Nothing is
    /// mutated; the caller applies the plan.
    /// </summary>
    /// <param name="rows">The validated rows, in any order. Not mutated.</param>
    /// <param name="selectedId">The identifier of the row being promoted.</param>
    /// <returns>The promotion plan, or a typed refusal that mutates nothing.</returns>
    /// <exception cref="ArgumentNullException">Thrown when an argument is null.</exception>
    public static PromotePlan PlanPromoteToParent(
        IReadOnlyList<GanttEvent> rows,
        GanttRowId selectedId)
    {
        ArgumentNullException.ThrowIfNull(rows);
        ArgumentNullException.ThrowIfNull(selectedId);

        Dictionary<GanttRowId, GanttEvent> byId = [];
        foreach (GanttEvent row in rows)
        {
            // First-wins mirrors the validator's first-canonical policy: a
            // duplicated id is already a validation error, so last-wins would
            // be arbitrary.
            _ = byId.TryAdd(row.Id, row);
        }

        if (!byId.TryGetValue(selectedId, out GanttEvent? target))
        {
            return PromotePlan.Refused(ParentChildRefusalReason.SelectionUnknown);
        }

        if (target.ParentId is null)
        {
            return PromotePlan.Refused(ParentChildRefusalReason.NotAChild);
        }

        if (!EntityHierarchyCatalog.MayBeChild(target.Type))
        {
            return PromotePlan.Refused(ParentChildRefusalReason.TypeCannotBeChild);
        }

        if (target.Type == GanttEntityType.CriticalInterval)
        {
            return PromotePlan.Refused(ParentChildRefusalReason.PromotionForbiddenForType);
        }

        foreach (GanttEvent row in rows)
        {
            if (row.Id != selectedId && row.ParentId == selectedId)
            {
                return PromotePlan.Refused(ParentChildRefusalReason.HasChildren);
            }
        }

        GanttRowId parentId = target.ParentId;
        List<GanttEvent> siblings = [.. rows
            .Where(row => row.ParentId == parentId)
            .OrderBy(static row => row.RowNumber)];

        // The move target is the bottom of the run: the highest body index
        // among the siblings, which is the promoted row's own index when it
        // is already last.
        var moveTo = siblings.Count == 0
            ? target.RowNumber
            : siblings[^1].RowNumber;

        List<SiblingOrderWrite> writes = [];
        var order = 0;
        foreach (GanttEvent sibling in siblings)
        {
            if (sibling.Id == selectedId)
            {
                continue;
            }

            writes.Add(new SiblingOrderWrite(sibling.Id, order));
            order++;
        }

        // The promoted row joins the top-level run at its
        // new position -- the bottom of its former sibling
        // run -- and the run resequences densely around it,
        // exactly as a demote resequences the run a row
        // leaves. The promoted row's stale child-run value
        // is never carried into the top-level run: the
        // planner is the authority (R5.3 D4, D2 amended
        // 2026-10-11).
        writes.AddRange(ResequenceTopLevelRunWithJoin(rows, selectedId, moveTo));

        return new PromotePlan(selectedId, parentId, moveTo, writes, null);
    }

    /// <summary>
    /// Plans one parent-to-child demotion (R5.3 D3). The target must be a
    /// top-level row of a demotable type; it inherits the nearest eligible
    /// parent-capable row above it in body order, provided the selection
    /// lands directly beneath that parent's existing child run -- Excel's
    /// outline groups only the contiguous block directly beneath a parent,
    /// so a separated landing refuses with
    /// <see cref="ParentChildRefusalReason.LandingNotAdjacent"/>. Nothing
    /// is mutated; the caller applies the plan.
    /// </summary>
    /// <param name="rows">The validated rows, in any order. Not mutated.</param>
    /// <param name="selectedId">The identifier of the row being demoted.</param>
    /// <returns>The demotion plan, or a typed refusal that mutates nothing.</returns>
    /// <exception cref="ArgumentNullException">Thrown when an argument is null.</exception>
    public static DemotePlan PlanDemoteToChild(
        IReadOnlyList<GanttEvent> rows,
        GanttRowId selectedId)
    {
        ArgumentNullException.ThrowIfNull(rows);
        ArgumentNullException.ThrowIfNull(selectedId);

        Dictionary<GanttRowId, GanttEvent> byId = [];
        Dictionary<int, GanttEvent> byRowNumber = [];
        foreach (GanttEvent row in rows)
        {
            _ = byId.TryAdd(row.Id, row);
            _ = byRowNumber.TryAdd(row.RowNumber, row);
        }

        if (!byId.TryGetValue(selectedId, out GanttEvent? target))
        {
            return DemotePlan.Refused(ParentChildRefusalReason.SelectionUnknown);
        }

        if (target.ParentId is not null)
        {
            return DemotePlan.Refused(ParentChildRefusalReason.AlreadyAChild);
        }

        if (!EntityHierarchyCatalog.MayBeChild(target.Type))
        {
            return DemotePlan.Refused(ParentChildRefusalReason.TypeCannotBeChild);
        }

        Dictionary<GanttRowId, int> childCountByParent = [];
        foreach (GanttEvent row in rows)
        {
            if (row.ParentId is { } parent)
            {
                childCountByParent[parent] = childCountByParent.GetValueOrDefault(parent) + 1;
            }
        }

        List<GanttEvent> above = [.. rows
            .Where(row => row.RowNumber < target.RowNumber)
            .OrderByDescending(static row => row.RowNumber)];

        foreach (GanttEvent candidate in above)
        {
            if (candidate.Type == GanttEntityType.Delineator)
            {
                continue;
            }

            if (!EntityHierarchyCatalog.MayOwnChildren(candidate.Type))
            {
                continue;
            }

            if (candidate.ParentId is not null)
            {
                // A child cannot own: that would be depth 3.
                continue;
            }

            if (IsDescendantOf(byId, candidate.Id, selectedId))
            {
                return DemotePlan.Refused(ParentChildRefusalReason.Cycle);
            }

            if (!LandsDirectlyBeneath(byRowNumber, candidate, target))
            {
                // Excel's outline groups only the contiguous block
                // directly beneath a parent (OutlineGroupPlanner
                // anchors a group at parentRow + 1). A selection
                // separated from this candidate's run by rows that
                // are not its children would leave the demoted row
                // top-level on every Refresh despite its new
                // ParentId. No parent further up can be adjacent
                // either -- this candidate's own row is an
                // intervening top-level row for any of them -- so
                // the demotion refuses rather than writing a
                // hierarchy the worksheet cannot show (R5.3 D3).
                return DemotePlan.Refused(ParentChildRefusalReason.LandingNotAdjacent);
            }

            if (childCountByParent.GetValueOrDefault(candidate.Id) >= EntityHierarchyCatalog.MaxChildrenPerParent)
            {
                // The selection extends this parent's run, and no
                // parent further up can be adjacent, so the
                // capacity rule refuses instead of skipping to a
                // landing the outline cannot represent (R5.3 D3).
                return DemotePlan.Refused(ParentChildRefusalReason.ParentUnavailable);
            }

            if (target.Type == GanttEntityType.CriticalInterval && !IsSpan(candidate.Type))
            {
                return DemotePlan.Refused(ParentChildRefusalReason.ParentNotSpan);
            }

            var newRunSize = childCountByParent.GetValueOrDefault(candidate.Id);
            List<SiblingOrderWrite> writes =
            [
                new(target.Id, newRunSize),
                .. ResequenceTopLevelRun(rows, selectedId),
            ];

            return new DemotePlan(target.Id, candidate.Id, writes, null);
        }

        return DemotePlan.Refused(ParentChildRefusalReason.NoEligibleParentAbove);
    }

    /// <summary>
    /// Whether the selection lands directly beneath the candidate's
    /// existing child run: every row between them is already the
    /// candidate's child, so the selection extends the run and the
    /// demotion stays one contiguous outline group. A childless
    /// candidate qualifies only when it sits immediately above the
    /// selection.
    /// </summary>
    /// <param name="byRowNumber">The rows by body row number.</param>
    /// <param name="candidate">The parent the scan is considering.</param>
    /// <param name="target">The row being demoted.</param>
    /// <returns>Whether the demotion lands directly beneath the candidate.</returns>
    private static bool LandsDirectlyBeneath(
        Dictionary<int, GanttEvent> byRowNumber,
        GanttEvent candidate,
        GanttEvent target)
    {
        for (var rowNumber = candidate.RowNumber + 1; rowNumber < target.RowNumber; rowNumber++)
        {
            if (!byRowNumber.TryGetValue(rowNumber, out GanttEvent? between)
                || between.ParentId != candidate.Id)
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// Whether the parent chain of <paramref name="id"/> contains
    /// <paramref name="rootId"/>. Guards the demote scan against pathological
    /// input where the nearest-above row is already below the selection in
    /// the parent graph.
    /// </summary>
    private static bool IsDescendantOf(
        Dictionary<GanttRowId, GanttEvent> byId,
        GanttRowId id,
        GanttRowId rootId)
    {
        GanttRowId? cursor = byId.TryGetValue(id, out GanttEvent? current) ? current.ParentId : null;
        HashSet<GanttRowId> seen = [];
        while (cursor is not null)
        {
            if (cursor == rootId)
            {
                return true;
            }

            if (!seen.Add(cursor) || !byId.TryGetValue(cursor, out GanttEvent? next))
            {
                return false;
            }

            cursor = next.ParentId;
        }

        return false;
    }

    /// <summary>
    /// Dense resequencing of the top-level run after a row
    /// joins it at <paramref name="joinedRowNumber"/> (R5.3
    /// D2). The joining row takes its position in row order
    /// and the rows below it shift to make room. Existing
    /// <c>SiblingOrder</c> values are ignored: the planner
    /// is the authority (R5.3 D4).
    /// </summary>
    /// <param name="rows">The validated rows, in any order. Not mutated.</param>
    /// <param name="joinedId">The row joining the top-level run.</param>
    /// <param name="joinedRowNumber">The row's body index after its content moves.</param>
    /// <returns>The dense resequencing of the whole top-level run, ascending row order.</returns>
    private static List<SiblingOrderWrite> ResequenceTopLevelRunWithJoin(
        IReadOnlyList<GanttEvent> rows,
        GanttRowId joinedId,
        int joinedRowNumber)
    {
        List<SiblingOrderWrite> writes = [];
        var order = 0;
        foreach (GanttEvent row in rows
            .Where(row => row.ParentId is null && row.RowNumber < joinedRowNumber)
            .OrderBy(static row => row.RowNumber))
        {
            writes.Add(new SiblingOrderWrite(row.Id, order));
            order++;
        }

        writes.Add(new SiblingOrderWrite(joinedId, order));
        order++;

        foreach (GanttEvent row in rows
            .Where(row => row.ParentId is null && row.RowNumber > joinedRowNumber)
            .OrderBy(static row => row.RowNumber))
        {
            writes.Add(new SiblingOrderWrite(row.Id, order));
            order++;
        }

        return writes;
    }

    /// <summary>
    /// Dense resequencing of the top-level run after a row leaves it. Existing
    /// <c>SiblingOrder</c> values are ignored: the planner is the authority
    /// (R5.3 D4).
    /// </summary>
    private static List<SiblingOrderWrite> ResequenceTopLevelRun(
        IReadOnlyList<GanttEvent> rows,
        GanttRowId leavingId)
    {
        List<SiblingOrderWrite> writes = [];
        var order = 0;
        foreach (GanttEvent row in rows
            .Where(static row => row.ParentId is null)
            .OrderBy(static row => row.RowNumber))
        {
            if (row.Id == leavingId)
            {
                continue;
            }

            writes.Add(new SiblingOrderWrite(row.Id, order));
            order++;
        }

        return writes;
    }

    /// <summary>
    /// Whether the type is a span a <c>CriticalInterval</c> child may land on.
    /// Mirrors the validator's span rule without duplicating its diagnostics.
    /// </summary>
    private static bool IsSpan(GanttEntityType type) => type switch
    {
        GanttEntityType.Splitter => false,
        GanttEntityType.Spacer => false,
        GanttEntityType.AsBuiltActivity => true,
        GanttEntityType.AsPlannedActivity => true,
        GanttEntityType.BaselineActivity => true,
        GanttEntityType.CriticalInterval => false,
        GanttEntityType.DelayEvent => true,
        GanttEntityType.AsBuiltProcurement => true,
        GanttEntityType.AsPlannedProcurement => true,
        GanttEntityType.BaselineProcurement => true,
        GanttEntityType.CustomActivity => true,
        GanttEntityType.AsBuiltMilestone => false,
        GanttEntityType.AsPlannedMilestone => false,
        GanttEntityType.BaselineMilestone => false,
        GanttEntityType.CriticalMilestone => false,
        GanttEntityType.Delineator => false,
        _ => false,
    };
}
