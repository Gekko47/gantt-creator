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
}
