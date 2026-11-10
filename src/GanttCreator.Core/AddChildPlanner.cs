namespace GanttCreator.Core;

/// <summary>The reason an add-child could not be planned.</summary>
public enum AddChildRefusalReason
{
    /// <summary>The selected identifier names no row of the table.</summary>
    SelectionUnknown = 0,

    /// <summary>The selected row's type may never own children.</summary>
    ParentCannotOwnChildren = 1,

    /// <summary>The requested type may never be a child.</summary>
    TypeCannotBeChild = 2,

    /// <summary>The parent already owns the maximum number of children.</summary>
    ParentAtCapacity = 3,

    /// <summary>The parent is itself a child, so the new row would nest three deep.</summary>
    TooDeep = 4,
}

/// <summary>
/// One cell write the add-child plan needs: the new row's parent and its
/// position among its siblings.
/// </summary>
/// <param name="RowId">The new row's stable id.</param>
/// <param name="ParentId">The parent's stable id.</param>
/// <param name="SiblingOrder">The new row's zero-based position among the parent's children.</param>
public sealed record AddChildWrite(GanttRowId RowId, GanttRowId ParentId, int SiblingOrder);

/// <summary>
/// The result of planning an add-child: the writes to apply, or the reason
/// it could not be planned.
/// </summary>
/// <param name="Writes">The writes, or null when the plan was refused.</param>
/// <param name="Refusal">The typed refusal reason, or null when the plan succeeded.</param>
public sealed record AddChildPlan(IReadOnlyList<AddChildWrite>? Writes, AddChildRefusalReason? Refusal)
{
    /// <summary>Gets whether a plan was produced.</summary>
    public bool Succeeded => Writes is not null;
}

/// <summary>
/// Plans the insertion of a new row as a CHILD of a selected parent row
/// (R5.3 WI-E).
/// </summary>
/// <remarks>
/// <para>
/// The hierarchy's rules live in <see cref="EntityHierarchyCatalog"/>, and this
/// planner is the only place that turns them into a decision about a new row.
/// The Ribbon's permitted-type list is projected from the same catalogue, so
/// the types the UI offers and the types the planner accepts cannot
/// disagree.
/// </para>
/// <para>
/// The new row lands DIRECTLY BENEATH the parent's existing child run, so the
/// parent's children stay one contiguous run and the Excel outline group the
/// group writer builds stays one group. A child inserted elsewhere would
/// split the run, and the writer would then have to ungroup and regroup
/// around it.
/// </para>
/// </remarks>
public static class AddChildPlanner
{
    /// <summary>
    /// Plans the insertion of a new row as a child of the selected parent.
    /// </summary>
    /// <param name="rows">The validated rows of the table, in any order.</param>
    /// <param name="selectedId">The selected parent's stable id.</param>
    /// <param name="childType">The type the new row will have.</param>
    /// <param name="newRowId">The stable id the new row will carry.</param>
    /// <returns>The plan, or the typed refusal.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="rows"/> is null.</exception>
    public static AddChildPlan Plan(
        IReadOnlyList<GanttEvent> rows,
        GanttRowId selectedId,
        GanttEntityType childType,
        GanttRowId newRowId)
    {
        ArgumentNullException.ThrowIfNull(rows);

        GanttEvent? parent = rows.FirstOrDefault(row => row.Id == selectedId);
        if (parent is null)
        {
            return Refused(AddChildRefusalReason.SelectionUnknown);
        }

        List<GanttEvent> siblings =
        [
            .. rows
                .Where(row => row.ParentId == selectedId)
                .OrderBy(static row => row.RowNumber),
        ];

        AddChildRefusalReason? refusal = Evaluate(
            parent.Type,
            parent.Id,
            parent.ParentId,
            siblings.Count,
            childType);
        if (refusal is not null)
        {
            return Refused(refusal.Value);
        }

        // The new row joins the end of the parent's existing child run, so the
        // run stays contiguous and the outline group stays one group.
        return new AddChildPlan(
            [new AddChildWrite(newRowId, selectedId, siblings.Count)],
            null);
    }

    /// <summary>
    /// The hierarchy rule body for one candidate parent (R5.16). Returns null
    /// when the child would be accepted, otherwise the typed refusal.
    /// </summary>
    /// <remarks>
    /// This is the single rule implementation: <see cref="Plan"/> delegates to
    /// it, the Ribbon's data gate asks <see cref="IsEligibleParent"/>, and the
    /// row inserter calls it before writing a <c>ParentId</c> (so a positional
    /// insert cannot create a parent the model forbids). The order of the
    /// checks is deliberate: an unknown parent is reported first, then the
    /// parent's right to own children, the child's right to be one, the depth
    /// limit, and finally the capacity limit.
    /// </remarks>
    /// <param name="parentType">The candidate parent's type, or null when no row was resolved.</param>
    /// <param name="parentId">The candidate parent's id, or null when no row was resolved.</param>
    /// <param name="parentOfParentId">The candidate parent's own parent id, or null when it is top-level.</param>
    /// <param name="existingChildCount">How many children the candidate already owns.</param>
    /// <param name="childType">The type the new row will have.</param>
    /// <returns>The typed refusal, or <see langword="null"/> when the child is accepted.</returns>
    public static AddChildRefusalReason? Evaluate(
        GanttEntityType? parentType,
        GanttRowId? parentId,
        GanttRowId? parentOfParentId,
        int existingChildCount,
        GanttEntityType childType)
    {
        if (parentType is not GanttEntityType type || parentId is null)
        {
            return AddChildRefusalReason.SelectionUnknown;
        }

        if (!EntityHierarchyCatalog.MayOwnChildren(type))
        {
            return AddChildRefusalReason.ParentCannotOwnChildren;
        }

        if (!EntityHierarchyCatalog.MayBeChild(childType))
        {
            return AddChildRefusalReason.TypeCannotBeChild;
        }

        // A parent that is itself a child would make the new row depth 3,
        // which the catalogue refuses rather than supports.
        return parentOfParentId is not null
            ? AddChildRefusalReason.TooDeep
            : existingChildCount >= EntityHierarchyCatalog.MaxChildrenPerParent
                ? AddChildRefusalReason.ParentAtCapacity
                : null;
    }

    /// <summary>
    /// Whether a row can accept a new child right now (R5.16): its type may own
    /// children, it is not itself a child (so the new row stays depth 2), and it
    /// owns fewer than <see cref="EntityHierarchyCatalog.MaxChildrenPerParent"/>
    /// children.
    /// </summary>
    /// <remarks>
    /// This is the single eligibility predicate. The Ribbon's Critical Interval
    /// gate asks it of every row in the table; the inserter asks it while
    /// scanning the body for the nearest parent; <see cref="Plan"/> applies it
    /// through the same rule body. One predicate, three callers, so the surface
    /// and the command cannot disagree about who may own a child.
    /// </remarks>
    /// <param name="type">The candidate parent's type.</param>
    /// <param name="parentId">The candidate parent's own parent id, or null when it is top-level.</param>
    /// <param name="existingChildCount">How many children the candidate already owns.</param>
    /// <returns><see langword="true"/> when the candidate can accept one more child.</returns>
    public static bool IsEligibleParent(
        GanttEntityType type,
        GanttRowId? parentId,
        int existingChildCount) =>
        EntityHierarchyCatalog.MayOwnChildren(type)
        && parentId is null
        && existingChildCount < EntityHierarchyCatalog.MaxChildrenPerParent;

    /// <summary>
    /// The child types that may be added to a parent of the given type.
    /// </summary>
    /// <param name="parentType">The parent's type.</param>
    /// <returns>
    /// The permitted child types in ascending enum order, or an empty list
    /// when the parent may never own a child.
    /// </returns>
    /// <remarks>
    /// This is the projection the Ribbon's type list is built from, so the
    /// list and <see cref="Plan"/>'s acceptance cannot disagree. The parent
    /// type's own rules are applied: a <c>CriticalInterval</c> parent takes
    /// only a span child.
    /// </remarks>
    public static IReadOnlyList<GanttEntityType> PermittedChildTypes(GanttEntityType parentType) =>
        EntityHierarchyCatalog.MayOwnChildren(parentType)
            ? [.. EntityHierarchyCatalog.TypesThatMayBeChild().Order()]
            : [];

    /// <summary>Builds a refused plan.</summary>
    /// <param name="reason">The refusal reason.</param>
    /// <returns>The refused plan.</returns>
    private static AddChildPlan Refused(AddChildRefusalReason reason) => new(null, reason);
}
