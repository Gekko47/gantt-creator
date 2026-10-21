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

        if (!EntityHierarchyCatalog.MayOwnChildren(parent.Type))
        {
            return Refused(AddChildRefusalReason.ParentCannotOwnChildren);
        }

        if (!EntityHierarchyCatalog.MayBeChild(childType))
        {
            return Refused(AddChildRefusalReason.TypeCannotBeChild);
        }

        // A parent that is itself a child would make the new row depth 3,
        // which the catalogue refuses rather than supports.
        if (parent.ParentId is not null)
        {
            return Refused(AddChildRefusalReason.TooDeep);
        }

        List<GanttEvent> siblings =
        [
            .. rows
                .Where(row => row.ParentId == selectedId)
                .OrderBy(static row => row.RowNumber),
        ];

        if (siblings.Count >= EntityHierarchyCatalog.MaxChildrenPerParent)
        {
            return Refused(AddChildRefusalReason.ParentAtCapacity);
        }

        // The new row joins the end of the parent's existing child run, so the
        // run stays contiguous and the outline group stays one group.
        return new AddChildPlan(
            [new AddChildWrite(newRowId, selectedId, siblings.Count)],
            null);
    }

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
