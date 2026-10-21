namespace GanttCreator.Core;

/// <summary>Which direction a collapse or expand request moves the groups.</summary>
public enum OutlineCollapseDirection
{
    /// <summary>Hide the grouped child rows, collapsing every group.</summary>
    Collapse = 0,

    /// <summary>Show the grouped child rows, expanding every group.</summary>
    Expand = 1,
}

/// <summary>
/// One body row's outline level, as the sheet reports it.
/// </summary>
/// <param name="BodyRowNumber">The one-based body-row index.</param>
/// <param name="OutlineLevel">
/// The row's outline level: <see cref="OutlineCollapsePlanner.TopLevelOutlineLevel"/>
/// for a row that belongs to no group, and a deeper level for a grouped child.
/// </param>
public sealed record OutlineRowLevel(int BodyRowNumber, int OutlineLevel);

/// <summary>
/// The rows a collapse or expand request writes, derived from the outline the
/// sheet currently carries.
/// </summary>
/// <param name="Direction">Whether the rows are being hidden or shown.</param>
/// <param name="ChildRowNumbers">
/// The body-row numbers of the grouped child rows the request writes, ascending
/// and without duplicates. Empty when the sheet carries no group.
/// </param>
/// <remarks>
/// <para>
/// <b>Why the plan comes from the sheet's outline.</b> This is a PRESENTATION
/// command: it moves the collapsed state the user can already see through
/// Excel's own <c>-</c>/<c>+</c> controls. The groups the native control acts
/// on are the groups the sheet carries, so the request must act on exactly
/// those rows — otherwise a group the user can see would be left behind, or a
/// group they cannot see would be moved. The hierarchy in the <c>ParentId</c>
/// column is unaffected: it stays the model, and the next Refresh still
/// rebuilds the outline from it (ADR-0026 D5, R4.7D).
/// </para>
/// <para>
/// <b>Why a child row is written, not its parent.</b> Collapsing a group in
/// Excel hides the group's child rows; the parent stays visible and keeps its
/// <c>-</c> control. A live probe measured this: Excel's own collapse leaves
/// the child rows <c>Hidden = $true</c> and the parent <c>Hidden = $false</c>,
/// and Excel's own expand clears the child rows' flag and leaves the parent's
/// alone. Writing exactly those rows therefore reproduces what the native
/// control does, rather than a look-alike of it.
/// </para>
/// <para>
/// <b>Why the request is row-scoped.</b> The rows are named individually, so a
/// request can only ever touch the Gantt table's own body rows. The worksheet
/// level <c>Outline.ShowLevels</c> alternative was measured and rejected: it is
/// sheet-wide, so it would collapse the user's own unrelated outline groups.
/// </para>
/// </remarks>
public sealed record OutlineCollapsePlan(OutlineCollapseDirection Direction, IReadOnlyList<int> ChildRowNumbers)
{
    /// <summary>Gets how many child rows the request writes.</summary>
    public int RowCount => ChildRowNumbers.Count;

    /// <summary>Gets whether the request has nothing to write.</summary>
    public bool IsEmpty => ChildRowNumbers.Count == 0;
}

/// <summary>
/// Plans the rows a collapse-all or expand-all request writes, so the Ribbon
/// commands act on exactly the groups the sheet carries.
/// </summary>
public static class OutlineCollapsePlanner
{
    /// <summary>The outline level of a row that belongs to no group.</summary>
    public const int TopLevelOutlineLevel = 1;

    /// <summary>
    /// Plans the request for the outline the sheet carries.
    /// </summary>
    /// <param name="rows">
    /// Every body row's outline level, in any order. A row whose level is
    /// deeper than <see cref="TopLevelOutlineLevel"/> belongs to a group and
    /// is a candidate for the request.
    /// </param>
    /// <param name="direction">Whether the groups are being collapsed or expanded.</param>
    /// <returns>
    /// The rows to write: the grouped child rows, in ascending body-row order
    /// and without duplicates. Empty when the sheet carries no group.
    /// </returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="rows"/> is null.</exception>
    public static OutlineCollapsePlan Plan(
        IReadOnlyList<OutlineRowLevel> rows,
        OutlineCollapseDirection direction)
    {
        ArgumentNullException.ThrowIfNull(rows);

        List<int> childRows = [];
        HashSet<int> seen = [];
        foreach (OutlineRowLevel row in rows)
        {
            // A row at the top level belongs to no group, so it is never
            // written: a group has no child to hide without it, and the
            // parent keeps its own collapse control.
            if (row.OutlineLevel <= TopLevelOutlineLevel)
            {
                continue;
            }

            if (seen.Add(row.BodyRowNumber))
            {
                childRows.Add(row.BodyRowNumber);
            }
        }

        childRows.Sort();

        return new OutlineCollapsePlan(direction, childRows);
    }
}
