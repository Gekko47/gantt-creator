namespace GanttCreator.Core;

/// <summary>
/// One Excel outline group: a parent row and the contiguous child rows beneath it.
/// </summary>
/// <param name="ParentRowNumber">The parent row, which carries the group's <c>-</c>/<c>+</c> control.</param>
/// <param name="ChildRowNumbers">The child rows, in order, immediately following the parent.</param>
/// <remarks>
/// Equality is by value, including the child sequence. A record's synthesised
/// equality would compare <see cref="ChildRowNumbers"/> by reference, so two plans
/// built from the same hierarchy in different input orders would compare unequal --
/// which is precisely the drift this planner exists to prevent, and it would make
/// the order-independence test unsatisfiable.
/// </remarks>
public sealed record OutlineGroup(int ParentRowNumber, IReadOnlyList<int> ChildRowNumbers)
{
    /// <inheritdoc />
    public bool Equals(OutlineGroup? other) =>
        other is not null
        && ParentRowNumber == other.ParentRowNumber
        && ChildRowNumbers.SequenceEqual(other.ChildRowNumbers);

    /// <inheritdoc />
    public override int GetHashCode()
    {
        var hash = new HashCode();
        hash.Add(ParentRowNumber);
        foreach (var child in ChildRowNumbers)
        {
            hash.Add(child);
        }

        return hash.ToHashCode();
    }
}

/// <summary>The result of planning outline groups from a hierarchy.</summary>
/// <param name="Groups">The groups, in row order.</param>
/// <param name="RowsToUngroup">
/// Rows whose grouping the plan cannot represent -- a child above its parent, or a
/// run broken by an unrelated row -- so the caller clears them rather than leaving
/// a stale group that Excel will refuse to overwrite.
/// </param>
public sealed record OutlineGroupPlan(
    IReadOnlyList<OutlineGroup> Groups,
    IReadOnlyList<int> RowsToUngroup)
{
    /// <summary>How many rows the plan wants grouped.</summary>
    public int GroupedRowCount => Groups.Sum(static group => group.ChildRowNumbers.Count);
}

/// <summary>
/// Plans the Excel outline groups that represent a parent/child hierarchy, so the
/// native <c>-</c>/<c>+</c> controls stay a presentation of the metadata rather
/// than becoming a second source of truth.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why the plan is derived from <c>ParentId</c> and never read back.</b> Excel's
/// outline is presentation; the hierarchy lives in the hidden <c>ParentId</c>
/// column. Were the outline authoritative, collapsing a group would change the
/// model and expanding one would silently redefine the tree. This type takes the
/// validated events and produces the outline the sheet should have.
/// </para>
/// <para>
/// <b>Why non-contiguity is reported rather than repaired.</b> An Excel outline
/// group covers a contiguous run. A hierarchy whose children are not contiguous
/// cannot be shown as one group without moving rows, and moving a user's rows is
/// their decision. Such a child is listed for ungrouping instead of being silently
/// dropped or the rows silently moved.
/// </para>
/// <para>
/// <b>Why ungrouping belongs in the plan.</b> A sheet carrying a stale group will
/// refuse the new grouping, so the caller must know which rows to clear before
/// applying. Deriving both from one pass keeps the read and the clear consistent.
/// </para>
/// <para>
/// Row numbers are <see cref="GanttEvent.RowNumber"/>, the body-row index. Mapping
/// those to worksheet rows is the caller's job, because only the Office layer knows
/// where the table sits on the sheet.
/// </para>
/// </remarks>
public static class OutlineGroupPlanner
{
    /// <summary>
    /// Plans the outline groups for the supplied validated events.
    /// </summary>
    /// <param name="events">
    /// The validated events, in any order. Only rows naming a parent that is itself
    /// an event of this table take part.
    /// </param>
    /// <returns>The groups and the rows to ungroup, both in row order.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="events"/> is null.</exception>
    public static OutlineGroupPlan Plan(IReadOnlyList<GanttEvent> events)
    {
        ArgumentNullException.ThrowIfNull(events);

        Dictionary<GanttRowId, int> rowNumberById = [];
        HashSet<GanttRowId> parents = [];
        foreach (GanttEvent candidate in events)
        {
            // A duplicated id is a validation error, so last-wins would be arbitrary.
            // First-wins is deterministic, and the duplicate is reported by the
            // validator that owns it.
            _ = rowNumberById.TryAdd(candidate.Id, candidate.RowNumber);
            if (candidate.ParentId is { } parent)
            {
                _ = parents.Add(parent);
            }
        }

        Dictionary<int, List<int>> childrenByParentRow = [];
        HashSet<int> ungrouped = [];

        foreach (GanttEvent child in events)
        {
            if (child.ParentId is not { } parentId
                || !parents.Contains(parentId)
                || !rowNumberById.TryGetValue(parentId, out var parentRow)
                || !rowNumberById.TryGetValue(child.Id, out var childRow))
            {
                // A parent that is not a row of this table cannot own a group here.
                continue;
            }

            if (childRow <= parentRow)
            {
                // A child above its parent cannot form a group. Reported, not moved.
                _ = ungrouped.Add(child.RowNumber);
                continue;
            }

            if (!childrenByParentRow.TryGetValue(parentRow, out List<int>? children))
            {
                children = [];
                childrenByParentRow[parentRow] = children;
            }

            children.Add(childRow);
        }

        List<OutlineGroup> groups = [];
        foreach ((var parentRow, List<int> children) in childrenByParentRow.OrderBy(static pair => pair.Key))
        {
            children.Sort();

            // An Excel group under a parent is the CONTIGUOUS block directly
            // beneath it, so the run is anchored at parentRow + 1. A child outside
            // that block cannot be drawn under this parent without moving rows, so
            // it is reported rather than given a second group of its own -- one
            // child must never be both grouped and ungrouped by the same plan.
            var run = new List<int>();
            foreach (var child in children)
            {
                if (run.Count > 0 ? child == run[^1] + 1 : child == parentRow + 1)
                {
                    run.Add(child);
                }
                else
                {
                    _ = ungrouped.Add(child);
                }
            }

            if (run.Count > 0)
            {
                groups.Add(new OutlineGroup(parentRow, [.. run]));
            }
        }

        return new OutlineGroupPlan(
            [.. groups.OrderBy(static group => group.ParentRowNumber)],
            [.. ungrouped.Order()]);
    }
}
