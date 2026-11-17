namespace GanttCreator.Core;

/// <summary>
/// Which structural row actions the ribbon may offer for the row under the
/// active cell (R5.9b). Pure: a projection of the rows the table read already
/// produced and of <see cref="EntityHierarchyCatalog"/>, so the surface and the
/// commands answer from one rule instead of two.
/// </summary>
/// <remarks>
/// <para>
/// The four flags mirror the refusals the promote, demote and move commands
/// already produce: a disabled button always corresponds to a precondition the
/// command would refuse, and never to a condition the command can satisfy.
/// </para>
/// <para>
/// The sibling-scope rule is the honest one for the move pair: "up" is offered
/// only when a row above shares the active row's parent scope, so the first
/// child of a group is not offered an up that would silently do nothing. The
/// promote/demote rules read the hierarchy catalogue directly, and the demote
/// eligibility reuses <see cref="AddChildPlanner.IsEligibleParent"/> so the
/// ribbon and the inserter cannot disagree about who may own a child.
/// </para>
/// </remarks>
/// <param name="CanPromote">Whether the row may leave its group.</param>
/// <param name="CanDemote">Whether the row may become a child of an eligible parent above it.</param>
/// <param name="CanMoveUp">Whether a row above shares the row's parent scope.</param>
/// <param name="CanMoveDown">Whether a row below shares the row's parent scope.</param>
public readonly record struct RowActionAvailability(
    bool CanPromote,
    bool CanDemote,
    bool CanMoveUp,
    bool CanMoveDown)
{
    /// <summary>
    /// The "nothing is available" value: no resolved row, every action disabled.
    /// </summary>
    public static RowActionAvailability None => new(false, false, false, false);

    /// <summary>
    /// Evaluates the four structural actions for one body row.
    /// </summary>
    /// <param name="rows">The rows the table read returned, in body order.</param>
    /// <param name="activeBodyIndex">The one-based body index under the active cell, or null when it is outside the body.</param>
    /// <returns>
    /// The availability flags. An out-of-range index, an unknown type, or an
    /// unparsable <c>ParentId</c> reports <see cref="None"/>, because the row's
    /// depth is then unknown and every command would refuse it.
    /// </returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="rows"/> is null.</exception>
    public static RowActionAvailability Evaluate(IReadOnlyList<GanttRowDto> rows, int? activeBodyIndex)
    {
        ArgumentNullException.ThrowIfNull(rows);

        if (activeBodyIndex is not int index || index < 1 || index > rows.Count)
        {
            return None;
        }

        GanttRowDto selected = rows[index - 1];

        // An unparsable parent id makes the row's own depth unknown, so no
        // action is offered; the commands refuse the same row for the same
        // reason (SelectionUnknown / TableInvalid).
        if (!EntityTypeCatalog.TryParse(selected.TypeText, out GanttEntityType type)
            || !TryReadParentScope(selected.ParentId, out string? parentScope))
        {
            return None;
        }

        Dictionary<string, int> childCounts = CountChildren(rows);

        bool canPromote = parentScope is not null
            && ChildCountOf(childCounts, selected.Id) == 0
            && type != GanttEntityType.CriticalInterval;

        bool canDemote = parentScope is null
            && EntityHierarchyCatalog.MayBeChild(type)
            && HasEligibleParentAbove(rows, index, childCounts);

        return new RowActionAvailability(
            canPromote,
            canDemote,
            HasSiblingAbove(rows, index, parentScope),
            HasSiblingBelow(rows, index, parentScope));
    }

    /// <summary>
    /// Reads one <c>ParentId</c> cell as a parent scope.
    /// </summary>
    /// <param name="text">The raw cell text.</param>
    /// <param name="parentScope">The trimmed parent text, or null when the row is top-level.</param>
    /// <returns><see langword="false"/> when the text is present but not a valid identifier.</returns>
    private static bool TryReadParentScope(string? text, out string? parentScope)
    {
        parentScope = null;
        if (string.IsNullOrWhiteSpace(text))
        {
            return true;
        }

        if (!GanttRowId.TryParse(text, out GanttRowId? parsed) || parsed is null)
        {
            return false;
        }

        parentScope = parsed.Value;
        return true;
    }

    /// <summary>
    /// Counts how many rows each parent id owns.
    /// </summary>
    /// <param name="rows">The table body rows.</param>
    /// <returns>The child count per parent id, ordinal by id.</returns>
    private static Dictionary<string, int> CountChildren(IReadOnlyList<GanttRowDto> rows)
    {
        Dictionary<string, int> counts = new(StringComparer.Ordinal);
        foreach (GanttRowDto row in rows)
        {
            if (!TryReadParentScope(row.ParentId, out string? parentScope) || parentScope is null)
            {
                continue;
            }

            counts[parentScope] = counts.TryGetValue(parentScope, out int count) ? count + 1 : 1;
        }

        return counts;
    }

    /// <summary>Gets how many children one row owns.</summary>
    /// <param name="childCounts">The child counts per parent id.</param>
    /// <param name="idText">The row's raw id text.</param>
    /// <returns>The child count, or zero when the id is blank or names no children.</returns>
    private static int ChildCountOf(Dictionary<string, int> childCounts, string? idText)
    {
        string? id = idText?.Trim();
        return string.IsNullOrEmpty(id) ? 0 : childCounts.TryGetValue(id, out int count) ? count : 0;
    }

    /// <summary>
    /// Whether a row above the active one can accept it as a child.
    /// </summary>
    /// <param name="rows">The table body rows.</param>
    /// <param name="index">The one-based body index of the active row.</param>
    /// <param name="childCounts">The child counts per parent id.</param>
    /// <returns><see langword="true"/> when at least one eligible parent exists above.</returns>
    private static bool HasEligibleParentAbove(
        IReadOnlyList<GanttRowDto> rows,
        int index,
        Dictionary<string, int> childCounts)
    {
        for (int position = 1; position < index; position++)
        {
            GanttRowDto candidate = rows[position - 1];
            if (!EntityTypeCatalog.TryParse(candidate.TypeText, out GanttEntityType type)
                || !TryReadParentScope(candidate.ParentId, out string? scope))
            {
                continue;
            }

            GanttRowId? parentId = null;
            if (scope is not null && (!GanttRowId.TryParse(scope, out parentId) || parentId is null))
            {
                continue;
            }

            if (AddChildPlanner.IsEligibleParent(type, parentId, ChildCountOf(childCounts, candidate.Id)))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Whether a row above shares the active row's parent scope.
    /// </summary>
    /// <param name="rows">The table body rows.</param>
    /// <param name="index">The one-based body index of the active row.</param>
    /// <param name="parentScope">The active row's parent scope, or null when it is top-level.</param>
    /// <returns><see langword="true"/> when a preceding row has the same scope.</returns>
    private static bool HasSiblingAbove(IReadOnlyList<GanttRowDto> rows, int index, string? parentScope)
    {
        for (int position = 1; position < index; position++)
        {
            if (ScopeMatches(rows[position - 1], parentScope))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Whether a row below shares the active row's parent scope.
    /// </summary>
    /// <param name="rows">The table body rows.</param>
    /// <param name="index">The one-based body index of the active row.</param>
    /// <param name="parentScope">The active row's parent scope, or null when it is top-level.</param>
    /// <returns><see langword="true"/> when a following row has the same scope.</returns>
    private static bool HasSiblingBelow(IReadOnlyList<GanttRowDto> rows, int index, string? parentScope)
    {
        for (int position = index + 1; position <= rows.Count; position++)
        {
            if (ScopeMatches(rows[position - 1], parentScope))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Whether one row belongs to a parent scope. A row whose parent id cannot
    /// be parsed never matches: its depth is unknown, so it is not a sibling.
    /// </summary>
    /// <param name="row">The candidate row.</param>
    /// <param name="parentScope">The scope being matched, or null for top level.</param>
    /// <returns><see langword="true"/> when the row's scope equals the requested scope.</returns>
    private static bool ScopeMatches(GanttRowDto row, string? parentScope) =>
        TryReadParentScope(row.ParentId, out string? scope)
        && string.Equals(scope, parentScope, StringComparison.Ordinal);
}
