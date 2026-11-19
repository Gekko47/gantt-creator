namespace GanttCreator.Core;

/// <summary>
/// The visual indent each body row's <c>Description</c> cell carries, derived from the
/// outline groups the sheet is about to show (R5.18i).
/// </summary>
/// <remarks>
/// <para>
/// <b>Why this reads the group plan rather than <c>ParentId</c>.</b> The indent marks the
/// same relationship as the <c>-</c>/<c>+</c> bracket beside it. A row can name a parent
/// the sheet still refuses to group — the parent is not a row of this table, or the child
/// sits above it — and indenting that row would draw an indent with no bracket, or a
/// bracket with no indent. Either is worse than neither, so the plan
/// <see cref="OutlineGroupPlanner.Plan"/> produced is the only input here.
/// </para>
/// <para>
/// <b>Why depth becomes a constant.</b> "Indent by depth − 1" reduces to one step while the
/// supported hierarchy is two levels deep, and it must not silently become two steps the
/// day the hierarchy grows without the outline growing with it. The step is therefore
/// written as <see cref="EntityHierarchyCatalog.MaxDepth"/> minus one, and a test pins
/// that identity, so raising the limit fails there first.
/// </para>
/// <para>
/// No row value changes: the indent is cell formatting, so copy/paste, the reader, and the
/// chart are all untouched.
/// </para>
/// </remarks>
public static class HierarchyIndentation
{
    /// <summary>The indent a grouped child's description carries.</summary>
    /// <remarks>
    /// Numerically <c>EntityHierarchyCatalog.MaxDepth - 1</c>, which is <c>1</c> today: the
    /// parent/child pair is the only nesting the validator permits.
    /// </remarks>
    public const int ChildIndentLevel = EntityHierarchyCatalog.MaxDepth - 1;

    /// <summary>
    /// The width one indent step occupies in an exported data-panel cell, in points.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>A code constant, not a catalogue metric.</b> It is presentation for the editable
    /// copy only, exactly like <c>GanttTableAppearance.DefaultVisibleColumnWidthPt</c>: adding
    /// it to the VeryHidden catalogue would move the catalogue hash and force an Initialise
    /// for a value that changes no stored contract.
    /// </para>
    /// <para>
    /// The live sheet writes Excel's <c>IndentLevel</c>, whose step is host- and font-derived
    /// and is not exposed as a point value. The export draws the panel as shapes, so a point
    /// inset must stand in for that step; one body character is the closest honest
    /// equivalent at the body font size. It is an initial design default and may be tuned
    /// without changing the storage schema.
    /// </para>
    /// </remarks>
    public const double ChildIndentWidthPt = 6.0;

    /// <summary>
    /// Plans the non-zero indent for every body row the outline groups.
    /// </summary>
    /// <param name="plan">The outline plan the same pass is applying.</param>
    /// <returns>
    /// Body row number to indent level, containing an entry only for a row that is indented.
    /// A row absent from the map is one the plan does not group, and its documented indent
    /// is zero.
    /// </returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="plan"/> is null.</exception>
    public static IReadOnlyDictionary<int, int> Plan(OutlineGroupPlan plan)
    {
        ArgumentNullException.ThrowIfNull(plan);

        Dictionary<int, int> indents = [];
        foreach (OutlineGroup group in plan.Groups)
        {
            foreach (int childRowNumber in group.ChildRowNumbers)
            {
                indents[childRowNumber] = ChildIndentLevel;
            }
        }

        return indents;
    }

    /// <summary>
    /// Reads one body row's planned indent.
    /// </summary>
    /// <param name="indents">The map <see cref="Plan"/> produced.</param>
    /// <param name="bodyRowNumber">The body row, as <see cref="GanttEvent.RowNumber"/> numbers it.</param>
    /// <returns>The indent the row should carry; zero for an ungrouped row.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="indents"/> is null.</exception>
    public static int IndentFor(IReadOnlyDictionary<int, int> indents, int bodyRowNumber)
    {
        ArgumentNullException.ThrowIfNull(indents);

        return indents.TryGetValue(bodyRowNumber, out int indent) ? indent : 0;
    }
}
