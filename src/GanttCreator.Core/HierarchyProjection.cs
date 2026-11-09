namespace GanttCreator.Core;

/// <summary>
/// Projects the rows as they were READ into the neutral events the outline planner consumes, so
/// the outline has one source of truth regardless of which path asks for it.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why this exists (defect, 2026-11-09).</b> The Add activity command rebuilt the outline from
/// the rows it had just written, while the Refresh rebuilt it from the VALIDATED events. A newly
/// inserted activity has no dates yet, so it is not a usable event and is absent from the validated
/// list; the Refresh's plan therefore contained no group for it, and its reset pass ungrouped the
/// rows the insert had just grouped — 313 ms later, from the log. The insert was correct and the
/// Refresh contradicted it, because the two planned from different data.
/// </para>
/// <para>
/// <b>Hierarchy is not validity.</b> A row whose <c>ParentId</c> names a real row of this table
/// belongs in that parent's group whether or not its dates are filled in. Validation decides what
/// may be RENDERED; the outline is presentation of the <c>ParentId</c> column, so it follows the
/// column. A row whose <c>Id</c> cannot be parsed is skipped, because an outline group is addressed
/// by row number and an unaddressable row cannot be a member of anything.
/// </para>
/// <para>
/// The events carry only the fields the outline planner reads — <c>RowNumber</c>, <c>Id</c>,
/// <c>Type</c> and <c>ParentId</c>. A <c>Type</c> the catalogue cannot parse keeps the enum default
/// and matters not at all here, so it is not reported; this projection makes no claim about the
/// row being renderable.
/// </para>
/// </remarks>
public static class HierarchyProjection
{
    /// <summary>
    /// Builds the outline-planning events from the rows as read, including rows that are not yet
    /// valid activities.
    /// </summary>
    /// <param name="rows">The rows exactly as the table reader returned them.</param>
    /// <returns>
    /// One event per row with a parseable <c>Id</c>, in row order, carrying that row's
    /// <c>ParentId</c> when it has one.
    /// </returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="rows"/> is null.</exception>
    public static IReadOnlyList<GanttEvent> FromRows(IReadOnlyList<GanttRowDto> rows)
    {
        ArgumentNullException.ThrowIfNull(rows);

        List<GanttEvent> events = new(rows.Count);
        foreach (GanttRowDto row in rows)
        {
            if (!GanttRowId.TryParse(row.Id, out GanttRowId? id) || id is null)
            {
                continue;
            }

            GanttRowId? parentId =
                GanttRowId.TryParse(row.ParentId, out GanttRowId? parsedParent) && parsedParent is not null
                    ? parsedParent
                    : null;

            _ = EntityTypeCatalog.TryParse(row.TypeText, out GanttEntityType type);

            events.Add(new GanttEvent(
                row.RowNumber,
                id,
                null,
                null,
                type,
                null,
                null,
                null,
                parentId,
                null,
                null,
                null,
                null,
                true,
                null));
        }

        return events;
    }
}
