namespace GanttCreator.Core;

/// <summary>One row whose <c>StyleKey</c> cell must be rewritten on Refresh.</summary>
/// <param name="RowNumber">The one-based body-row index of the row to write.</param>
/// <param name="NewStyleKey">The row's current <c>Type</c>'s default named style key.</param>
public sealed record StyleKeyWrite(int RowNumber, string NewStyleKey);

/// <summary>
/// The typed outcome of planning the <c>StyleKey</c> write for a set of validated rows.
/// </summary>
/// <param name="Writes">
/// One entry per row whose <c>StyleKey</c> differs from its <c>Type</c>'s default.
/// A row already carrying its type's default is absent, so a correctly written
/// sheet produces an empty write set and the adapter performs no work at all.
/// </param>
public sealed record StyleKeyWritePlan(IReadOnlyList<StyleKeyWrite> Writes)
{
    /// <summary>How many cells the plan would write.</summary>
    public int WriteCount => Writes.Count;
}

/// <summary>
/// Plans the <c>StyleKey</c> write from validated rows: column K is brought in line
/// with each row's current <c>Type</c>.
/// </summary>
/// <remarks>
/// <para>
/// <b>Pure and Office-free.</b> It is a function over the validated events, so the
/// planner is fully testable without a workbook and the Office layer only performs
/// the write. That mirrors <see cref="DurationCalculator.Plan"/>.
/// </para>
/// <para>
/// <b>Why only rows that differ.</b> Duration establishes the convention: a sheet
/// already in the desired state produces an empty plan, so Refresh does not mark
/// the workbook dirty when nothing changed.
/// </para>
/// <para>
/// <b>Blank resolves to the default.</b> A blank cell reads as <see langword="null"/>
/// on the event, which never <see cref="string.Equals(string, StringComparison)"/>s a non-empty default, so a
/// blank <c>StyleKey</c> is filled with the type default. A stale non-null cell (for
/// example K4 stamped <c>AsPlannedActivity</c> when the row is now <c>As-Built
/// Activity</c>) is overwritten with <c>AsBuiltActivity</c>. Refresh cannot
/// distinguish a deliberate per-row override from a stale insert-time stamp, so the
/// current <c>Type</c> always wins.
/// </para>
/// <para>
/// <b>Why a <c>Custom Activity</c> or <c>Custom Milestone</c> choice is not
/// overwritten.</b> Both require an explicit style key
/// (<see cref="EntityTypeDefinition.RequiresStyleKey"/> is true), so the planner
/// never replaces a choice that still resolves: the user picked that style, and
/// the cell is theirs. Two cases are not a choice. A blank cell resolves to the
/// type's code-owned built-in default at render time and is left blank, so the
/// ribbon's <c>None</c> item keeps meaning "use the built-in default". A cell
/// naming a style that is no longer in the registry is a dangling reference left
/// behind by deleting that style; the registry-aware overload heals it to the
/// type's default, so the row renders instead of refusing and the sheet stops
/// carrying a reference to something that no longer exists.
/// </para>
/// <para>
/// <b>Undefined behaviour is an unknown, not a guess.</b> A validated event always
/// carries a catalogue type, so the lookup below cannot fail on production input;
/// if it ever does, the planner throws rather than writing a wrong key.
/// </para>
/// </remarks>
public static class StyleKeyWritePlanner
{
    /// <summary>
    /// Plans the <c>StyleKey</c> write for a batch of validated events.
    /// </summary>
    /// <param name="events">The validated events, in any order.</param>
    /// <returns>
    /// The write plan, ordered by row number so the plan is deterministic.
    /// </returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="events"/> is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when an event's <see cref="GanttEvent.Type"/> is not in the catalogue.
    /// </exception>
    public static StyleKeyWritePlan Plan(IReadOnlyList<GanttEvent> events) =>
        Plan(events, null);

    /// <summary>
    /// Plans the <c>StyleKey</c> write from validated rows against the workbook's
    /// named-style registry, healing a row whose named style has been deleted.
    /// </summary>
    /// <param name="events">The validated events, in any order.</param>
    /// <param name="styleRegistry">
    /// The workbook's named styles, or <see langword="null"/> for the compatibility
    /// path, which cannot tell a deleted style from an unchosen one and therefore
    /// leaves every style-key type's cell alone.
    /// </param>
    /// <returns>The write plan, ordered by row number so the plan is deterministic.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="events"/> is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when an event's <see cref="GanttEvent.Type"/> is not in the catalogue.
    /// </exception>
    public static StyleKeyWritePlan Plan(
        IReadOnlyList<GanttEvent> events,
        GanttStyleRegistry? styleRegistry)
    {
        ArgumentNullException.ThrowIfNull(events);

        List<StyleKeyWrite> writes = [];
        foreach (GanttEvent @event in events.OrderBy(static e => e.RowNumber))
        {
            EntityTypeDefinition? definition =
                EntityTypeCatalog.GetDefinition(@event.Type)
                ?? throw new ArgumentOutOfRangeException(
                    nameof(events),
                    @event.Type,
                    "The event's type is not in the catalogue.");

            string target = definition.DefaultStyleKey;

            if (definition.RequiresStyleKey)
            {
                // Custom Activity / Custom Milestone: see the class remarks. A
                // choice that still resolves is never overwritten; a dangling
                // reference is healed to the type's default.
                if (styleRegistry is null || target.Length == 0)
                {
                    continue;
                }

                string? chosen = @event.StyleKey;
                if (string.IsNullOrEmpty(chosen) || styleRegistry.TryGet(chosen, out _))
                {
                    continue;
                }

                writes.Add(new StyleKeyWrite(@event.RowNumber, target));
                continue;
            }

            if (target.Length == 0)
            {
                continue;
            }

            string? current = @event.StyleKey; // null == blank, which resolves to the default
            if (!string.Equals(current, target, StringComparison.Ordinal))
            {
                writes.Add(new StyleKeyWrite(@event.RowNumber, target));
            }
        }

        return new StyleKeyWritePlan(writes);
    }
}
