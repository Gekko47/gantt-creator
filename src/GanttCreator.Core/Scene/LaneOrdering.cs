namespace GanttCreator.Core.Scene;

/// <summary>Deterministic lane and effective-stack ordering rules.</summary>
public static class LaneOrdering
{
    /// <summary>Orders events by the approved lane key and stable tie-breaks.</summary>
    /// <param name="events">The layout inputs.</param>
    /// <returns>The deterministic event order.</returns>
    public static IReadOnlyList<LaneEventInput> OrderForLanes(IEnumerable<LaneEventInput> events) =>
        [
            .. events
                .OrderBy(input => input.Event.SortOrder ?? int.MaxValue)
                .ThenBy(input => input.Event.RowNumber)
                .ThenBy(input => input.Event.Id.Value, StringComparer.Ordinal),
        ];

    /// <summary>
    /// Gets the stable lane key for an event.
    /// </summary>
    /// <param name="input">The layout input.</param>
    /// <returns>The lane-bound key or a deterministic row-only key.</returns>
    /// <remarks>
    /// <para>
    /// A <c>Splitter</c> or <c>Spacer</c> always gets a row-scoped key, even when the
    /// row carries a <c>LaneId</c>. This is load-bearing rather than cosmetic: entity
    /// guide §10 puts a Splitter in "a complete lane across the included data panel
    /// and plot" and §11 makes a Spacer a blank lane, so neither may share vertical
    /// space with an activity. <see cref="LaneLayoutBuilder"/> picks its lane branch
    /// from the first input in a group, so a shared <c>LaneId</c> would group a
    /// Splitter with activities and build it as an ordinary event lane.
    /// </para>
    /// </remarks>
    public static string LaneKey(LaneEventInput input)
    {
        ArgumentNullException.ThrowIfNull(input);
        return OwnsItsOwnLane(input.Event) || input.Event.LaneId is null
            ? $"row:{input.Event.Id.Value}"
            : input.Event.LaneId.Value;
    }

    /// <summary>Whether a row type occupies a fixed-height lane of its own.</summary>
    /// <param name="atEvent">The validated event whose type is tested.</param>
    /// <returns><see langword="true"/> for a <c>Splitter</c> or <c>Spacer</c>.</returns>
    public static bool OwnsItsOwnLane(GanttEvent atEvent)
    {
        ArgumentNullException.ThrowIfNull(atEvent);
        return atEvent.Type is GanttEntityType.Splitter or GanttEntityType.Spacer;
    }
}
