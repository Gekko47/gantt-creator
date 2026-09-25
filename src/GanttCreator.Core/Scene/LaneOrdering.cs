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

    /// <summary>Gets the stable lane key for an event.</summary>
    /// <param name="input">The layout input.</param>
    /// <returns>The lane-bound key or a deterministic row-only key.</returns>
    public static string LaneKey(LaneEventInput input)
    {
        ArgumentNullException.ThrowIfNull(input);
        return input.Event.LaneId?.Value ?? $"row:{input.Event.Id.Value}";
    }
}
