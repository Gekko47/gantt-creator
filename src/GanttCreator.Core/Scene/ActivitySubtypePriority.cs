using System.Collections.Frozen;

namespace GanttCreator.Core.Scene;

/// <summary>Deterministic activity-body ordering within scene layer 40.</summary>
public static class ActivitySubtypePriority
{
    private static readonly FrozenDictionary<GanttEntityType, int> _priorities =
        new Dictionary<GanttEntityType, int>
        {
            [GanttEntityType.BaselineActivity] = 0,
            [GanttEntityType.BaselineProcurement] = 1,
            [GanttEntityType.AsPlannedActivity] = 2,
            [GanttEntityType.AsPlannedProcurement] = 3,
            [GanttEntityType.AsBuiltActivity] = 4,
            [GanttEntityType.AsBuiltProcurement] = 5,
            [GanttEntityType.CustomActivity] = 6,
            [GanttEntityType.DelayEvent] = 7,
        }.ToFrozenDictionary();

    /// <summary>Gets the back-to-front priority for an activity subtype.</summary>
    /// <param name="type">The activity subtype.</param>
    /// <returns>The documented layer-40 priority; non-activity values sort last.</returns>
    public static int For(GanttEntityType? type) =>
        type is not null && _priorities.TryGetValue(type.Value, out var priority)
            ? priority
            : int.MaxValue;
}
