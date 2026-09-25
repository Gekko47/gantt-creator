namespace GanttCreator.Core.Scene;

/// <summary>One vertical visual slot within a lane.</summary>
/// <param name="VisualSlotIndex">The consecutive visual slot number.</param>
/// <param name="EffectiveStackIndex">The Core-assigned compatibility stack value.</param>
/// <param name="Top">The slot top in points.</param>
/// <param name="Centre">The slot centre in points.</param>
/// <param name="Bottom">The slot bottom in points.</param>
/// <param name="Height">The resolved slot height in points.</param>
/// <param name="EventIds">The stable event IDs sharing this slot.</param>
public sealed record SlotGeometry(
    int VisualSlotIndex,
    int EffectiveStackIndex,
    double Top,
    double Centre,
    double Bottom,
    double Height,
    IReadOnlyList<GanttRowId> EventIds
);
