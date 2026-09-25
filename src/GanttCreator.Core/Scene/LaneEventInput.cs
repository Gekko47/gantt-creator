namespace GanttCreator.Core.Scene;

/// <summary>One validated event plus the already-resolved height used for lane layout.</summary>
/// <param name="Event">The validated event.</param>
/// <param name="ResolvedHeightPt">The resolved rectangle or milestone height.</param>
/// <param name="EffectiveStackIndex">
/// An optional Core-assigned compatibility stack value. Normal authoring leaves
/// this null so the builder derives the value from row position. It is never
/// populated from the visible worksheet cell.
/// </param>
public sealed record LaneEventInput(GanttEvent Event, double ResolvedHeightPt, int? EffectiveStackIndex = null);
