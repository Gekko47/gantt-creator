namespace GanttCreator.Core.Scene;

/// <summary>One validated event plus the already-resolved height used for lane layout.</summary>
/// <param name="Event">The validated event.</param>
/// <param name="ResolvedHeightPt">The resolved rectangle or milestone height.</param>
/// <param name="EffectiveStackIndex">
/// An optional Core-assigned compatibility stack value. Normal authoring leaves
/// this null so the builder derives the value from row position. It is never
/// populated from the visible worksheet cell.
/// </param>
/// <param name="RenderLaneOwner">
/// The event whose lane <paramref name="Event"/> draws on, when R4.7B's
/// <see cref="ProjectionResolver"/> resolved it onto a parent. Null for a
/// top-level row, which owns its own lane. The owner is always top-level, so a
/// projected child never creates a lane and never grows one.
/// </param>
public sealed record LaneEventInput(
    GanttEvent Event,
    double ResolvedHeightPt,
    int? EffectiveStackIndex = null,
    GanttEvent? RenderLaneOwner = null)
{
    /// <summary>
    /// Whether this event draws on a lane other than its own row's. A projected
    /// child creates no lane and consumes no lane height.
    /// </summary>
    public bool IsProjected => RenderLaneOwner is not null && RenderLaneOwner.Id != Event.Id;
}
