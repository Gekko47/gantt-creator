namespace GanttCreator.Core.Scene;

/// <summary>
/// The §22 description-label placement priority for an entity type, highest first.
/// </summary>
/// <remarks>
/// <para>
/// §22 fixes the order as "critical milestones, other milestones, delay labels,
/// actual labels, planned labels, baseline labels, procurement/custom labels, date
/// labels, and delineator labels". The date and delineator label passes place their
/// own text outside this list, so those ranks need no member.
/// </para>
/// <para>
/// The order is published rather than kept private to <see cref="SceneBuilder"/>
/// because it is a product contract in its own right, exactly like
/// <see cref="ActivitySubtypePriority"/>: a rank that is only observable through a
/// side effect cannot be tested when the lane layout keeps two rows' labels in
/// separate vertical bands, which is the normal case.
/// </para>
/// </remarks>
public static class LabelPlacementPriority
{
    /// <summary>Gets the placement rank for an entity type, lowest value first.</summary>
    /// <param name="type">The entity type whose description label is being placed.</param>
    /// <returns>
    /// The §22 rank. A type §22 does not name sorts last, which is deterministic
    /// and cannot outrank a named one.
    /// </returns>
    /// <remarks>
    /// <para>
    /// All three procurement types share the rank with <c>CustomActivity</c>.
    /// §22 lists "procurement/custom labels" as one group after the baseline group,
    /// so procurement is not interleaved with the activity family it belongs to:
    /// ranking as-built procurement alongside as-built <em>activity</em> let a
    /// procurement label outrank a planned or baseline activity label, which is the
    /// reverse of the stated order.
    /// </para>
    /// <para>
    /// Every member is listed explicitly rather than folded into a default arm, so
    /// a new <see cref="GanttEntityType"/> cannot be added without deciding its
    /// label priority.
    /// </para>
    /// </remarks>
    public static int For(GanttEntityType type) =>
        type switch
        {
            GanttEntityType.CriticalMilestone => 0,
            GanttEntityType.AsPlannedMilestone
                or GanttEntityType.AsBuiltMilestone
                or GanttEntityType.BaselineMilestone => 1,
            GanttEntityType.DelayEvent => 2,
            GanttEntityType.AsBuiltActivity => 3,
            GanttEntityType.AsPlannedActivity => 4,
            GanttEntityType.BaselineActivity => 5,
            GanttEntityType.AsBuiltProcurement
                or GanttEntityType.AsPlannedProcurement
                or GanttEntityType.BaselineProcurement
                or GanttEntityType.CustomActivity => 6,

            // A Critical Interval and a Delineator have no §22 description label of
            // their own (the first is an overlay, the second places its own corner
            // text), and a Splitter or Spacer emits no label at all. They are ranked
            // anyway so the switch provably covers the enum.
            GanttEntityType.CriticalInterval => 7,
            GanttEntityType.Delineator => 8,
            GanttEntityType.Splitter or GanttEntityType.Spacer => 9,
            _ => int.MaxValue,
        };
}
