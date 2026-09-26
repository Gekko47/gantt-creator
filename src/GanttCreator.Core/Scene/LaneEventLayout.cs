namespace GanttCreator.Core.Scene;

/// <summary>One laid-out event placed into its visual slot.</summary>
/// <param name="Event">The validated event.</param>
/// <param name="VisualSlotIndex">The consecutive visual slot the event occupies.</param>
/// <param name="EffectiveStackIndex">The Core-derived compatibility stack value.</param>
/// <param name="SlotCentreY">The slot's vertical centre in points.</param>
/// <param name="LaneOrder">The lane ordering value.</param>
public sealed record LaneEventPlacement(
    GanttEvent Event,
    int VisualSlotIndex,
    int EffectiveStackIndex,
    double SlotCentreY,
    int LaneOrder
);

/// <summary>The result of the per-lane event grouping feed.</summary>
/// <param name="Placements">
/// Every event, in the deterministic order the scene must emit: lane, then
/// visual slot, then activity-subtype priority, then <c>SortOrder</c>, then the
/// stable row ID.
/// </param>
/// <param name="Warnings">The deterministic non-blocking layout warnings.</param>
public sealed record LaneEventLayoutResult(
    IReadOnlyList<LaneEventPlacement> Placements,
    IReadOnlyList<SceneWarning> Warnings
);

/// <summary>The reason a lane event layout could not be produced.</summary>
public enum LaneEventLayoutRefusal
{
    /// <summary>The input collection was null.</summary>
    NullInput = 0,

    /// <summary>The lane layout was null or carried no lanes.</summary>
    MissingLayout = 1,

    /// <summary>An input item or its event was null.</summary>
    NullEvent = 2,

    /// <summary>Two input events carried the same stable row ID.</summary>
    DuplicateEventId = 3,

    /// <summary>One or more slot indices were negative.</summary>
    InvalidSlotIndex = 4,
}

/// <summary>The typed result of attempting to build a lane event layout.</summary>
/// <param name="Result">The successful result, or <see langword="null"/>.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/>.</param>
public sealed record LaneEventLayoutCreationOutcome(LaneEventLayoutResult? Result, LaneEventLayoutRefusal? Refusal)
{
    /// <summary>Gets whether layout creation succeeded.</summary>
    public bool Succeeded => Result is not null;
}

/// <summary>
/// Groups events by lane and binds each to the visual slot R3.4 already derived,
/// so the span, milestone, and critical builders all receive one authoritative
/// slot centre.
/// </summary>
/// <remarks>
/// <para>
/// This is the integration feed described by entity guide §9, not a second set
/// of geometry rules. It re-derives nothing: the slot centres, the effective
/// stack assignment, and the lane growth all come from <see cref="LaneGeometry"/>.
/// </para>
/// <para>
/// The visible <c>StackIndex</c> cell is never read. ADR-0012 makes the effective
/// stack a Core-derived value, and the compatibility cell is neither trusted nor
/// required for layout, so an event's visible value cannot move it.
/// </para>
/// </remarks>
public static class LaneEventLayout
{
    /// <summary>
    /// The warning code for events sharing one effective slot and overlapping in
    /// time. This mirrors <see cref="LaneLayoutBuilder"/>'s code, which is the
    /// authoritative emitter; this feed re-emits nothing so a user never sees the
    /// same ambiguity twice.
    /// </summary>
    public const string AmbiguousStackOverlapCode = "AmbiguousStackOverlap";

    /// <summary>Attempts to build the per-lane event placement feed.</summary>
    /// <param name="events">The validated event layout inputs, as R3.4 consumes them.</param>
    /// <param name="layout">The lane layout produced by <see cref="LaneLayoutBuilder"/>.</param>
    /// <returns>A typed result or refusal.</returns>
    public static LaneEventLayoutCreationOutcome TryBuild(
        IReadOnlyList<LaneEventInput>? events,
        LaneLayoutResult? layout
    )
    {
        if (events is null)
        {
            return Refused(LaneEventLayoutRefusal.NullInput);
        }

        if (layout is null || layout.Lanes.Count == 0)
        {
            return Refused(LaneEventLayoutRefusal.MissingLayout);
        }

        if (events.Any(input => input?.Event is null))
        {
            return Refused(LaneEventLayoutRefusal.NullEvent);
        }

        // Index each event's slot by the lane key and the effective stack value
        // that R3.4 already assigned. Nothing here re-derives a slot, and the
        // visible StackIndex cell is never consulted.
        Dictionary<string, Dictionary<int, SlotGeometry>> slotsByLane = [];
        foreach (LaneGeometry lane in layout.Lanes)
        {
            Dictionary<int, SlotGeometry> slots = [];
            foreach (SlotGeometry slot in lane.Slots)
            {
                if (slot.VisualSlotIndex < 0 || slot.EffectiveStackIndex < 0)
                {
                    return Refused(LaneEventLayoutRefusal.InvalidSlotIndex);
                }

                slots[slot.EffectiveStackIndex] = slot;
            }

            slotsByLane[lane.LaneKey] = slots;
        }

        Dictionary<GanttRowId, LaneEventPlacement> placements = [];
        HashSet<GanttRowId> seen = [];

        foreach (LaneGeometry lane in layout.Lanes)
        {
            // R3.4 assigns the effective stack from the row's position within the
            // lane when the compatibility value is absent, and a supplied value is
            // compacted into the slot it names. This feed resolves an event to its
            // slot by matching that same rule, so no slot is re-derived here and
            // the visible StackIndex cell is never read.
            LaneEventInput[] laneEvents =
            [
                .. events
                    .Where(input => input?.Event is not null && LaneOrdering.LaneKey(input) == lane.LaneKey)
                    .OrderBy(input => input.Event.RowNumber)
                    .ThenBy(input => input.Event.Id.Value, StringComparer.Ordinal),
            ];

            for (var index = 0; index < laneEvents.Length; index++)
            {
                LaneEventInput input = laneEvents[index];
                GanttEvent @event = input.Event;
                if (!seen.Add(@event.Id))
                {
                    return Refused(LaneEventLayoutRefusal.DuplicateEventId);
                }

                var effective = input.EffectiveStackIndex ?? index;
                if (effective < 0)
                {
                    return Refused(LaneEventLayoutRefusal.InvalidSlotIndex);
                }

                if (!slotsByLane[lane.LaneKey].TryGetValue(effective, out var slot))
                {
                    return Refused(LaneEventLayoutRefusal.MissingLayout);
                }

                placements[@event.Id] = new LaneEventPlacement(
                    @event,
                    slot.VisualSlotIndex,
                    effective,
                    slot.Centre,
                    lane.LaneOrder
                );
            }
        }

        IReadOnlyList<LaneEventPlacement> ordered =
        [
            .. placements.Values
                .OrderBy(placement => placement.LaneOrder)
                .ThenBy(placement => placement.VisualSlotIndex)
                .ThenBy(placement => ActivitySubtypePriority.For(placement.Event.Type))
                .ThenBy(placement => placement.Event.SortOrder ?? int.MaxValue)
                .ThenBy(placement => placement.Event.Id.Value, StringComparer.Ordinal),
        ];

        // §9's ambiguity warning is emitted by LaneLayoutBuilder, which owns the
        // lane geometry. The warnings are carried through so the scene assembly
        // has the complete set without this feed double-reporting one condition.
        return new LaneEventLayoutCreationOutcome(
            new LaneEventLayoutResult(ordered, layout.Warnings),
            null
        );
    }

    private static LaneEventLayoutCreationOutcome Refused(LaneEventLayoutRefusal refusal) => new(null, refusal);
}
