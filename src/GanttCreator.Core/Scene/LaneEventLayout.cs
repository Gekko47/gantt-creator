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

    /// <summary>The lane layout was null.</summary>
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

        if (layout is null)
        {
            return Refused(LaneEventLayoutRefusal.MissingLayout);
        }

        // An empty event set with a lane layout is a successful empty feed, not a
        // refusal. A §24-only scene has plot-global entities and no lanes, so
        // requiring at least one lane here would refuse the very scenes the empty
        // input from LaneLayoutBuilder exists to allow. A layout that fails to cover
        // a non-empty input is still caught, by the completeness check below.
        if (events.Count == 0)
        {
            return new LaneEventLayoutCreationOutcome(
                new LaneEventLayoutResult([], layout.Warnings),
                null
            );
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

        // Duplicates are detected before any numbering, not during it. The derived
        // map is keyed by row ID, so a repeated ID would be silently overwritten by
        // the second occurrence and the duplicate would then surface as a missing
        // slot -- reporting a layout problem for what is really two rows claiming one
        // identity.
        Dictionary<GanttRowId, LaneEventPlacement> placements = [];
        HashSet<GanttRowId> seen = [];
        foreach (LaneEventInput input in events)
        {
            if (!seen.Add(input.Event.Id))
            {
                return Refused(LaneEventLayoutRefusal.DuplicateEventId);
            }
        }

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

            // R4.7B: a projected child shares its lane owner's slot, so it takes the
            // owner's index rather than one of its own. This MUST number exactly as
            // `LaneLayoutBuilder.BuildEventLane` numbers, because the two passes read
            // the same map: the builder decides which slots exist, and this feed
            // resolves each event to one of them. Numbering the whole array -- rather
            // than only the non-projected events -- gave the two passes different
            // answers whenever a child sat above its parent, and the child was then
            // resolved to a slot that did not exist.
            //
            // Owners are indexed first, in row order, and only then do children look
            // one up: an authoring UI may place a child above its parent, so a child
            // is not guaranteed to precede the owner it depends on.
            Dictionary<GanttRowId, int> effectiveById = [];
            var ownIndex = 0;
            foreach (LaneEventInput input in laneEvents)
            {
                if (input.IsProjected)
                {
                    continue;
                }

                effectiveById[input.Event.Id] = ownIndex;
                ownIndex++;
            }

            foreach (LaneEventInput input in laneEvents)
            {
                if (!input.IsProjected)
                {
                    continue;
                }

                if (input.RenderLaneOwner is { } projectedOwner
                    && effectiveById.TryGetValue(projectedOwner.Id, out var ownerEffective))
                {
                    effectiveById[input.Event.Id] = ownerEffective;
                }
            }

            foreach (LaneEventInput input in laneEvents)
            {
                GanttEvent @event = input.Event;

                // A supplied compatibility value takes precedence, exactly as
                // `LaneLayoutBuilder.BuildEventLane` applies it, so a caller that
                // names its own slot is never overridden. Only a projected child
                // whose owner is not in this lane has no derived index at all, and
                // that is a broken projection rather than a placement decision.
                int effective;
                if (input.EffectiveStackIndex is { } supplied)
                {
                    effective = supplied;
                }
                else if (effectiveById.TryGetValue(@event.Id, out var derived))
                {
                    effective = derived;
                }
                else
                {
                    return Refused(LaneEventLayoutRefusal.MissingLayout);
                }

                if (effective < 0)
                {
                    return Refused(LaneEventLayoutRefusal.InvalidSlotIndex);
                }

                if (!slotsByLane[lane.LaneKey].TryGetValue(effective, out SlotGeometry? slot))
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

        // Every input event must reach a placement. The per-lane filter above
        // selects only events whose lane key matches a layout lane, so an event on
        // a lane the layout never produced is silently dropped — a caller would
        // then render a scene missing a row with no refusal and no warning. The
        // count is the authoritative completeness check, and MissingLayout is the
        // honest reason: the layout does not cover the inputs.
        if (placements.Count != events.Count)
        {
            return Refused(LaneEventLayoutRefusal.MissingLayout);
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
