namespace GanttCreator.Core.Scene;

/// <summary>The reason a lane layout could not be created.</summary>
public enum LaneLayoutRefusal
{
    /// <summary>The event collection was null.</summary>
    NullInput = 0,

    /// <summary>The event collection was empty.</summary>
    EmptyInput = 1,

    /// <summary>The metrics input was null.</summary>
    NullMetrics = 2,

    /// <summary>One or more metrics were non-finite or negative.</summary>
    InvalidMetrics = 3,

    /// <summary>An input item or its event was null.</summary>
    NullEvent = 4,

    /// <summary>A resolved event height was non-finite or invalid for its type.</summary>
    InvalidEventHeight = 5,

    /// <summary>An effective compatibility stack value was negative.</summary>
    InvalidEffectiveStack = 6,

    /// <summary>Two input events carried the same stable row ID.</summary>
    DuplicateEventId = 7,
}

/// <summary>The typed result of attempting to build lane geometry.</summary>
/// <param name="Layout">The successful layout, or <see langword="null"/>.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/>.</param>
public sealed record LaneLayoutCreationOutcome(LaneLayoutResult? Layout, LaneLayoutRefusal? Refusal)
{
    /// <summary>Gets whether layout creation succeeded.</summary>
    public bool Succeeded => Layout is not null;
}

/// <summary>Builds deterministic vertical lane geometry without rendering or measuring.</summary>
public static class LaneLayoutBuilder
{
    private const string _ambiguousStackOverlapCode = "AmbiguousStackOverlap";

    /// <summary>Attempts to build deterministic lane geometry.</summary>
    /// <param name="events">The validated event layout inputs.</param>
    /// <param name="metrics">The resolved lane metrics.</param>
    /// <returns>A typed layout outcome.</returns>
    public static LaneLayoutCreationOutcome TryBuild(IReadOnlyList<LaneEventInput>? events, LaneLayoutMetrics? metrics)
    {
        if (events is null)
        {
            return Refused(LaneLayoutRefusal.NullInput);
        }

        if (events.Count == 0)
        {
            return Refused(LaneLayoutRefusal.EmptyInput);
        }

        if (metrics is null)
        {
            return Refused(LaneLayoutRefusal.NullMetrics);
        }

        if (!ValidMetrics(metrics))
        {
            return Refused(LaneLayoutRefusal.InvalidMetrics);
        }

        HashSet<GanttRowId> ids = [];
        foreach (LaneEventInput? input in events)
        {
            if (input?.Event is null)
            {
                return Refused(LaneLayoutRefusal.NullEvent);
            }

            if (!double.IsFinite(input.ResolvedHeightPt) || input.ResolvedHeightPt < 0)
            {
                return Refused(LaneLayoutRefusal.InvalidEventHeight);
            }

            if (input.EffectiveStackIndex is { } stack && stack < 0)
            {
                return Refused(LaneLayoutRefusal.InvalidEffectiveStack);
            }

            if (!ids.Add(input.Event.Id))
            {
                return Refused(LaneLayoutRefusal.DuplicateEventId);
            }
        }

        IReadOnlyList<LaneEventInput> ordered = LaneOrdering.OrderForLanes(events);
        List<LaneGeometry> lanes = [];
        List<SceneWarning> warnings = [];
        double laneTop = 0;
        var laneOrder = 0;
        foreach (IGrouping<string, LaneEventInput> group in ordered.GroupBy(LaneOrdering.LaneKey))
        {
            LaneEventInput[] laneInputs = [.. group];
            bool isSplitter = laneInputs[0].Event.Type == GanttEntityType.Splitter;
            bool isSpacer = laneInputs[0].Event.Type == GanttEntityType.Spacer;
            LaneGeometry lane =
                isSplitter || isSpacer
                    ? BuildFixedLane(group.Key, laneOrder, laneTop, laneInputs, isSplitter, isSpacer, metrics)
                    : BuildEventLane(group.Key, laneOrder, laneTop, laneInputs, metrics, warnings);
            lanes.Add(lane);
            laneTop += lane.Height;
            laneOrder++;
        }

        return new LaneLayoutCreationOutcome(new LaneLayoutResult(lanes, warnings), null);
    }

    private static LaneGeometry BuildFixedLane(
        string laneKey,
        int laneOrder,
        double laneTop,
        IReadOnlyList<LaneEventInput> inputs,
        bool isSplitter,
        bool isSpacer,
        LaneLayoutMetrics metrics
    )
    {
        double height = isSplitter ? metrics.SplitterHeightPt : metrics.SpacerHeightPt;
        GanttRowId[] eventIds = [.. inputs.Select(input => input.Event.Id)];
        SlotGeometry slot = new(0, 0, laneTop, laneTop + (height / 2), laneTop + height, height, eventIds);
        return new LaneGeometry(laneKey, laneOrder, laneTop, height, [slot], eventIds, isSplitter, isSpacer);
    }

    private static LaneGeometry BuildEventLane(
        string laneKey,
        int laneOrder,
        double laneTop,
        IReadOnlyList<LaneEventInput> inputs,
        LaneLayoutMetrics metrics,
        List<SceneWarning> warnings
    )
    {
        LaneEventInput[] rowOrdered =
        [
            .. inputs.OrderBy(input => input.Event.RowNumber).ThenBy(input => input.Event.Id.Value, StringComparer.Ordinal),
        ];
        Dictionary<int, List<LaneEventInput>> effectiveSlots = [];
        for (var index = 0; index < rowOrdered.Length; index++)
        {
            int effective = rowOrdered[index].EffectiveStackIndex ?? index;
            if (!effectiveSlots.TryGetValue(effective, out List<LaneEventInput>? slotEvents))
            {
                slotEvents = [];
                effectiveSlots.Add(effective, slotEvents);
            }

            slotEvents.Add(rowOrdered[index]);
        }

        int[] effectiveValues = [.. effectiveSlots.Keys.OrderBy(value => value)];
        double contentHeight = metrics.LanePaddingTopPt + metrics.LanePaddingBottomPt;
        contentHeight += Math.Max(0, effectiveValues.Length - 1) * metrics.StackGapPt;
        contentHeight += effectiveValues.Sum(value => effectiveSlots[value].Max(input => input.ResolvedHeightPt));
        double laneHeight = Math.Max(metrics.LaneHeightPt, contentHeight);
        List<SlotGeometry> slots = [];
        double slotTop = laneTop + metrics.LanePaddingTopPt;
        for (var visualIndex = 0; visualIndex < effectiveValues.Length; visualIndex++)
        {
            int effective = effectiveValues[visualIndex];
            List<LaneEventInput> slotEvents = effectiveSlots[effective];
            double height = slotEvents.Max(input => input.ResolvedHeightPt);
            GanttRowId[] eventIds =
            [
                .. slotEvents
                    .OrderBy(input => input.Event.SortOrder ?? int.MaxValue)
                    .ThenBy(input => input.Event.Id.Value, StringComparer.Ordinal)
                    .Select(input => input.Event.Id),
            ];
            slots.Add(new SlotGeometry(visualIndex, effective, slotTop, slotTop + (height / 2), slotTop + height, height, eventIds));
            AddAmbiguityWarning(slotEvents, warnings);
            slotTop += height + (visualIndex < effectiveValues.Length - 1 ? metrics.StackGapPt : 0);
        }

        GanttRowId[] laneEventIds =
        [
            .. inputs
                .OrderBy(input => input.Event.SortOrder ?? int.MaxValue)
                .ThenBy(input => input.Event.Id.Value, StringComparer.Ordinal)
                .Select(input => input.Event.Id),
        ];
        return new LaneGeometry(laneKey, laneOrder, laneTop, laneHeight, slots, laneEventIds, false, false);
    }

    private static void AddAmbiguityWarning(List<LaneEventInput> slotEvents, List<SceneWarning> warnings)
    {
        for (var firstIndex = 0; firstIndex < slotEvents.Count; firstIndex++)
        {
            for (var secondIndex = firstIndex + 1; secondIndex < slotEvents.Count; secondIndex++)
            {
                if (Overlaps(slotEvents[firstIndex].Event, slotEvents[secondIndex].Event))
                {
                    warnings.Add(
                        new SceneWarning(
                            SceneOwnerId.ForRow(slotEvents[firstIndex].Event.Id),
                            _ambiguousStackOverlapCode,
                            "Events share an effective stack slot and overlap in time."
                        )
                    );
                    return;
                }
            }
        }
    }

    private static bool Overlaps(GanttEvent first, GanttEvent second)
    {
        if (first.Start is not { } firstStart || second.Start is not { } secondStart)
        {
            return false;
        }

        DateOnly firstFinish = first.Finish ?? firstStart;
        DateOnly secondFinish = second.Finish ?? secondStart;
        return firstStart <= secondFinish && secondStart <= firstFinish;
    }

    private static bool ValidMetrics(LaneLayoutMetrics metrics) =>
        IsFiniteNonNegative(metrics.LaneHeightPt)
        && metrics.LaneHeightPt > 0
        && IsFiniteNonNegative(metrics.LanePaddingTopPt)
        && IsFiniteNonNegative(metrics.LanePaddingBottomPt)
        && IsFiniteNonNegative(metrics.StackGapPt)
        && IsFiniteNonNegative(metrics.SplitterHeightPt)
        && IsFiniteNonNegative(metrics.SpacerHeightPt);

    private static bool IsFiniteNonNegative(double value) => double.IsFinite(value) && value >= 0;

    private static LaneLayoutCreationOutcome Refused(LaneLayoutRefusal refusal) => new(null, refusal);
}
