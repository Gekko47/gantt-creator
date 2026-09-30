namespace GanttCreator.Core.Scene;

/// <summary>The reason a lane layout could not be created.</summary>
public enum LaneLayoutRefusal
{
    /// <summary>The event collection was null.</summary>
    NullInput = 0,

    /// <summary>The metrics input was null.</summary>
    NullMetrics = 1,

    /// <summary>One or more metrics were non-finite or negative.</summary>
    InvalidMetrics = 2,

    /// <summary>An input item or its event was null.</summary>
    NullEvent = 3,

    /// <summary>A resolved event height was non-finite or invalid for its type.</summary>
    InvalidEventHeight = 4,

    /// <summary>An effective compatibility stack value was negative.</summary>
    InvalidEffectiveStack = 5,

    /// <summary>Two input events carried the same stable row ID.</summary>
    DuplicateEventId = 6,
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
    /// <param name="events">
    /// The validated event layout inputs. An empty collection is valid and produces an
    /// empty layout: §24 makes a Delineator a plot-global entity that consumes no lane.
    /// </param>
    /// <param name="metrics">The resolved lane metrics.</param>
    /// <returns>A typed layout outcome.</returns>
    public static LaneLayoutCreationOutcome TryBuild(IReadOnlyList<LaneEventInput>? events, LaneLayoutMetrics? metrics)
    {
        if (events is null)
        {
            return Refused(LaneLayoutRefusal.NullInput);
        }

        if (metrics is null)
        {
            return Refused(LaneLayoutRefusal.NullMetrics);
        }

        if (!ValidMetrics(metrics))
        {
            return Refused(LaneLayoutRefusal.InvalidMetrics);
        }

        // An empty input is a successful empty layout, not a refusal. Entity guide
        // §24 makes a Delineator a full-height plot line that consumes no lane, so a
        // scene of only delineators legitimately has no lanes at all. Refusing here
        // forced `SceneBuilder` to call that scene "empty" and refuse it as
        // `EmptyEvents`, which wrongly reported a renderable scene as having nothing
        // to render. Whether the scene has any content at all is the orchestrator's
        // question, not this builder's.
        if (events.Count == 0)
        {
            return new LaneLayoutCreationOutcome(new LaneLayoutResult([], []), null);
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
            var isSplitter = laneInputs[0].Event.Type == GanttEntityType.Splitter;
            var isSpacer = laneInputs[0].Event.Type == GanttEntityType.Spacer;
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
        var height = isSplitter ? metrics.SplitterHeightPt : metrics.SpacerHeightPt;
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

        // R4.7B: a projected child takes its lane owner's effective stack index, so
        // it lands in the SAME slot and overlaps. Assigning it its own index by
        // position would give it a second slot, stacking it below the parent and
        // growing the lane -- which ADR-0026 D7 forbids and which is the whole
        // reason a child is projected rather than given a lane of its own.
        //
        // Two passes, because the owner is not guaranteed to precede the child in
        // row order: an authoring UI may place a child above its parent. Owners are
        // therefore indexed first, and only then do children look one up.
        var effectiveById = new Dictionary<GanttRowId, int>();
        var ownIndex = 0;

        foreach (LaneEventInput input in rowOrdered)
        {
            if (input.IsProjected)
            {
                continue;
            }

            effectiveById[input.Event.Id] = ownIndex;
            ownIndex++;
        }
        foreach (LaneEventInput input in rowOrdered)
        {
            if (!input.IsProjected)
            {
                continue;
            }

            if (input.RenderLaneOwner is { } owner
                && effectiveById.TryGetValue(owner.Id, out var ownerIndex))
            {
                effectiveById[input.Event.Id] = ownerIndex;
            }
        }

        Dictionary<int, List<LaneEventInput>> effectiveSlots = [];
        foreach (LaneEventInput input in rowOrdered)
        {
            var effective = input.EffectiveStackIndex ?? effectiveById[input.Event.Id];
            if (!effectiveSlots.TryGetValue(effective, out List<LaneEventInput>? slotEvents))
            {
                slotEvents = [];
                effectiveSlots.Add(effective, slotEvents);
            }

            slotEvents.Add(input);
        }

        int[] effectiveValues = [.. effectiveSlots.Keys.OrderBy(value => value)];
        var contentHeight = metrics.LanePaddingTopPt + metrics.LanePaddingBottomPt;
        contentHeight += Math.Max(0, effectiveValues.Length - 1) * metrics.StackGapPt;
        contentHeight += effectiveValues.Sum(value => effectiveSlots[value].Max(input => input.ResolvedHeightPt));
        var laneHeight = Math.Max(metrics.LaneHeightPt, contentHeight);
        var slots = new List<SlotGeometry>();
        var slotTop = laneTop + metrics.LanePaddingTopPt;
        for (var visualIndex = 0; visualIndex < effectiveValues.Length; visualIndex++)
        {
            var effective = effectiveValues[visualIndex];
            List<LaneEventInput> slotEvents = effectiveSlots[effective];
            var height = slotEvents.Max(input => input.ResolvedHeightPt);
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

    /// <summary>
    /// Emits the same-stack overlap warning for a slot, once per slot.
    /// </summary>
    /// <remarks>
    /// <para>
    /// R4.7B: a projected child shares its parent's slot <b>by design</b>, and
    /// overlapping its parent is the intended presentation — REV6 §8 says several
    /// children on one parent lane "may overlap", distinguished by date geometry,
    /// Type, style, z-order and label. Emitting <c>AmbiguousStackOverlap</c> for
    /// that case would report the product working, and it would also collide with
    /// any genuine ambiguity on the same owner, because <c>SceneValidator</c> keys
    /// duplicate warnings by (owner, code) and a parent can own more than one
    /// slot.
    /// </para>
    /// <para>
    /// So a pair where one event is projected onto the other is not ambiguous and
    /// warns nothing; genuine same-slot overlap between unrelated rows still does.
    /// </para>
    /// </remarks>
    private static void AddAmbiguityWarning(List<LaneEventInput> slotEvents, List<SceneWarning> warnings)
    {
        for (var firstIndex = 0; firstIndex < slotEvents.Count; firstIndex++)
        {
            for (var secondIndex = firstIndex + 1; secondIndex < slotEvents.Count; secondIndex++)
            {
                if (IsProjectionPair(slotEvents[firstIndex], slotEvents[secondIndex]))
                {
                    continue;
                }

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

    /// <summary>
    /// Whether one of the two events is projected onto the lane the other owns, in
    /// either direction. Such a pair shares a slot because the hierarchy says so.
    /// </summary>
    private static bool IsProjectionPair(LaneEventInput first, LaneEventInput second)
    {
        GanttRowId? firstParent = first.RenderLaneOwner?.Id;
        GanttRowId? secondParent = second.RenderLaneOwner?.Id;
        return (firstParent == second.Event.Id) || (secondParent == first.Event.Id);
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
