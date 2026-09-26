using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

public sealed class LaneEventLayoutTests
{
    private static readonly LaneLayoutMetrics _metrics = new(18, 3, 3, 2, 18, 9);
    private static readonly GanttRowId _laneId = GanttRowId.New();

    [Fact]
    public void Places_a_two_event_stack_with_the_exact_stack_gap_between_centres()
    {
        LaneEventInput[] events = [Input(1, 1), Input(2, 2)];

        LaneEventLayoutResult result = Layout(events);

        Assert.Equal(2, result.Placements.Count);
        Assert.Equal([0, 1], result.Placements.Select(p => p.VisualSlotIndex));

        // Each slot is 8pt tall, separated by StackGapPt = 2, from lane top 3.
        // The centres are therefore 7 and 17, exactly 10pt apart.
        Assert.Equal(7, result.Placements[0].SlotCentreY);
        Assert.Equal(17, result.Placements[1].SlotCentreY);
        Assert.Equal(10, result.Placements[1].SlotCentreY - result.Placements[0].SlotCentreY, 10);
    }

    [Fact]
    public void Places_a_three_event_stack_with_two_gaps()
    {
        LaneEventLayoutResult result = Layout([Input(1, 1), Input(2, 2), Input(3, 3)]);

        Assert.Equal(3, result.Placements.Count);
        Assert.Equal([0, 1, 2], result.Placements.Select(p => p.VisualSlotIndex));
        Assert.Equal(7, result.Placements[0].SlotCentreY);
        Assert.Equal(17, result.Placements[1].SlotCentreY);
        Assert.Equal(27, result.Placements[2].SlotCentreY);
    }

    [Fact]
    public void Grows_the_lane_when_stacked_content_exceeds_the_minimum_height()
    {
        LaneEventInput[] events = [Input(1, 1), Input(2, 2), Input(3, 3)];

        LaneLayoutResult layout = Lane(events);
        LaneEventLayoutResult result = Layout(events, layout);

        // Three 8pt slots plus two 2pt gaps and 6pt of padding is 34pt, so the
        // lane grew past the 18pt minimum rather than compressing the events.
        LaneGeometry lane = Assert.Single(layout.Lanes);
        Assert.Equal(34, lane.Height);
        Assert.Equal(3, result.Placements.Count);
    }

    [Fact]
    public void Gives_same_slot_events_one_shared_centre_without_warning_when_sequential()
    {
        // Two events on one effective slot whose dates do not intersect: they
        // share the line and never warn, because nothing is ambiguous.
        LaneEventInput first = Input(1, 1, new DateOnly(2024, 1, 1), new DateOnly(2024, 1, 3), effective: 0);
        LaneEventInput second = Input(2, 2, new DateOnly(2024, 1, 5), new DateOnly(2024, 1, 7), effective: 0);

        LaneEventLayoutResult result = Layout([first, second]);

        Assert.Equal(2, result.Placements.Count);
        Assert.Equal(result.Placements[0].SlotCentreY, result.Placements[1].SlotCentreY);
        Assert.Equal(0, result.Placements[0].VisualSlotIndex);
        Assert.Equal(0, result.Placements[1].VisualSlotIndex);
        Assert.Empty(result.Warnings);
    }

    [Fact]
    public void Warns_once_for_same_slot_events_that_overlap_and_keeps_one_centre()
    {
        // §9: same effective slot, overlapping dates. The events are never moved
        // or hidden; they share the centre and the ambiguity is reported.
        LaneEventInput first = Input(1, 1, new DateOnly(2024, 1, 1), new DateOnly(2024, 1, 5), effective: 0);
        LaneEventInput second = Input(2, 2, new DateOnly(2024, 1, 3), new DateOnly(2024, 1, 7), effective: 0);

        LaneEventLayoutResult result = Layout([first, second]);

        Assert.Equal(2, result.Placements.Count);
        Assert.Equal(result.Placements[0].SlotCentreY, result.Placements[1].SlotCentreY);
        SceneWarning warning = Assert.Single(result.Warnings);
        Assert.Equal(LaneEventLayout.AmbiguousStackOverlapCode, warning.Code);
    }

    [Fact]
    public void Separates_planned_and_actual_on_one_lane_at_full_length()
    {
        // §13: different effective stacks separate vertically and neither event
        // is shortened, moved, or hidden because the other overlaps it.
        LaneEventInput planned = Input(
            1,
            1,
            new DateOnly(2024, 1, 1),
            new DateOnly(2024, 1, 10),
            GanttEntityType.AsPlannedActivity
        );
        LaneEventInput actual = Input(
            2,
            2,
            new DateOnly(2024, 1, 5),
            new DateOnly(2024, 1, 15),
            GanttEntityType.AsBuiltActivity
        );

        LaneEventLayoutResult result = Layout([planned, actual]);

        Assert.Equal(2, result.Placements.Count);
        Assert.NotEqual(result.Placements[0].SlotCentreY, result.Placements[1].SlotCentreY);
        // The subtype order is deterministic: planned is behind actual at layer 40.
        Assert.Equal(GanttEntityType.AsPlannedActivity, result.Placements[0].Event.Type);
        Assert.Equal(GanttEntityType.AsBuiltActivity, result.Placements[1].Event.Type);
    }

    [Fact]
    public void Keeps_both_events_when_planned_and_actual_share_one_effective_stack()
    {
        // §13: equal values deliberately share the same line, and neither event is
        // hidden. The layout reports both placements.
        LaneEventInput planned = Input(
            1,
            1,
            new DateOnly(2024, 1, 1),
            new DateOnly(2024, 1, 10),
            GanttEntityType.AsPlannedActivity,
            effective: 0
        );
        LaneEventInput actual = Input(
            2,
            2,
            new DateOnly(2024, 1, 5),
            new DateOnly(2024, 1, 15),
            GanttEntityType.AsBuiltActivity,
            effective: 0
        );

        LaneEventLayoutResult result = Layout([planned, actual]);

        Assert.Equal(2, result.Placements.Count);
        Assert.All(result.Placements, placement => Assert.NotNull(placement.Event));
        Assert.Equal(result.Placements[0].SlotCentreY, result.Placements[1].SlotCentreY);
    }

    [Fact]
    public void Ignores_the_visible_stack_index_cell_entirely()
    {
        // ADR-0012: the visible cell is compatibility data. Two events with
        // wildly different visible values but no effective assignment still get
        // consecutive Core-derived slots.
        GanttEvent first = Event(1, visibleStack: 7, effectiveType: GanttEntityType.AsPlannedActivity);
        GanttEvent second = Event(2, visibleStack: 2, effectiveType: GanttEntityType.AsPlannedActivity);

        LaneEventLayoutResult result = Layout([new LaneEventInput(first, 8), new LaneEventInput(second, 8)]);

        Assert.Equal([0, 1], result.Placements.Select(p => p.VisualSlotIndex));
        Assert.Equal([0, 1], result.Placements.Select(p => p.EffectiveStackIndex));
    }

    [Fact]
    public void Preserves_order_without_creating_empty_height_for_sparse_values()
    {
        LaneEventInput first = Input(1, 1, effective: 2);
        LaneEventInput second = Input(2, 2, effective: 7);

        LaneEventLayoutResult result = Layout([first, second]);

        // R3.4 compacts sparse compatibility values into consecutive visual
        // slots, so no empty band appears between them.
        Assert.Equal([0, 1], result.Placements.Select(p => p.VisualSlotIndex));
        Assert.Equal([2, 7], result.Placements.Select(p => p.EffectiveStackIndex));
    }

    [Fact]
    public void Places_mixed_bar_and_diamond_heights_on_one_lane()
    {
        // A milestone is 12pt tall and a bar 8pt, so each slot takes the maximum
        // of what it holds (R3.4's rule) and the feed honours that centre.
        LaneEventInput bar = new LaneEventInput(
            Event(1, effectiveType: GanttEntityType.AsPlannedActivity),
            8
        );
        LaneEventInput diamond = new LaneEventInput(
            Event(2, effectiveType: GanttEntityType.AsPlannedMilestone),
            12
        );

        LaneEventLayoutResult result = Layout([bar, diamond]);

        Assert.Equal(2, result.Placements.Count);
        // Slot 0 is 8pt tall from lane top 3, so its centre is 7. Slot 1 is the
        // 12pt diamond after a 2pt gap, spanning 13..25, so its centre is 19.
        Assert.Equal(7, result.Placements[0].SlotCentreY);
        Assert.Equal(19, result.Placements[1].SlotCentreY);
    }

    [Fact]
    public void Orders_identical_z_layer_entities_deterministically_by_the_shared_key()
    {
        LaneEventLayoutResult result = Layout(
            [Input(3, 1), Input(1, 2), Input(2, 3)]
        );

        // Within one slot the order is activity-subtype priority, then SortOrder,
        // then the stable ID — never the arrival order.
        Assert.Equal([1, 2, 3], result.Placements.Select(p => p.Event.RowNumber));
    }

    [Fact]
    public void Produces_identical_placements_when_the_input_order_is_shuffled()
    {
        LaneEventInput a = Input(1, 1);
        LaneEventInput b = Input(2, 2);
        LaneEventInput c = Input(3, 3);

        string[] first = Layout([a, b, c]).Placements.Select(p => p.Event.Id.Value).ToArray();
        string[] shuffled = Layout([c, a, b]).Placements.Select(p => p.Event.Id.Value).ToArray();
        string[] reversed = Layout([b, c, a]).Placements.Select(p => p.Event.Id.Value).ToArray();

        Assert.Equal(first, shuffled);
        Assert.Equal(first, reversed);
    }

    [Fact]
    public void Refuses_null_input()
    {
        LaneEventLayoutCreationOutcome outcome = LaneEventLayout.TryBuild(null, Lane([Input(1, 1)]));

        Assert.False(outcome.Succeeded);
        Assert.Equal(LaneEventLayoutRefusal.NullInput, outcome.Refusal);
    }

    [Fact]
    public void Refuses_a_missing_layout()
    {
        LaneEventLayoutCreationOutcome outcome = LaneEventLayout.TryBuild([Input(1, 1)], null);

        Assert.False(outcome.Succeeded);
        Assert.Equal(LaneEventLayoutRefusal.MissingLayout, outcome.Refusal);
    }

    [Fact]
    public void Refuses_a_null_event()
    {
        LaneEventLayoutCreationOutcome outcome = LaneEventLayout.TryBuild(
            [new LaneEventInput(null!, 8)],
            Lane([Input(1, 1)])
        );

        Assert.False(outcome.Succeeded);
        Assert.Equal(LaneEventLayoutRefusal.NullEvent, outcome.Refusal);
    }

    [Fact]
    public void Refuses_duplicate_event_identity()
    {
        GanttRowId shared = GanttRowId.New();
        LaneEventLayoutCreationOutcome outcome = LaneEventLayout.TryBuild(
            [
                new LaneEventInput(Event(1, id: shared), 8),
                new LaneEventInput(Event(2, id: shared), 8),
            ],
            Lane([Input(1, 1)])
        );

        Assert.False(outcome.Succeeded);
        Assert.Equal(LaneEventLayoutRefusal.DuplicateEventId, outcome.Refusal);
    }

    [Fact]
    public void Refuses_a_negative_effective_slot()
    {
        LaneEventLayoutCreationOutcome outcome = LaneEventLayout.TryBuild(
            [Input(1, 1, effective: -1)],
            Lane([Input(1, 1)])
        );

        Assert.False(outcome.Succeeded);
        Assert.Equal(LaneEventLayoutRefusal.InvalidSlotIndex, outcome.Refusal);
    }

    private static LaneEventLayoutResult Layout(
        IReadOnlyList<LaneEventInput> events,
        LaneLayoutResult? layout = null
    )
    {
        LaneLayoutResult resolved = layout ?? Lane(events);
        LaneEventLayoutCreationOutcome outcome = LaneEventLayout.TryBuild(events, resolved);
        Assert.True(outcome.Succeeded);
        return outcome.Result!;
    }

    private static LaneLayoutResult Lane(IReadOnlyList<LaneEventInput> events) =>
        LaneLayoutBuilder.TryBuild(events, _metrics).Layout!;

    private static LaneEventInput Input(
        int rowNumber,
        int dayOffset,
        DateOnly? start = null,
        DateOnly? finish = null,
        GanttEntityType type = GanttEntityType.AsPlannedActivity,
        int? effective = null
    ) =>
        new(
            Event(
                rowNumber,
                start: start ?? new DateOnly(2024, 1, 1).AddDays(dayOffset),
                finish: finish ?? new DateOnly(2024, 1, 1).AddDays(dayOffset + 2),
                effectiveType: type
            ),
            8,
            effective
        );

    private static GanttEvent Event(
        int rowNumber,
        int? visibleStack = null,
        GanttEntityType effectiveType = GanttEntityType.AsPlannedActivity,
        GanttRowId? id = null,
        DateOnly? start = null,
        DateOnly? finish = null
    ) =>
        new(
            rowNumber,
            id ?? GanttRowId.New(),
            _laneId,
            visibleStack,
            effectiveType,
            "Event",
            start ?? new DateOnly(2024, 1, 1).AddDays(rowNumber),
            finish ?? new DateOnly(2024, 1, 1).AddDays(rowNumber + 2),
            null,
            null,
            null,
            null,
            null,
            true,
            null
        );
}
