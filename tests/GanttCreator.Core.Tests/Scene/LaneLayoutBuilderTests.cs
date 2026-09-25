using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

public sealed class LaneLayoutBuilderTests
{
    private static readonly LaneLayoutMetrics _metrics = new(18, 3, 3, 2, 18, 9);
    private static readonly GanttRowId _laneId = GanttRowId.New();

    [Fact]
    public void Effective_stack_is_derived_from_row_position_and_ignores_visible_values()
    {
        GanttEvent second = Event(2, "second", visibleStack: 7);
        GanttEvent first = Event(1, "first", visibleStack: 2);
        LaneLayoutCreationOutcome outcome = LaneLayoutBuilder.TryBuild(
            [new LaneEventInput(second, 8), new LaneEventInput(first, 8)],
            _metrics
        );

        Assert.True(outcome.Succeeded);
        LaneGeometry lane = Assert.Single(outcome.Layout!.Lanes);
        Assert.Equal([0, 1], lane.Slots.Select(slot => slot.EffectiveStackIndex));
        Assert.Equal([0, 1], lane.Slots.Select(slot => slot.VisualSlotIndex));
        Assert.Equal(3, lane.Slots[0].Top);
        Assert.Equal(7, lane.Slots[0].Centre);
        Assert.Equal(13, lane.Slots[1].Top);
        Assert.Equal(17, lane.Slots[1].Centre);
    }

    [Fact]
    public void Lane_grows_for_stacked_content_without_compressing_events()
    {
        LaneEventInput[] events = [new(Event(1, "first"), 8), new(Event(2, "second"), 8), new(Event(3, "third"), 8)];

        LaneLayoutCreationOutcome outcome = LaneLayoutBuilder.TryBuild(events, _metrics);

        LaneGeometry lane = Assert.Single(outcome.Layout!.Lanes);
        Assert.Equal(34, lane.Height);
        Assert.Equal(3, lane.Slots[0].Top);
        Assert.Equal(31, lane.Slots[2].Bottom);
    }

    [Fact]
    public void Sparse_and_duplicate_effective_values_are_compacted_and_duplicate_overlap_warns()
    {
        GanttEvent first = Event(1, "first", start: new DateOnly(2024, 1, 1), finish: new DateOnly(2024, 1, 3));
        GanttEvent second = Event(2, "second", start: new DateOnly(2024, 1, 2), finish: new DateOnly(2024, 1, 4));
        GanttEvent third = Event(3, "third", start: new DateOnly(2024, 1, 4), finish: new DateOnly(2024, 1, 5));

        LaneLayoutCreationOutcome outcome = LaneLayoutBuilder.TryBuild(
            [new LaneEventInput(first, 8, 2), new LaneEventInput(second, 8, 7), new LaneEventInput(third, 8, 7)],
            _metrics
        );

        Assert.True(outcome.Succeeded);
        LaneGeometry lane = Assert.Single(outcome.Layout!.Lanes);
        Assert.Equal([0, 1], lane.Slots.Select(slot => slot.VisualSlotIndex));
        Assert.Equal([2, 7], lane.Slots.Select(slot => slot.EffectiveStackIndex));
        Assert.Equal(8, lane.Slots[0].Height);
        Assert.Single(outcome.Layout.Warnings);
        Assert.Equal("AmbiguousStackOverlap", outcome.Layout.Warnings[0].Code);
        Assert.Equal(2, lane.Slots[1].EventIds.Count);
    }

    [Fact]
    public void Splitter_and_spacer_lanes_use_fixed_heights_and_preserve_sequence()
    {
        LaneEventInput spacer = new(Event(1, "spacer", GanttEntityType.Spacer), 1);
        LaneEventInput splitter = new(Event(2, "splitter", GanttEntityType.Splitter), 1);

        LaneLayoutCreationOutcome outcome = LaneLayoutBuilder.TryBuild([splitter, spacer], _metrics);

        Assert.True(outcome.Succeeded);
        Assert.Equal([9, 18], outcome.Layout!.Lanes.Select(lane => lane.Height));
        Assert.True(outcome.Layout.Lanes[0].IsSpacer);
        Assert.True(outcome.Layout.Lanes[1].IsSplitter);
        Assert.Equal(9, outcome.Layout.Lanes[1].Top);
    }

    [Fact]
    public void Reordered_inputs_produce_identical_lane_geometry()
    {
        LaneEventInput first = new(Event(1, "first"), 8);
        LaneEventInput second = new(Event(2, "second"), 10);
        LaneEventInput third = new(Event(3, "third"), 6);
        LaneLayoutResult expected = LaneLayoutBuilder.TryBuild([first, second, third], _metrics).Layout!;
        LaneLayoutResult actual = LaneLayoutBuilder.TryBuild([third, first, second], _metrics).Layout!;

        Assert.Equal(Signature(expected), Signature(actual));
    }

    [Fact]
    public void Invalid_inputs_return_typed_refusals_with_positive_cases()
    {
        Assert.Equal(LaneLayoutRefusal.NullInput, LaneLayoutBuilder.TryBuild(null, _metrics).Refusal);
        Assert.Equal(LaneLayoutRefusal.EmptyInput, LaneLayoutBuilder.TryBuild([], _metrics).Refusal);
        Assert.Equal(LaneLayoutRefusal.NullMetrics, LaneLayoutBuilder.TryBuild([new LaneEventInput(Event(1, "one"), 8)], null).Refusal);
        Assert.Equal(
            LaneLayoutRefusal.InvalidMetrics,
            LaneLayoutBuilder.TryBuild([new LaneEventInput(Event(1, "one"), 8)], new(0, 0, 0, 0, 0, 0)).Refusal
        );
        Assert.Equal(LaneLayoutRefusal.NullEvent, LaneLayoutBuilder.TryBuild([null!], _metrics).Refusal);
        Assert.Equal(
            LaneLayoutRefusal.InvalidEventHeight,
            LaneLayoutBuilder.TryBuild([new LaneEventInput(Event(1, "one"), double.NaN)], _metrics).Refusal
        );
        Assert.Equal(
            LaneLayoutRefusal.InvalidEffectiveStack,
            LaneLayoutBuilder.TryBuild([new LaneEventInput(Event(1, "one"), 8, -1)], _metrics).Refusal
        );
        GanttEvent duplicate = Event(1, "same");
        Assert.Equal(
            LaneLayoutRefusal.DuplicateEventId,
            LaneLayoutBuilder.TryBuild([new LaneEventInput(duplicate, 8), new LaneEventInput(duplicate, 8)], _metrics).Refusal
        );
        Assert.Throws<ArgumentNullException>(() => LaneOrdering.LaneKey(null!));
    }

    private static GanttEvent Event(
        int row,
        string suffix,
        GanttEntityType type = GanttEntityType.AsPlannedActivity,
        int? visibleStack = null,
        DateOnly? start = null,
        DateOnly? finish = null
    ) =>
        new(
            row,
            GanttRowId.New(),
            type is GanttEntityType.Splitter or GanttEntityType.Spacer ? null : _laneId,
            visibleStack,
            type,
            null,
            start,
            finish,
            null,
            null,
            null,
            null,
            null,
            true,
            null
        );

    private static string Signature(LaneLayoutResult result) =>
        string.Join(
            "|",
            result.Lanes.Select(lane =>
                $"{lane.LaneKey}:{lane.LaneOrder}:{lane.Top:R}:{lane.Height:R}:"
                + string.Join(
                    ",",
                    lane.Slots.Select(slot =>
                        $"{slot.VisualSlotIndex}:{slot.EffectiveStackIndex}:{slot.Top:R}:{slot.Centre:R}:{slot.Bottom:R}:{slot.Height:R}:"
                        + string.Join(",", slot.EventIds.Select(id => id.Value))
                    )
                )
            )
        );
}
