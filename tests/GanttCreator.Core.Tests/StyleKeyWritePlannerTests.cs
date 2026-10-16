namespace GanttCreator.Core.Tests;

/// <summary>
/// <see cref="StyleKeyWritePlanner"/> brings column <c>K</c> (<c>StyleKey</c>) into
/// line with each row's current <c>Type</c> on every Refresh: a type change must
/// update the style key accordingly.
/// </summary>
public sealed class StyleKeyWritePlannerTests
{
    private static readonly GanttRowId RowId = GanttRowId.Parse("G-" + new string('0', 32));

    private static GanttEvent Event(int rowNumber, GanttEntityType type, string? styleKey) =>
        new(
            rowNumber,
            RowId,
            null,
            0,
            type,
            $"Row {rowNumber}",
            null,
            null,
            null,
            styleKey,
            null,
            null,
            null,
            true,
            null);

    [Fact]
    public void A_blank_style_key_resolves_to_the_type_default()
    {
        StyleKeyWritePlan plan = StyleKeyWritePlanner.Plan(
            [Event(2, GanttEntityType.AsPlannedActivity, null)]);

        StyleKeyWrite write = Assert.Single(plan.Writes);
        Assert.Equal(2, write.RowNumber);
        Assert.Equal("AsPlannedActivity", write.NewStyleKey);
        Assert.Equal(1, plan.WriteCount);
    }

    [Fact]
    public void A_type_change_rewrites_the_style_key_to_the_new_type_default()
    {
        // The row was stamped as As-Built when inserted, but its Type cell now says
        // As-Planned. Refresh must not leave the stale key in place.
        StyleKeyWritePlan plan = StyleKeyWritePlanner.Plan(
            [Event(2, GanttEntityType.AsPlannedActivity, "AsBuiltActivity")]);

        StyleKeyWrite write = Assert.Single(plan.Writes);
        Assert.Equal(2, write.RowNumber);
        Assert.Equal("AsPlannedActivity", write.NewStyleKey);
    }

    [Theory]
    [InlineData(GanttEntityType.AsPlannedActivity, "AsPlannedActivity")]
    [InlineData(GanttEntityType.AsBuiltActivity, "AsBuiltActivity")]
    [InlineData(GanttEntityType.Delineator, "DefaultDelineator")]
    [InlineData(GanttEntityType.Splitter, "Splitter")]
    public void Every_type_writes_its_own_default_when_the_cell_is_blank(
        GanttEntityType type,
        string expectedKey)
    {
        StyleKeyWritePlan plan = StyleKeyWritePlanner.Plan([Event(3, type, null)]);

        StyleKeyWrite write = Assert.Single(plan.Writes);
        Assert.Equal(3, write.RowNumber);
        Assert.Equal(expectedKey, write.NewStyleKey);
    }

    [Fact]
    public void A_row_already_carrying_the_default_produces_no_write()
    {
        StyleKeyWritePlan plan = StyleKeyWritePlanner.Plan(
            [Event(2, GanttEntityType.AsPlannedActivity, "AsPlannedActivity")]);

        Assert.Empty(plan.Writes);
        Assert.Equal(0, plan.WriteCount);
    }

    [Fact]
    public void Custom_activity_is_left_alone_because_it_has_no_default()
    {
        // Custom Activity requires an explicit, user-chosen style key: there is no
        // default to sync to, so the cell is left exactly as the user left it.
        StyleKeyWritePlan plan = StyleKeyWritePlanner.Plan(
            [Event(2, GanttEntityType.CustomActivity, "Fancy")]);

        Assert.Empty(plan.Writes);
        Assert.Equal(0, plan.WriteCount);
    }

    [Fact]
    public void Writes_are_ordered_by_row_number_regardless_of_input_order()
    {
        StyleKeyWritePlan plan = StyleKeyWritePlanner.Plan(
        [
            Event(5, GanttEntityType.AsPlannedActivity, null),
            Event(2, GanttEntityType.AsBuiltActivity, null),
            Event(4, GanttEntityType.Delineator, null),
        ]);

        Assert.Equal([2, 4, 5], plan.Writes.Select(write => write.RowNumber));
    }

    [Fact]
    public void An_uncatalogued_type_throws_rather_than_writing_a_wrong_key()
    {
        // Undefined behaviour is an unknown: the planner must not invent a key for
        // a type the catalogue does not define.
        Assert.Throws<ArgumentOutOfRangeException>(() => StyleKeyWritePlanner.Plan(
            [Event(2, (GanttEntityType)99, null)]));
    }

    [Fact]
    public void Plan_refuses_null_events()
    {
        Assert.Throws<ArgumentNullException>(() => StyleKeyWritePlanner.Plan(null!));
    }
}
