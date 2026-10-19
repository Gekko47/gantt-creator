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

    [Theory]
    [InlineData(GanttEntityType.CustomActivity, "CustomActivity")]
    [InlineData(GanttEntityType.CustomMilestone, "CustomMilestone")]
    public void A_deleted_custom_style_is_healed_to_the_type_default(
        GanttEntityType type,
        string expectedKey)
    {
        // R5.6a: the user deleted the custom style this row was using, so the
        // cell names a key that is no longer in the registry. The registry-aware
        // planner rewrites the cell to the type's code-owned default, which
        // always resolves, so the row renders instead of carrying a reference
        // to something that no longer exists.
        var registry = new GanttStyleRegistry(
        [
            new GanttStyleDefinition(
                "SomeOtherStyle",
                new HashSet<GanttLabelPosition> { GanttLabelPosition.Auto },
                EntityColourCapability.Fill,
                GanttLabelPosition.Auto,
                FillColour: "#112233",
                HatchPattern: GanttHatchPattern.None),
        ]);

        StyleKeyWritePlan plan = StyleKeyWritePlanner.Plan(
            [Event(4, type, "DeletedCustomStyle")],
            registry);

        StyleKeyWrite write = Assert.Single(plan.Writes);
        Assert.Equal(4, write.RowNumber);
        Assert.Equal(expectedKey, write.NewStyleKey);
    }

    [Fact]
    public void A_custom_row_whose_style_still_resolves_is_left_alone()
    {
        // The heal must never overwrite a live user choice: a style that is
        // still in the registry is the user's, and the planner leaves it.
        var registry = new GanttStyleRegistry(
        [
            new GanttStyleDefinition(
                "Fancy",
                new HashSet<GanttLabelPosition> { GanttLabelPosition.Auto },
                EntityColourCapability.Fill,
                GanttLabelPosition.Auto,
                FillColour: "#112233",
                HatchPattern: GanttHatchPattern.None),
        ]);

        StyleKeyWritePlan plan = StyleKeyWritePlanner.Plan(
            [Event(4, GanttEntityType.CustomActivity, "Fancy")],
            registry);

        Assert.Empty(plan.Writes);
        Assert.Equal(0, plan.WriteCount);
    }

    [Fact]
    public void A_blank_custom_style_key_is_left_blank_without_a_registry()
    {
        // The compatibility path (no registry) cannot tell a deleted style from
        // an unchosen one, so it leaves every style-key row's cell alone. The
        // blank cell resolves to the type's code-owned default at render time,
        // so the ribbon's None item keeps meaning "use the built-in default".
        StyleKeyWritePlan plan = StyleKeyWritePlanner.Plan(
            [Event(2, GanttEntityType.CustomActivity, null)]);

        Assert.Empty(plan.Writes);
        Assert.Equal(0, plan.WriteCount);
    }

    [Fact]
    public void A_blank_custom_style_key_is_left_blank_even_with_a_registry()
    {
        // A blank cell is an unchosen style, not a dangling reference: the
        // heal only fires for a key that no longer resolves.
        var registry = new GanttStyleRegistry(
        [
            new GanttStyleDefinition(
                "Fancy",
                new HashSet<GanttLabelPosition> { GanttLabelPosition.Auto },
                EntityColourCapability.Fill,
                GanttLabelPosition.Auto,
                FillColour: "#112233",
                HatchPattern: GanttHatchPattern.None),
        ]);

        StyleKeyWritePlan plan = StyleKeyWritePlanner.Plan(
            [Event(2, GanttEntityType.CustomActivity, null)],
            registry);

        Assert.Empty(plan.Writes);
        Assert.Equal(0, plan.WriteCount);
    }
}
