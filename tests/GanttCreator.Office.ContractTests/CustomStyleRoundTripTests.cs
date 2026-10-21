using GanttCreator.Core;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// The write to read round trip for a user-created custom style.
/// </summary>
/// <remarks>
/// R5.6a: the ribbon's custom-style dropdown projects its items from the
/// configuration registry, so a created style is only reachable when the
/// write is visible to a later read. This class runs the real production
/// <c>WriteTheme</c> and <c>Read</c> over the shared config-sheet fake, with
/// only the COM seams substituted.
/// </remarks>
public sealed class CustomStyleRoundTripTests
{
    [Fact]
    public void A_created_custom_style_is_visible_to_a_later_read()
    {
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();

        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            [new CustomStyleEdit(
                CustomStyleOperation.Create,
                "UserStyle",
                DisplayName: "User Style",
                FillColour: "#112233",
                StrokeColour: "#445566",
                TextColour: "#FFFFFF")]);

        var outcome = ConfigGraph.BuildWriter(fake).WriteTheme(request);
        Assert.Equal(ConfigWriteOutcome.Ok(), outcome);

        ConfigReadOutcome read = ConfigGraph.BuildReader(fake).Read();

        Assert.True(read.Succeeded, $"The catalogue read refused: {read.Refusal}.");
        Assert.NotNull(read.Styles);
        Assert.Contains("UserStyle", read.Styles.StyleKeys);
    }

    /// <summary>
    /// The bar height a custom style names survives the write and the read.
    /// </summary>
    /// <remarks>
    /// WI-A: the custom style form's bar height is only a real capability if it
    /// reaches the workbook and comes back. The round trip was previously
    /// covered for the style's EXISTENCE and for nothing else, so the size
    /// columns -- the ones the renderer reads its bar height from -- were
    /// never proven to survive the write. This test writes a height of 14pt
    /// and reads it back from the registry's own definition, which is the
    /// value the scene resolves the bar from.
    /// </remarks>
    [Fact]
    public void A_custom_style_bar_height_survives_the_write_and_the_read()
    {
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();

        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            [new CustomStyleEdit(
                CustomStyleOperation.Create,
                "UserStyle",
                DisplayName: "User Style",
                FillColour: "#112233",
                StrokeColour: "#445566",
                TextColour: "#FFFFFF",
                ActivityHeightPt: 14)]);

        var outcome = ConfigGraph.BuildWriter(fake).WriteTheme(request);
        Assert.Equal(ConfigWriteOutcome.Ok(), outcome);

        ConfigReadOutcome read = ConfigGraph.BuildReader(fake).Read();

        Assert.True(read.Succeeded, $"The catalogue read refused: {read.Refusal}.");
        Assert.NotNull(read.Styles);
        Assert.True(read.Styles.TryGet("UserStyle", out var definition));
        Assert.NotNull(definition);
        Assert.Equal(14, definition.ActivityHeightPt, 3);
    }

    /// <summary>
    /// A custom style that names NO bar height keeps the zero, so the renderer
    /// inherits the height the row's own kind carries.
    /// </summary>
    /// <remarks>
    /// The counterpart of the round trip above: a zero must NOT be silently
    /// replaced by the code-owned default on the way through the workbook,
    /// because a zero is the style contract's "not applicable to this entity
    /// kind" value and the renderer reads it as "inherit".
    /// </remarks>
    [Fact]
    public void A_custom_style_with_no_bar_height_keeps_its_zero_through_the_round_trip()
    {
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();

        var request = new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            [new CustomStyleEdit(
                CustomStyleOperation.Create,
                "UserStyle",
                DisplayName: "User Style",
                FillColour: "#112233",
                StrokeColour: "#445566",
                TextColour: "#FFFFFF",
                ActivityHeightPt: 0)]);

        var outcome = ConfigGraph.BuildWriter(fake).WriteTheme(request);
        Assert.Equal(ConfigWriteOutcome.Ok(), outcome);

        ConfigReadOutcome read = ConfigGraph.BuildReader(fake).Read();

        Assert.True(read.Succeeded, $"The catalogue read refused: {read.Refusal}.");
        Assert.NotNull(read.Styles);
        Assert.True(read.Styles.TryGet("UserStyle", out var definition));
        Assert.NotNull(definition);
        Assert.Equal(0, definition.ActivityHeightPt, 3);
    }
}
