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
}
