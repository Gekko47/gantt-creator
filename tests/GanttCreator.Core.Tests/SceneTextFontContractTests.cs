using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests;

/// <summary>
/// The scene's font-family contract for text (R8.8 live defect, 2026-11-23): a text
/// primitive whose style named no family made the raster renderer refuse the whole export
/// with <c>FamilyMismatch</c>, because only the raster path reads the family — the sheet
/// draws shape text from the measured metrics. The scene now names it at the primitive, so
/// the property is testable without a renderer at all.
/// </summary>
public sealed class SceneTextFontContractTests
{
    [Fact]
    public void A_text_primitive_names_the_label_family_when_its_style_has_none()
    {
        SceneText text = Text(new SceneStyle("TitleStyle", fontSizePt: 9));

        Assert.Equal(GanttCatalogues.LabelFontFamily, text.Style.FontFamily);
    }

    [Fact]
    public void A_text_primitive_keeps_an_explicit_family()
    {
        SceneText text = Text(new SceneStyle("TitleStyle", fontFamily: "Consolas", fontSizePt: 9));

        Assert.Equal("Consolas", text.Style.FontFamily);
    }

    /// <summary>
    /// The fill is a rebuild through the validated constructor, so every other member must
    /// survive it — a family fix that dropped a colour or a size would be a silent visual
    /// change dressed as a contract fix.
    /// </summary>
    [Fact]
    public void Filling_the_family_preserves_every_other_style_member()
    {
        var style = new SceneStyle(
            "TitleStyle",
            ColourHex.Parse("#112233"),
            ColourHex.Parse("#445566"),
            outlineWidthPt: 1.25,
            GanttHatchPattern.ForwardDiagonal,
            fontFamily: null,
            fontSizePt: 11.5,
            bold: true,
            GanttTextAlignment.Right,
            ColourHex.Parse("#778899"),
            GanttStrokeDash.Dash);

        SceneStyle filled = style.WithLabelFontFamilyIfMissing();

        Assert.Equal(style.StyleKey, filled.StyleKey);
        Assert.Equal(style.FillColour, filled.FillColour);
        Assert.Equal(style.StrokeColour, filled.StrokeColour);
        Assert.Equal(style.OutlineWidthPt, filled.OutlineWidthPt);
        Assert.Equal(style.HatchPattern, filled.HatchPattern);
        Assert.Equal(style.FontSizePt, filled.FontSizePt);
        Assert.Equal(style.Bold, filled.Bold);
        Assert.Equal(style.Alignment, filled.Alignment);
        Assert.Equal(style.TextColour, filled.TextColour);
        Assert.Equal(style.StrokeDash, filled.StrokeDash);
        Assert.Equal(GanttCatalogues.LabelFontFamily, filled.FontFamily);
    }

    [Fact]
    public void An_explicit_family_is_returned_unchanged()
    {
        var style = new SceneStyle("TitleStyle", fontFamily: "Consolas");

        Assert.Same(style, style.WithLabelFontFamilyIfMissing());
    }

    /// <summary>
    /// The contract over a whole scene: no text primitive a renderer can receive carries a
    /// family-less style, whichever builder produced it.
    /// </summary>
    [Fact]
    public void A_scene_of_text_primitives_never_carries_a_family_less_style()
    {
        SceneText[] texts =
        [
            Text(new SceneStyle("TitleStyle", fontSizePt: 9), "chart:title-text"),
            Text(new SceneStyle("YearHeader", fontSizePt: 9), "chart:year:2026"),
            Text(new SceneStyle("PanelCell", fontFamily: "Consolas", fontSizePt: 9), "row:1:type"),
        ];

        SceneCreationOutcome creation = GanttScene.TryCreate(
            new RectD(0, 0, 100, 50),
            new RectD(0, 0, 100, 50),
            texts,
            []);

        Assert.True(creation.Succeeded, creation.Refusal?.ToString());
        Assert.All(
            creation.Scene!.Primitives.OfType<SceneText>(),
            text => Assert.False(string.IsNullOrWhiteSpace(text.Style.FontFamily)));
    }

    private static SceneText Text(SceneStyle style, string id = "chart:title-text") =>
        new(
            id,
            SceneOwnerId.Chart,
            ZLayer.Title,
            "T",
            new RectD(10, 10, 50, 10),
            style,
            GanttTextAlignment.Centre);
}
