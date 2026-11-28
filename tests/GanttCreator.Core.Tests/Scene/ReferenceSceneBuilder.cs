using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// The resolved styles and measured geometry the reference scene is built with.
/// </summary>
/// <remarks>
/// These are fixture inputs, not product defaults: they are the values a Phase 4
/// adapter would read from a live workbook. They are declared once here so the
/// golden snapshot has a single, stated source rather than values repeated across
/// tests, and so a change to them is a deliberate, visible act.
/// </remarks>
internal static class ReferenceSceneBuilder
{
    /// <summary>Gets the named-style registry covering every Type the fixture uses.</summary>
    internal static GanttStyleRegistry StyleRegistry { get; } = new(
    [
        Style("AsPlannedActivity", "#92D050", "#375623", 8),
        Style("AsBuiltActivity", "#A9D18E", "#375623", 8),
        Style("BaselineActivity", "#D9D9D9", "#7F7F7F", 8),
        Style("CriticalInterval", "#C00000", "#000000", 8),
        Style("DelayEvent", "#FFC7CE", "#C00000", 8),
        Style("AsPlannedProcurement", "#BFBFBF", "#595959", 8),
        Style("AsBuiltProcurement", "#D0CECE", "#595959", 8),
        Style("AsPlannedMilestone", "#92D050", "#375623", 8),
        Style("AsBuiltMilestone", "#A9D18E", "#375623", 8),
        Style("CriticalMilestone", "#C00000", "#000000", 8),
    ]);

    /// <summary>Gets the chart-owned frame and band styles.</summary>
    /// <remarks>
    /// The three text-bearing styles carry typography and a text colour because a
    /// scene that names neither is not renderable: the R8.3 pinned-font raster
    /// renderer fails closed on a blank family and paints nothing for a null text
    /// colour (ADR-0070). The values are the code-owned catalogue defaults, not new
    /// literals, so the fixture states what a Phase 4 adapter would read rather than
    /// inventing a font. Geometry is unaffected: header and title boxes are measured
    /// by the injected metrics seam, never derived from these style members.
    /// </remarks>
    internal static FrameBandsTheme FrameTheme { get; } =
        new(
            new SceneStyle("Background", fillColour: ColourHex.Parse("#FFFFFF")),
            new SceneStyle("AlternateBand", fillColour: ColourHex.Parse("#F2F2F2")),
            new SceneStyle("MajorGrid", strokeColour: ColourHex.Parse("#808080"), outlineWidthPt: 1),
            new SceneStyle(
                "YearHeader",
                fillColour: ColourHex.Parse("#D9D9D9"),
                bold: true,
                alignment: GanttTextAlignment.Centre,
                fontFamily: GanttCatalogues.LabelFontFamily,
                fontSizePt: GanttCatalogues.TypographyDefault("YearFontSizePt"),
                textColour: ColourHex.Parse("#000000")),
            new SceneStyle(
                "PeriodHeader",
                fillColour: ColourHex.Parse("#F2F2F2"),
                alignment: GanttTextAlignment.Centre,
                fontFamily: GanttCatalogues.LabelFontFamily,
                fontSizePt: GanttCatalogues.TypographyDefault("HeaderFontSizePt"),
                textColour: ColourHex.Parse("#000000")),
            new SceneStyle(
                "Title",
                bold: true,
                fontFamily: GanttCatalogues.LabelFontFamily,
                fontSizePt: GanttCatalogues.TypographyDefault("TitleFontSizePt"),
                textColour: ColourHex.Parse("#000000")));

    private static GanttStyleDefinition Style(string key, string fill, string stroke, double height) =>
        new(
            key,
            new HashSet<GanttLabelPosition>
            {
                GanttLabelPosition.Inside,
                GanttLabelPosition.Left,
                GanttLabelPosition.Right,
                GanttLabelPosition.TopRight,
                GanttLabelPosition.BottomLeft,
                GanttLabelPosition.Auto,
            },
            EntityColourCapability.Fill | EntityColourCapability.Stroke,
            GanttLabelPosition.Inside,
            fill,
            stroke,
            "#000000",
            GanttHatchPattern.None,
            0,
            0,
            0.5,
            height,
            8);
}
