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
    internal static FrameBandsTheme FrameTheme { get; } =
        new(
            new SceneStyle("Background", fillColour: ColourHex.Parse("#FFFFFF")),
            new SceneStyle("AlternateBand", fillColour: ColourHex.Parse("#F2F2F2")),
            new SceneStyle("MinorGrid", strokeColour: ColourHex.Parse("#D9D9D9"), outlineWidthPt: 0.5),
            new SceneStyle("MajorGrid", strokeColour: ColourHex.Parse("#808080"), outlineWidthPt: 1),
            new SceneStyle("YearHeader", fillColour: ColourHex.Parse("#D9D9D9"), bold: true, alignment: GanttTextAlignment.Centre),
            new SceneStyle("PeriodHeader", fillColour: ColourHex.Parse("#F2F2F2"), alignment: GanttTextAlignment.Centre),
            new SceneStyle("Title", bold: true));

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
