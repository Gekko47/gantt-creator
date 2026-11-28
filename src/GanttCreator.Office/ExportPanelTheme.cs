using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office;

/// <summary>
/// Resolves the editable-export data panel's style set from the code-owned catalogue
/// (R6.3), so the staged panel and the live table cannot drift apart.
/// </summary>
/// <remarks>
/// <para>
/// <b>One authority.</b> Every value is read through <see cref="GanttTableAppearance"/> or
/// <see cref="GanttCatalogues"/>, the same authorities the live table presentation and
/// the chart's own export path read. No hex value or point size is restated here: a
/// second literal would be a second theme.
/// </para>
/// <para>
/// The panel is code-owned presentation (ADR-0060): a workbook's chart-chrome settings
/// change the chart's frame and bands, not the exported table's own cells, exactly as
/// Initialise and Repair write the live table's presentation from these tokens.
/// </para>
/// </remarks>
public static class ExportPanelTheme
{
    /// <summary>Builds the resolved export panel style set.</summary>
    /// <param name="themedMetrics">
    /// The workbook's themed metrics, or <see langword="null"/> to use the catalogue defaults.
    /// The panel's border WIDTH comes from <c>MajorBoundaryPt</c> (owner ruling 2026-11-27,
    /// superseding ADR-0068 D5's use of <c>GridLinePt</c>): it is the width every border on
    /// the chart takes, and the grid metric it used to read no longer exists.
    /// </param>
    /// <returns>The panel theme <see cref="ExportBoundsResolver"/> and the panel builder consume.</returns>
    public static PanelTheme Resolve(IReadOnlyDictionary<string, double>? themedMetrics = null) =>
        new(
            new SceneStyle("DataPanelFill", fillColour: GanttTableAppearance.BodyFill),
            new SceneStyle(
                "DefaultText",
                fontFamily: GanttTableAppearance.FontFamily,
                fontSizePt: GanttTableAppearance.BodyFontSizePt,
                textColour: GanttTableAppearance.BodyTextColour),
            new SceneStyle("HeaderFill", fillColour: GanttTableAppearance.HeaderFill),
            new SceneStyle(
                "HeaderFontSizePt",
                fontFamily: GanttTableAppearance.FontFamily,
                fontSizePt: GanttTableAppearance.HeaderFontSizePt,
                bold: GanttTableAppearance.HeaderIsBold,
                textColour: GanttTableAppearance.HeaderTextColour),
            new SceneStyle(
                "Border",
                strokeColour: GanttTableAppearance.BorderColour,
                outlineWidthPt: BorderWidthFrom(themedMetrics)));

    /// <summary>The panel border width: the workbook's major-boundary metric, or the default.</summary>
    /// <param name="themedMetrics">The workbook's themed metrics, or <see langword="null"/>.</param>
    /// <returns>The width in points.</returns>
    /// <remarks>
    /// A missing or unusable value falls back to the catalogue default rather than refusing:
    /// the border is presentation, and a panel with no border is worse than one at the
    /// approved weight. The value is never zero, because a zero-width outline would delete the
    /// shared cell edges the panel exists to reproduce.
    /// </remarks>
    private static double BorderWidthFrom(IReadOnlyDictionary<string, double>? themedMetrics) =>
        themedMetrics is not null
        && themedMetrics.TryGetValue("MajorBoundaryPt", out double majorBoundaryPt)
        && double.IsFinite(majorBoundaryPt)
        && majorBoundaryPt > 0
            ? majorBoundaryPt
            : GanttCatalogues.MetricDefault("MajorBoundaryPt");
}
