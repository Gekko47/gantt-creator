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
    /// <returns>The panel theme <see cref="ExportBoundsResolver"/> and the panel builder consume.</returns>
    public static PanelTheme Resolve() =>
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
                outlineWidthPt: GanttCatalogues.MetricDefault("GridLinePt")));
}
