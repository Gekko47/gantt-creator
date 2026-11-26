using GanttCreator.Core;
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// Builds the theme dialog's opening request from what the workbook currently stores.
/// </summary>
/// <remarks>
/// This is the inverse of the dialog's own request builder: the dialog turns the user's edits into
/// a request to write, and this turns the stored catalogue into the request the dialog opens on.
/// It lived inside <c>RibbonStateService.ShowThemeSettingsDialog</c>, where it was the larger half
/// of a 187-line method and had no test of its own — so "the dialog opens showing what the chart
/// actually draws" was proven only through the form's tests.
/// </remarks>
internal static class ThemeEditRequestFactory
{
    /// <summary>Projects a successful catalogue read into the request the dialog opens with.</summary>
    /// <param name="outcome">The successful catalogue read.</param>
    /// <returns>The opening request, carrying one style entry per built-in preset.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="outcome"/> is <see langword="null"/>.</exception>
    internal static ThemeEditRequest FromWorkbook(ConfigReadOutcome outcome)
    {
        ArgumentNullException.ThrowIfNull(outcome);

        List<ThemeStyleEdit> currentStyles = [];
        Dictionary<string, double> currentMetrics = [];
        List<CustomStyleDef> currentCustomDefs = [];

        HashSet<string> builtInKeys = new HashSet<string>(StringComparer.Ordinal);
        foreach (GanttStylePreset preset in GanttCatalogues.StylePresets)
        {
            _ = builtInKeys.Add(preset.StyleKey);

            string fill = preset.FillColour;
            string stroke = preset.StrokeColour;
            string text = preset.TextColour;
            GanttStrokeDash? dash = null;

            if (outcome.Styles.TryGet(preset.StyleKey, out GanttStyleDefinition? definition)
                && definition is { HasFormatting: true })
            {
                fill = definition.FillColour ?? preset.FillColour;
                stroke = definition.StrokeColour ?? preset.StrokeColour;
                text = definition.TextColour ?? preset.TextColour;

                // The delineator's line type is the one built-in dash the
                // user may change, so the reader projects it from the row
                // rather than exact-matching it. The dialog must open on
                // the line type the chart actually draws.
                if (GanttStyleDashPolicy.IsUserEditable(preset.StyleKey))
                {
                    dash = definition.StrokeDash ?? preset.StrokeDash;
                }
            }

            currentStyles.Add(new ThemeStyleEdit(preset.StyleKey, fill, stroke, text, dash));
        }

        // Custom styles are registry entries that are not built-in presets.
        foreach (string key in outcome.Styles.StyleKeys)
        {
            if (builtInKeys.Contains(key))
            {
                continue;
            }

            if (outcome.Styles.TryGet(key, out GanttStyleDefinition? definition)
                && definition is { HasFormatting: true })
            {
                // The user row's display name is projected by the reader;
                // the key is the fallback when the row carries no name.
                currentCustomDefs.Add(new CustomStyleDef(
                    key,
                    definition.DisplayName is { Length: > 0 } name ? name : key,
                    definition.FillColour ?? "",
                    definition.StrokeColour ?? "",
                    definition.TextColour ?? "",
                    definition.HatchPattern ?? GanttHatchPattern.None,
                    definition.StandardOutlinePt,
                    definition.ActivityHeightPt,
                    definition.MilestoneSizePt,
                    definition.DefaultLabelPosition ?? GanttLabelPosition.Auto,
                    definition.AllowedLabelPositions,
                    definition.StrokeDash ?? GanttStrokeDash.Solid));
            }
        }

        foreach (GanttMetricToken token in GanttCatalogues.Metrics)
        {
            if (ThemeEditRequest.EditableMetrics.Contains(token.Name)
                && outcome.Metrics.TryGetValue(token.Name, out double parsed)
                && double.IsFinite(parsed))
            {
                currentMetrics[token.Name] = parsed;
            }
            else if (ThemeEditRequest.EditableMetrics.Contains(token.Name))
            {
                currentMetrics[token.Name] = token.DefaultValue;
            }
        }

        // The chart-chrome colours: the workbook's current values, so the
        // dialog opens showing what the chart actually renders.
        Dictionary<string, string> currentChromeColours = [];
        foreach (GanttChartChromeRole role in GanttChartChrome.Roles)
        {
            string key = GanttChartChrome.SettingKeyFor(role);
            currentChromeColours[key] = outcome.Settings.TryGetValue(key, out string? stored)
                && GanttColourToken.IsValidHex(stored)
                ? stored
                : GanttChartChrome.DefaultHexFor(role);
        }

        return new ThemeEditRequest(
            currentStyles,
            currentMetrics,
            null,
            currentCustomDefs,
            currentChromeColours);
    }
}
