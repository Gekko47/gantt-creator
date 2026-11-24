using GanttCreator.Core.Scene;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// The shared export-composition selection rule (R7.6/R8.7): which composition an
/// export covers.
/// </summary>
/// <remarks>
/// The per-refresh export-aspect record that once accompanied this rule was withdrawn
/// with the R8.7 width/preview surface (owner ruling 2026-11-23); the Export PNG command
/// sizes itself from its own pre-flight's bounds. Only the selection authority remains,
/// shared by Copy Editable, Send to PowerPoint and Export PNG.
/// </remarks>
public sealed class ExportCompositionTests
{
    /// <summary>
    /// The selection is chart-only only on an explicit case-insensitive FALSE;
    /// everything else fails open to table-plus-chart (the catalogue default).
    /// </summary>
    [Theory]
    [InlineData(null, ExportSelection.TablePlusChart)]
    [InlineData("TRUE", ExportSelection.TablePlusChart)]
    [InlineData("true", ExportSelection.TablePlusChart)]
    [InlineData("garbage", ExportSelection.TablePlusChart)]
    [InlineData("FALSE", ExportSelection.ChartOnly)]
    [InlineData(" false ", ExportSelection.ChartOnly)]
    public void The_selection_is_chart_only_only_on_an_explicit_false(
        string? stored,
        ExportSelection expected)
    {
        Dictionary<string, string> settings = [];
        if (stored is not null)
        {
            settings[ExportComposition.ExportIncludeDataPanelSettingKey] = stored;
        }

        Assert.Equal(expected, ExportComposition.SelectExportSelection(settings));
    }

    /// <summary>Null settings are a caller defect and are refused.</summary>
    [Fact]
    public void Null_settings_are_refused()
    {
        Assert.Throws<ArgumentNullException>(
            () => ExportComposition.SelectExportSelection(null!));
    }
}
