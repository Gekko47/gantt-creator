namespace GanttCreator.Core.Scene;

/// <summary>What a Copy Editable composition includes beside the chart.</summary>
/// <remarks>
/// <para>
/// Entity guide section 1 owns the two-mode contract: an export is either
/// "chart-only" or "data-plus-chart". These two members are the whole domain, and a
/// third mode is a stop condition rather than an extension point — the guide would
/// have to be amended first.
/// </para>
/// <para>
/// The selection adds no second composition authority. It only decides whether the
/// export <see cref="SceneBuildRequest"/> supplies a <see cref="PanelTheme"/>: the
/// chart-only composition leaves it <see langword="null"/> and the table-plus-chart
/// composition sets it. <see cref="SceneBuilder"/> already refuses a live request that
/// carries a panel, so a drawn panel can only ever reach an export profile.
/// </para>
/// </remarks>
public enum ExportSelection
{
    /// <summary>The chart alone; the composition draws no data panel.</summary>
    ChartOnly = 0,

    /// <summary>The data table beside the chart, drawn as the section 3 panel.</summary>
    TablePlusChart = 1,
}
