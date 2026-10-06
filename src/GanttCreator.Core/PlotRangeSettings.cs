namespace GanttCreator.Core;

/// <summary>The typed settings object passed to <see cref="PlotRangeResolver.ValidateSettings"/>.</summary>
public sealed record PlotRangeSettings
{
    /// <summary>The plot-range mode text (must match a <see cref="PlotRangeMode"/> name).</summary>
    public string Mode { get; init; } = string.Empty;
    /// <summary>The explicit start date text (used only when <see cref="PlotRangeMode.Explicit"/>).</summary>
    public string? StartDate { get; init; }
    /// <summary>The explicit finish date text (used only when <see cref="PlotRangeMode.Explicit"/>).</summary>
    public string? FinishDate { get; init; }
}
