namespace GanttCreator.Core.Scene;

/// <summary>The bounds a Copy Editable composition covers, translated to a zero origin.</summary>
/// <param name="ChartBounds">
/// The scene's own chart bounds: the union of the title band, the data panel (when one
/// is composed), the year and period headers, and the plot, plus the chart padding on
/// every side (entity guide section 1).
/// </param>
/// <remarks>
/// <para>
/// <see cref="Bounds"/> is the translation entity guide section 1 requires of every
/// renderer: the chart is moved so its top-left corner sits at the origin, because an
/// export composition has no worksheet to anchor to. The translation is derived from
/// <see cref="ChartBounds"/> rather than stored beside it, so the untranslated rectangle
/// and the delta cannot contradict each other.
/// </para>
/// <para>
/// <see cref="PlacementDelta"/> is the offset a renderer must carry. Section 1 is
/// explicit that the delta is part of the export contract: the outer padding sits
/// <em>inside</em> <see cref="ChartBounds"/>, so a crop computed from an untranslated
/// rectangle while ignoring the delta would lose exactly <c>ChartOuterPaddingPt</c> of
/// margin on two sides — the defect this type exists to make unrepresentable.
/// </para>
/// </remarks>
public sealed record ExportBounds(RectD ChartBounds)
{
    /// <summary>Gets the export bounds, translated so the top-left corner is the origin.</summary>
    public RectD Bounds => new(0, 0, ChartBounds.Width, ChartBounds.Height);

    /// <summary>Gets the translation delta from scene coordinates to export coordinates.</summary>
    /// <remarks>
    /// A scene point maps into the export composition as
    /// <c>(point.X - PlacementDelta.X, point.Y - PlacementDelta.Y)</c>.
    /// </remarks>
    public PointD PlacementDelta => new(ChartBounds.X, ChartBounds.Y);
}
