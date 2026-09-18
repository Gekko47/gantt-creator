namespace GanttCreator.Core;

/// <summary>
/// A label position that a row type may permit. Blank <c>LabelPosition</c>
/// resolves to the type default or <see cref="Auto"/>; the permitted set per
/// type is defined by <see cref="EntityTypeCatalog"/>.
/// </summary>
public enum GanttLabelPosition
{
    /// <summary>No label is rendered.</summary>
    None = 0,

    /// <summary>Resolve the position automatically by collision policy.</summary>
    Auto = 1,

    /// <summary>To the left of the entity.</summary>
    Left = 2,

    /// <summary>To the right of the entity.</summary>
    Right = 3,

    /// <summary>Inside the entity body.</summary>
    Inside = 4,

    /// <summary>Above the entity.</summary>
    Above = 5,

    /// <summary>Below the entity.</summary>
    Below = 6,

    /// <summary>Top-left of the full-height line (delineators).</summary>
    TopLeft = 7,

    /// <summary>Top-right of the full-height line (delineators).</summary>
    TopRight = 8,

    /// <summary>Bottom-left of the full-height line (delineators).</summary>
    BottomLeft = 9,

    /// <summary>Bottom-right of the full-height line (delineators).</summary>
    BottomRight = 10,

    /// <summary>In the data panel to the left of the plot (Splitters).</summary>
    DataPanelLeft = 11,

    /// <summary>Centred on the plot area (Splitters).</summary>
    PlotCentre = 12,

    /// <summary>Both the data panel and the plot centre (Splitters).</summary>
    Both = 13,
}
