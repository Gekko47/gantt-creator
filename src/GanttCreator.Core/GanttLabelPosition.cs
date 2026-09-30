namespace GanttCreator.Core;

/// <summary>
/// A label position that a row type may permit. Blank <c>LabelPosition</c>
/// resolves to the type default or <see cref="Auto"/>; the permitted set per
/// type is defined by <see cref="EntityTypeCatalog"/>.
/// </summary>
/// <remarks>
/// <para>
/// <c>Above</c> and <c>Below</c> were removed by owner ruling 2026-09-30: they
/// are not used in any permitted format, and a vertical label placement has no
/// guaranteed room once lane geometry is fixed (ADR-0026). Removing the members
/// rather than merely denying them makes an unresolvable stored value
/// impossible to construct — ADR-0029 D6's "reported, never coerced" then has
/// nothing to coerce, because the parse fails at the name.
/// </para>
/// <para>
/// The former <c>Above</c>/<c>Below</c> members were removed and the remaining
/// values were <b>renumbered contiguously</b> rather than left as a 5/6 gap. The
/// enum is stored by NAME — <c>GanttTypeCatalogueRow.AllowedLabelPositions</c>
/// serialises <c>position.ToString()</c> — so no workbook carries a number and
/// renumbering cannot orphan one. A retained gap would have meant a marker
/// member that nothing consumes, which CA1700 correctly rejects.
/// </para>
/// </remarks>
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

    /// <summary>Top-left of the full-height line (delineators).</summary>
    TopLeft = 5,

    /// <summary>Top-right of the full-height line (delineators).</summary>
    TopRight = 6,

    /// <summary>Bottom-left of the full-height line (delineators).</summary>
    BottomLeft = 7,

    /// <summary>Bottom-right of the full-height line (delineators).</summary>
    BottomRight = 8,

    /// <summary>In the data panel to the left of the plot (Splitters).</summary>
    DataPanelLeft = 9,

    /// <summary>Centred on the plot area (Splitters).</summary>
    PlotCentre = 10,

    /// <summary>Both the data panel and the plot centre (Splitters).</summary>
    Both = 11,
}
