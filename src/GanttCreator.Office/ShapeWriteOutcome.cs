using GanttCreator.Core;

namespace GanttCreator.Office;

/// <summary>The Office shape family the live renderer creates.</summary>
/// <remarks>
/// The members are the shape kinds the Phase-4 renderer needs and the probed
/// PIA can express exactly. <see cref="Diamond"/> was originally a freeform
/// <c>Polygon</c>, on R4.5's unprobed assumption that a milestone had to be a
/// four-point freeform; measurement later showed a host freeform quantises each
/// vertex to a whole EMU while a diamond auto-shape is derived from its bounding
/// box and is exact, so the member is now a diamond and the port carries its box.
/// </remarks>
public enum OfficeShapeKind
{
    /// <summary>An auto shape rectangle: activity bodies, bands, grid, frames.</summary>
    Rectangle = 0,

    /// <summary>A straight line: critical intervals, grid lines, delineators.</summary>
    Line = 1,

    /// <summary>A text box: every label and text primitive.</summary>
    TextBox = 2,

    /// <summary>
    /// A diamond auto-shape: the milestone marker, placed at the scene
    /// polygon's bounding box.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This member was originally a freeform polygon. Measurement changed it: a
    /// host freeform stores each vertex on a whole-EMU grid, so its span lands
    /// +1 EMU (1/12700 pt) off the scene value, whereas a diamond auto-shape is
    /// derived from its bounding box and measured exactly (0.000 EMU across
    /// exact spans of 10, 20, 37.5, 100 and 253 pt). The diamond form therefore
    /// needs no tolerance of its own, so <c>GeometryMath.Epsilon</c> governs
    /// every family.
    /// </para>
    /// <para>
    /// The scene still models a milestone as a four-point <c>ScenePolygon</c>;
    /// the renderer verifies the four points form a symmetric axis-aligned
    /// diamond and refuses them otherwise rather than approximating.
    /// </para>
    /// </remarks>
    Diamond = 3,
}

/// <summary>Point geometry for one shape, in points.</summary>
/// <param name="Bounds">
/// The rectangle bounds for <see cref="OfficeShapeKind.Rectangle"/> and
/// <see cref="OfficeShapeKind.TextBox"/>; unused by the other kinds.
/// </param>
/// <param name="From">
/// The start point for <see cref="OfficeShapeKind.Line"/>; unused by the
/// other kinds.
/// </param>
/// <param name="To">
/// The end point for <see cref="OfficeShapeKind.Line"/>; unused by the other
/// kinds.
/// </param>
/// <param name="Points">
/// The ordered draw-order points for a <see cref="OfficeShapeKind.Diamond"/>,
/// which the adapter uses only to place the auto-shape; unused by the other
/// kinds.
/// </param>
/// <remarks>
/// Values are scene points and are converted to the host's <c>Single</c>
/// precision at the single port boundary by
/// <c>GeometryMath.SnapToDisplayPrecision</c>. A caller never pre-rounds.
/// </remarks>
public sealed record OfficeShapeGeometry(
    RectD? Bounds = null,
    PointD? From = null,
    PointD? To = null,
    IReadOnlyList<PointD>? Points = null);

/// <summary>Why a shape-write operation refused without mutating the workbook.</summary>
public enum ShapeWriteRefusal
{
    /// <summary>The request or one of its members was null.</summary>
    NullRequest = 0,

    /// <summary>The application object or active workbook was absent.</summary>
    NoActiveWorkbook = 1,

    /// <summary>The Gantt worksheet was not found on the active workbook.</summary>
    WorksheetMissing = 2,

    /// <summary>The target worksheet is protected; the write refused before any COM mutation.</summary>
    TargetProtected = 3,

    /// <summary>A shape with the requested name already exists.</summary>
    AlreadyExists = 4,

    /// <summary>No shape with the requested name exists.</summary>
    NotFound = 5,

    /// <summary>The geometry did not carry the members its kind requires, or carried non-finite values.</summary>
    InvalidGeometry = 6,

    /// <summary>The primitive identifier was blank.</summary>
    BlankIdentifier = 7,

    /// <summary>The ownership tag was blank or did not carry the add-in prefix.</summary>
    InvalidOwnershipTag = 8,

    /// <summary>The host refused or rejected the write.</summary>
    HostRejected = 9,
}

/// <summary>The typed result of one shape-write operation.</summary>
/// <param name="Refusal">The refusal reason, or <see langword="null"/> on success.</param>
public sealed record ShapeWriteOutcome(ShapeWriteRefusal? Refusal)
{
    /// <summary>Gets whether the write completed.</summary>
    public bool Succeeded => Refusal is null;

    /// <summary>Creates a successful outcome.</summary>
    /// <returns>The successful outcome.</returns>
    public static ShapeWriteOutcome Ok() => new((ShapeWriteRefusal?)null);

    /// <summary>Creates a refusal.</summary>
    /// <param name="refusal">The refusal reason.</param>
    /// <returns>The refusal outcome.</returns>
    public static ShapeWriteOutcome Refused(ShapeWriteRefusal refusal) => new(refusal);
}
