using GanttCreator.Core;

namespace GanttCreator.Office;

/// <summary>The Office shape family the live renderer creates.</summary>
/// <remarks>
/// The members are the shape kinds the Phase-4 renderer needs and the probed
/// PIA can express exactly. <see cref="Polygon"/> is present because R4.5
/// renders milestone diamonds as four-point freeforms, but the freeform point
/// semantics are that row's own unprobed obligation, so this port carries the
/// kind and the points and leaves the host call to the adapter.
/// </remarks>
public enum OfficeShapeKind
{
    /// <summary>An auto shape rectangle: activity bodies, bands, grid, frames.</summary>
    Rectangle = 0,

    /// <summary>A straight line: critical intervals, grid lines, delineators.</summary>
    Line = 1,

    /// <summary>A text box: every label and text primitive.</summary>
    TextBox = 2,

    /// <summary>A freeform polygon: the milestone diamond.</summary>
    Polygon = 3,
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
/// The ordered draw-order points for <see cref="OfficeShapeKind.Polygon"/>;
/// unused by the other kinds.
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
