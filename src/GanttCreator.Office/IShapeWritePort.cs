using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office;

/// <summary>
/// One shape write requested by the live renderer: create, update, or delete
/// one owned shape by its stable scene identifier.
/// </summary>
/// <param name="PrimitiveId">
/// The stable role-derived scene primitive identifier. It becomes the shape's
/// <c>Name</c> and is the R4.7 reconciliation key (ADR-0019 D3).
/// </param>
/// <param name="Kind">The Office shape family to create.</param>
/// <param name="Geometry">The point geometry, in scene points.</param>
/// <param name="ZLayer">The scene layer, recorded for R4.5's z-order assertions.</param>
/// <param name="FillColour">The resolved fill colour, or <see langword="null"/> for none.</param>
/// <param name="StrokeColour">The resolved stroke colour, or <see langword="null"/> for none.</param>
/// <param name="LineWidthPt">The resolved outline or line width, or <see langword="null"/>.</param>
/// <param name="FontFamily">The resolved font family for a text shape, or <see langword="null"/>.</param>
/// <param name="FontSizePt">The resolved font size for a text shape, or <see langword="null"/>.</param>
/// <param name="Bold">Whether the resolved typography is bold, or <see langword="null"/>.</param>
/// <param name="Text">The text content for a text shape, or <see langword="null"/>.</param>
/// <remarks>
/// Every value is a primitive or a Core-owned type. No interop type crosses
/// this request, so the AddIn compilation never names
/// <c>Microsoft.Office.Interop.Excel</c>.
/// </remarks>
public sealed record OfficeShapeRequest(
    string PrimitiveId,
    OfficeShapeKind Kind,
    OfficeShapeGeometry Geometry,
    ZLayer ZLayer,
    ColourHex? FillColour = null,
    ColourHex? StrokeColour = null,
    double? LineWidthPt = null,
    string? FontFamily = null,
    double? FontSizePt = null,
    bool? Bold = null,
    string? Text = null);

/// <summary>
/// Narrow port over the live renderer's shape writes: create, update, and
/// delete one owned Excel shape by its stable scene identifier, and report
/// which owned shapes currently exist.
/// </summary>
/// <remarks>
/// <para>
/// The surface is deliberately primitive — no interop type crosses the port —
/// so the AddIn compilation never names <c>Microsoft.Office.Interop.Excel</c>
/// (the CS0433 duplicate-type hazard documented on
/// <see cref="IExcelApplicationAdapter"/>).
/// </para>
/// <para>
/// Every write carries the two ownership members ADR-0019 defines: the shape
/// <c>Name</c> is the full <see cref="OfficeShapeRequest.PrimitiveId"/> and the
/// shape <c>AlternativeText</c> is the bounded
/// <see cref="ShapeOwnershipTag"/> value derived from it. The adapter writes
/// both; the caller never composes either. This is what makes R4.7's
/// reconciliation key and R4.8's preservation filter possible from day one.
/// </para>
/// <para>
/// Reconciliation itself is <em>not</em> this port's job. This port performs
/// single-shape operations by identifier; R4.7 owns the algorithm that
/// compares the scene against <see cref="ListOwned"/> and decides what to
/// update, create, or delete.
/// </para>
/// </remarks>
public interface IShapeWritePort
{
    /// <summary>Creates one owned shape.</summary>
    /// <param name="request">The shape to create.</param>
    /// <returns>The typed result; on a refusal nothing was mutated.</returns>
    ShapeWriteOutcome Create(OfficeShapeRequest request);

    /// <summary>Updates one existing owned shape in place, preserving its name and tag.</summary>
    /// <param name="request">
    /// The shape to update. <see cref="OfficeShapeRequest.PrimitiveId"/> selects
    /// the existing shape; the adapter never renames a shape on update, because
    /// the name is the reconciliation key.
    /// </param>
    /// <returns>The typed result; on a refusal nothing was mutated.</returns>
    ShapeWriteOutcome Update(OfficeShapeRequest request);

    /// <summary>Deletes one owned shape by its scene identifier.</summary>
    /// <param name="primitiveId">The stable scene primitive identifier.</param>
    /// <returns>
    /// The typed result. A shape that is not found is
    /// <see cref="ShapeWriteRefusal.NotFound"/> rather than a silent success, so
    /// R4.7 can distinguish "already absent" from "deleted".
    /// </returns>
    ShapeWriteOutcome Delete(string primitiveId);

    /// <summary>Lists the scene identifiers of every owned shape on the Gantt worksheet.</summary>
    /// <returns>
    /// The owned scene identifiers in ordinal order. A shape is owned only when
    /// its alternative text carries a valid <see cref="ShapeOwnershipTag"/>; an
    /// unowned or malformed shape is excluded, never repaired here.
    /// </returns>
    IReadOnlyList<string> ListOwned();

    /// <summary>Applies z-order by moving shapes within the worksheet's z-order.</summary>
    /// <param name="backToFront">
    /// The scene identifiers in the exact back-to-front order the scene declares,
    /// lowest <see cref="ZLayer"/> first. The adapter applies the order by
    /// repeatedly moving the next shape forward, because the host exposes
    /// <c>ZOrder(MsoZOrderCmd)</c> as a relative command and
    /// <c>ZOrderPosition</c> as a read-only observation.
    /// </param>
    /// <returns>The typed result; on a refusal nothing was mutated.</returns>
    ShapeWriteOutcome ApplyZOrder(IReadOnlyList<string> backToFront);
}
