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
/// <param name="HatchPattern">
/// The resolved hatch pattern, or <see cref="GanttHatchPattern.None"/> for a solid
/// fill. R4.6: the scene's own enum crosses the port so the Core token table and
/// the host mapping cannot drift; the adapter maps it to
/// <c>MsoPatternType</c> and never invents a pattern.
/// </param>
/// <param name="Bold">Whether the resolved typography is bold, or <see langword="null"/>.</param>
/// <param name="Text">The text content for a text shape, or <see langword="null"/>.</param>
/// <param name="Alignment">
/// The resolved horizontal text alignment, or <see langword="null"/> for a
/// shape that carries no text. The scene resolves this (ADR-0018 D4); the
/// adapter only maps it to the host's alignment constant.
/// </param>
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
    GanttHatchPattern HatchPattern = GanttHatchPattern.None,
    string? FontFamily = null,
    double? FontSizePt = null,
    bool? Bold = null,
    string? Text = null,
    GanttTextAlignment? Alignment = null);

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
    /// <remarks>
    /// <para>
    /// An update re-applies the <strong>geometry of every kind</strong> and, for
    /// a text box, its <strong>content</strong> — the string, the resolved font
    /// tokens, and the resolved alignment. Those are owned generated data, so a
    /// shape that must match a new scene has to be able to say the new thing, not
    /// merely move to the new place.
    /// </para>
    /// <para>
    /// Two members are deliberately NOT rewritten. <see cref="OfficeShapeRequest.PrimitiveId"/>
    /// is the reconciliation key this method's caller matched on, and the
    /// alternative-text ownership tag (ADR-0019) is the ownership proof;
    /// re-stamping either would silently "repair" a shape a user edited, which
    /// is R9.4's job to report.
    /// </para>
    /// <para>
    /// That same reasoning is why the update is <strong>authorised by the
    /// ownership tag, not by the name</strong>, exactly as
    /// <see cref="Delete"/> is: a shape whose alternative text does not carry a
    /// valid <see cref="ShapeOwnershipTag"/> is refused as
    /// <see cref="ShapeWriteRefusal.NotFound"/> and never touched, so a refresh cannot
    /// move or rewrite a user's shape that happens to share an identifier (R4.8).
    /// </para>
    /// <para>
    /// <strong>Not yet applied by an update:</strong> the fill, stroke, and
    /// pattern members, which R4.6 owns. Until that row lands, a refresh that
    /// changes only a colour would move the shape and leave the old colour — the
    /// same staleness class the text path avoids. R4.6 must close this, and this
    /// note is what tells it to.
    /// </para>
    /// </remarks>
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
    /// lowest <see cref="ZLayer"/> first.
    /// </param>
    /// <returns>The typed result; on a refusal nothing was mutated.</returns>
    /// <remarks>
    /// <para>
    /// The host exposes <c>ZOrder(MsoZOrderCmd)</c> as a relative command and
    /// <c>ZOrderPosition</c> as a read-only observation, so the order is applied in one
    /// deterministic pass that sends each shape to the back <strong>in reverse list
    /// order</strong>. <c>msoSendToBack</c> is absolute, so walking the scene's
    /// back-to-front list from the front end leaves the final order equal to the scene's
    /// whatever the host's prior order was; forward order would produce its exact
    /// reverse. This is the single pass R4.5 D2 requires: no per-shape
    /// <c>BringToFront</c>, and no z-order call outside this method.
    /// </para>
    /// <para>
    /// <strong>Preflight, then commit.</strong> Every identifier is resolved and proved
    /// owned <em>before</em> the first command is issued. A shape that is absent, or that
    /// is present but does not carry the ownership tag for its identifier, refuses the
    /// <strong>entire list</strong> as <see cref="ShapeWriteRefusal.NotFound"/> with zero
    /// <c>ZOrder</c> calls — the same reason <see cref="Update"/> and
    /// <see cref="Delete"/> report, and deliberately indistinguishable from an absent
    /// shape, because re-stamping an unowned shape's tag would silently "repair" a
    /// user-drawn shape that is R9.4's job to report. Ownership is checked in the same
    /// preflight as existence so that an unowned shape late in the list cannot leave the
    /// earlier shapes already moved.
    /// </para>
    /// </remarks>
    ShapeWriteOutcome ApplyZOrder(IReadOnlyList<string> backToFront);
}
