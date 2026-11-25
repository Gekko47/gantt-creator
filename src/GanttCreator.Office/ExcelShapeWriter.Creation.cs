using System.Diagnostics;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

public partial class ExcelShapeWriter
{
    /// <inheritdoc />
    public ShapeWriteOutcome Create(OfficeShapeRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);

        if (!Validate(request, out ShapeWriteRefusal? invalid))
        {
            return RefuseCreate(request, invalid!.Value, "validate");
        }

        // Refused before any host call. The name is written during ApplyOwnership,
        // which is AFTER AddShape, so without this check an over-long identifier
        // would place a shape on the sheet only to fail on the name write.
        if (NameFor(request.PrimitiveId).Length > MaxShapeNameLength)
        {
            return RefuseCreate(request, ShapeWriteRefusal.HostRejected, "name-length");
        }

        // ADR-0008 D4's first statement, before any workbook access. ADR-0054: inside a
        // reconcile batch the target was already checked when it was resolved and events are
        // off for the whole pass, so the repeated query can only return the same answer.
        if (!_batch.Active && RefusalFor(_protectionGuard.Query()) is { } guardRefusal)
        {
            return RefuseCreate(request, guardRefusal.Refusal!.Value, "guard");
        }

        Excel.Shapes? shapes = ResolveShapes(out ShapeWriteOutcome? refusal);
        if (shapes is null)
        {
            return RefuseCreate(request, refusal!.Refusal!.Value, "resolve");
        }

        long stepStart = _trace is null ? 0L : Stopwatch.GetTimestamp();

        if (FindShapeByName(shapes, NameFor(request.PrimitiveId)) is not null)
        {
            return RefuseCreate(request, ShapeWriteRefusal.AlreadyExists, "already-exists");
        }

        if (_trace is not null)
        {
            long now = Stopwatch.GetTimestamp();
            _batch.Timings.Lookup += now - stepStart;
            stepStart = now;
        }

        // R6.8 follow-up: a batch that has already created a shape with the same
        // appearance duplicates it instead of paying AddShape plus the style writes
        // again. A host that refuses the duplicate falls through to the normal create,
        // so this is an optimisation and never a new refusal path.
        ShapeTemplateKey templateKey = ShapeTemplateKey.For(request);
        if (_batch.Templates is not null
            && _batch.Templates.TryGetValue(templateKey, out Excel.Shape? template)
            && CreateFromTemplate(request, template) is { } duplicatedOutcome)
        {
            return duplicatedOutcome;
        }

        Excel.Shape? created = AddShape(shapes, request);
        if (created is null)
        {
            return RefuseCreate(request, ShapeWriteRefusal.HostRejected, "create");
        }

        // The host must not own the shape's SIZE -- except for the one class that
        // spans the plot's full height, where the host owning it IS the wanted
        // behaviour (ADR-0052 D6-D7). Excel creates shapes with
        // Placement = xlMoveAndSize, which ties the displayed height to the rows the
        // shape spans: for a bar or a label that resizes the shape behind the renderer's
        // back, so xlMove is written. For a delineator or a plot-spanning band the
        // scene's height is derived from those same rows, so the host resizing it is the
        // chart following its own rows and is kept. Measured 2026-11-03: the scene
        // declared the splitter band at 18pt on every refresh while the live band was
        // four rows tall in one state and a line in another -- which is why the
        // classification is a geometric one (full plot height) and not "all shapes".
        ApplyPlacement(created, request.HostOwnsExtent);
        if (_trace is not null)
        {
            long now = Stopwatch.GetTimestamp();
            _batch.Timings.Add += now - stepStart;
            stepStart = now;
        }

        // R4.6. The style is written on the create path for the same reason the
        // text is: it is generated output the scene owns, and a shape created in
        // the host's default accent colour with the resolved colours never
        // reaching it is the same silent-staleness class the R4.4 content write
        // closed. Order matters — the style is applied BEFORE the ownership stamp,
        // so a shape that cannot be styled is discarded by the same
        // ApplyOwnershipOrDiscard path rather than left half-written.
        //
        // The failure is reported, not swallowed. ApplyStyleOrDiscard returning
        // normally after discarding the shape would leave Create to stamp
        // ownership onto a shape that is no longer on the sheet: a write to a
        // deleted shape that either throws (surfacing as a second, misleading
        // refusal) or succeeds and returns Ok for a shape that does not exist.
        // Neither is acceptable, so a failed style write ends the create here and
        // the ownership stamp is never attempted.
        if (!ApplyStyleOrDiscard(created, request))
        {
            return RefuseCreate(request, ShapeWriteRefusal.HostRejected, "style");
        }

        if (_trace is not null)
        {
            long now = Stopwatch.GetTimestamp();
            _batch.Timings.Style += now - stepStart;
            stepStart = now;
        }

        // No per-shape z-order call here. The caller creates the scene's
        // primitives in Translate's back-to-front order and a new shape is
        // inserted in front of the existing ones, so the default insertion order
        // already IS the scene's order. ApplyZOrder is the single authoritative
        // pass, and it establishes the final order independently of whatever
        // order the host happened to be in.
        ShapeWriteOutcome stamped = ApplyOwnershipOrDiscard(created, request.PrimitiveId);
        if (_trace is not null)
        {
            _batch.Timings.Ownership += Stopwatch.GetTimestamp() - stepStart;
            _batch.Timings.Count++;
        }

        if (stamped.Succeeded)
        {
            // The first shape of this appearance is the batch's template. Only shapes
            // created in this batch are cached: a reconcile delete later in the same plan
            // could otherwise remove a template a later create still needs.
            if (_batch.Templates is not null)
            {
                _ = _batch.Templates.TryAdd(templateKey, created);
            }

            // The new shape is a valid proxy for the rest of the batch: the geometry check
            // and the z-order preflight would otherwise look it up again (ADR-0054).
            _ = Remember(NameFor(request.PrimitiveId), created);
            TraceBandShape("create", created, request.PrimitiveId, before: null);
            return stamped;
        }

        return RefuseCreate(request, stamped.Refusal!.Value, "ownership");
    }

    /// <summary>
    /// Creates one shape by duplicating a batch template with the same appearance, then
    /// writing the per-shape geometry, placement, text, and ownership identity.
    /// </summary>
    /// <param name="request">The shape to create.</param>
    /// <param name="template">A shape created earlier in this batch with the same appearance.</param>
    /// <returns>
    /// The create outcome, or <see langword="null"/> when the host did not produce a
    /// duplicate, in which case the caller creates the shape normally.
    /// </returns>
    /// <remarks>
    /// <para>
    /// <b>Measured saving (2026-11-19, Excel 16.0 x64).</b> Duplicate plus the per-shape
    /// writes costs 2.3 ms for a rectangle, 2.1 ms for a line and 8.6 ms for a text box,
    /// against 24.9 / 14.4 / 34.6 ms to create the same shapes. On a 340-shape editable
    /// copy the style block disappears almost entirely because the template already
    /// carries the resolved fill, stroke, dash, and typography.
    /// </para>
    /// <para>
    /// <b>What is deliberately not written again.</b> The template's fill, stroke, dash,
    /// and text formatting are inherited from a request with the identical appearance
    /// signature, so no style write is issued. The per-shape members - name, ownership
    /// tag, geometry, placement, and text content - are written exactly as the normal
    /// create path writes them. A failure after the duplicate exists discards it, the
    /// same rule <see cref="AddShape"/> failures follow.
    /// </para>
    /// </remarks>
    private ShapeWriteOutcome? CreateFromTemplate(OfficeShapeRequest request, Excel.Shape template)
    {
        long stepStart = _trace is null ? 0L : Stopwatch.GetTimestamp();

        Excel.Shape? duplicate;
        try
        {
            duplicate = DuplicateShape(template);
        }
        catch (Exception exception) when (exception is ArgumentException or COMException or InvalidOperationException)
        {
            // The host refused to duplicate. Fall back to the normal create rather than
            // turning an optional optimisation into an operation failure.
            duplicate = null;
        }

        if (duplicate is null)
        {
            return null;
        }

        if (!ApplyGeometry(duplicate, request.Kind, request.Geometry, request.HostOwnsExtent))
        {
            DiscardUnownedShape(duplicate);
            return RefuseCreate(request, ShapeWriteRefusal.InvalidGeometry, "geometry");
        }

        // A duplicate inherits the template's placement class, and the template key
        // includes HostOwnsExtent, so the write is skipped unless the host lost it.
        if (ReadPlacement(duplicate) != RequiredPlacement(request))
        {
            ApplyPlacement(duplicate, request.HostOwnsExtent);
        }

        if (request.Kind == OfficeShapeKind.TextBox && !ApplyTextContentOrDiscard(duplicate, request))
        {
            return RefuseCreate(request, ShapeWriteRefusal.HostRejected, "text");
        }

        if (_trace is not null)
        {
            long now = Stopwatch.GetTimestamp();
            _batch.Timings.Add += now - stepStart;
            stepStart = now;
        }

        ShapeWriteOutcome stamped = ApplyDuplicateOwnershipOrDiscard(duplicate, request.PrimitiveId);
        if (_trace is not null)
        {
            _batch.Timings.Ownership += Stopwatch.GetTimestamp() - stepStart;
            _batch.Timings.Count++;
        }

        if (!stamped.Succeeded)
        {
            return RefuseCreate(request, stamped.Refusal!.Value, "ownership");
        }

        _ = Remember(NameFor(request.PrimitiveId), duplicate);
        TraceBandShape("create", duplicate, request.PrimitiveId, before: null);
        return stamped;
    }

    /// <summary>
    /// Stamps a duplicate's ownership identity, discarding it if the host refuses.
    /// </summary>
    /// <param name="shape">The duplicate.</param>
    /// <param name="primitiveId">The stable scene primitive identifier.</param>
    /// <returns>The typed result.</returns>
    /// <remarks>
    /// Same containment as <see cref="ApplyOwnershipOrDiscard"/>: a host refusal is a
    /// typed refusal, and an identity-less duplicate must not survive on the sheet.
    /// </remarks>
    private ShapeWriteOutcome ApplyDuplicateOwnershipOrDiscard(Excel.Shape shape, string primitiveId)
    {
        // CA1031: see ApplyOwnershipOrDiscard.
#pragma warning disable CA1031
        try
        {
            ApplyDuplicateOwnershipStamp(shape, primitiveId);
            return ShapeWriteOutcome.Ok();
        }
        catch (Exception)
        {
            DiscardUnownedShape(shape);
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.HostRejected);
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Duplicates one shape. Test seam over <c>Shape.Duplicate</c>, which a Moq expression
    /// cannot substitute because it is a parameterised COM method.
    /// </summary>
    /// <param name="template">The batch template to duplicate.</param>
    /// <returns>The new shape, or <see langword="null"/> when the host produced none.</returns>
    internal virtual Excel.Shape? DuplicateShape(Excel.Shape template)
    {
        ArgumentNullException.ThrowIfNull(template);
        return template.Duplicate();
    }

}
