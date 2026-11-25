using GanttCreator.Core;
using GanttCreator.Core.Scene;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

public partial class ExcelShapeWriter
{
    /// <inheritdoc />
    public bool GeometryMatches(OfficeShapeRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);

        if (request.Geometry is null || !TryGetBoundingBox(request.Geometry, out RectD desired))
        {
            return false;
        }

        Excel.Shapes? shapes = ResolveShapes(out _);
        if (shapes is null)
        {
            return false;
        }

        Excel.Shape? shape = FindShapeByName(shapes, NameFor(request.PrimitiveId));
        return shape is not null
            && CarriesOwnershipTagFor(shape, request.PrimitiveId)
            && MatchesGeometry(shape, request, desired);
    }

    /// <summary>Compares one resolved shape's live box and placement with the request.</summary>
    /// <param name="shape">The resolved shape.</param>
    /// <param name="request">The desired shape.</param>
    /// <param name="desired">The request's bounding box, already derived.</param>
    /// <returns><see langword="true"/> when the host already holds it.</returns>
    /// <remarks>
    /// Placement is part of the match, and the required placement is the request's class. A
    /// scene-owned shape still on the host's default xlMoveAndSize can be resized by a
    /// row-visibility change at any moment, so skipping it would preserve exactly the drift
    /// this verification exists to catch; the rewrite re-places it with xlMove. The reverse
    /// holds for a host-owned shape: it must carry xlMoveAndSize, or the live stretch
    /// ADR-0052 D6 exists to restore is absent.
    /// </remarks>
    private static bool MatchesGeometry(Excel.Shape shape, OfficeShapeRequest request, RectD desired)
    {
        if (shape.Placement != RequiredPlacement(request))
        {
            return false;
        }

        // The host-resized extent is the one member the host is allowed to have moved for
        // a host-owned shape: Excel resizes an xlMoveAndSize shape proportionally, so a
        // collapse/expand cycle (height) or a column resize (width) can leave a fraction
        // of a point of drift that is not a scene change. The wider tolerance is stated
        // (ADR-0052 D8), not tuned.
        //
        // A host-owned LINE takes it on both axes, because a line has a near-zero extent
        // on the axis it does not span; a host-owned RECT keeps the strict width, because
        // the host never resizes a band's width and a mismatch there IS a real difference.
        // Left and top keep the strict tolerance for every class: the host moves a shape,
        // it does not reposition it within its own anchor.
        bool hostOwnsWidth = request.HostOwnsExtent && request.Kind == OfficeShapeKind.Line;
        return MatchesLive(shape.Left, desired.X)
            && MatchesLive(shape.Top, desired.Y)
            && (hostOwnsWidth
                ? MatchesLiveWithin(shape.Width, desired.Width, HostOwnedGeometry.HostResizeTolerancePt)
                : MatchesLive(shape.Width, desired.Width))
            && (request.HostOwnsExtent
                ? MatchesLiveWithin(shape.Height, desired.Height, HostOwnedGeometry.HostResizeTolerancePt)
                : MatchesLive(shape.Height, desired.Height));
    }

    /// <summary>Compares one resolved shape's live box and placement with the request.</summary>
    private static bool MatchesGeometry(Excel.Shape shape, OfficeShapeRequest request) =>
        request.Geometry is not null
        && TryGetBoundingBox(request.Geometry, out RectD desired)
        && MatchesGeometry(shape, request, desired);

    /// <summary>The tolerance for a live-versus-scene geometry comparison, in points.</summary>
    /// <remarks>
    /// The host stores geometry as <see cref="float"/> and the scene snaps to display
    /// precision at the port boundary, so exact equality would test rounding rather than
    /// geometry. One hundredth of a point is far below anything Excel can show and far
    /// above the <see cref="float"/> rounding.
    /// </remarks>
    internal const double GeometryMatchTolerancePt = 0.01;

    /// <summary>Compares one live host value to the scene's expected point value.</summary>
    private static bool MatchesLive(float actual, double expected) =>
        MatchesLiveWithin(actual, expected, GeometryMatchTolerancePt);

    /// <summary>
    /// Compares one live host value to the scene's expected point value within an
    /// explicit tolerance.
    /// </summary>
    /// <param name="actual">The host's value.</param>
    /// <param name="expected">The scene's value.</param>
    /// <param name="tolerancePt">The tolerance, in points.</param>
    /// <returns><see langword="true"/> when the values are within the tolerance.</returns>
    private static bool MatchesLiveWithin(float actual, double expected, double tolerancePt) =>
        Math.Abs(actual - GeometryMath.SnapToDisplayPrecision(expected)) <= tolerancePt;

    /// <summary>
    /// Derives the bounding box a shape of this kind occupies, exactly as
    /// <see cref="ApplyGeometry"/> writes it.
    /// </summary>
    /// <param name="geometry">The scene geometry.</param>
    /// <param name="bounds">The bounding box, when the geometry carries one.</param>
    /// <returns><see langword="true"/> when a bounding box could be derived.</returns>
    private static bool TryGetBoundingBox(OfficeShapeGeometry geometry, out RectD bounds)
    {
        if (geometry.Bounds is { } rect)
        {
            bounds = rect;
            return true;
        }

        if (geometry.From is { } from && geometry.To is { } to)
        {
            bounds = new RectD(
                Math.Min(from.X, to.X),
                Math.Min(from.Y, to.Y),
                Math.Abs(to.X - from.X),
                Math.Abs(to.Y - from.Y));
            return true;
        }

        bounds = default;
        return false;
    }

    /// <inheritdoc />
    public ShapeWriteOutcome ApplyZOrder(IReadOnlyList<string> backToFront)
    {
        ArgumentNullException.ThrowIfNull(backToFront);

        // ADR-0008 D4's first statement, before any workbook access. ADR-0054: inside a
        // reconcile batch the target was already checked when it was resolved and events are
        // off for the whole pass, so the repeated query can only return the same answer.
        if (!_batch.Active && RefusalFor(_protectionGuard.Query()) is { } guardRefusal)
        {
            return guardRefusal;
        }

        Excel.Shapes? shapes = ResolveShapes(out ShapeWriteOutcome? refusal);
        if (shapes is null)
        {
            return refusal!;
        }

        if (backToFront.Count == 0)
        {
            return ShapeWriteOutcome.Ok();
        }

        // Resolve AND prove ownership of every named shape BEFORE mutating any.
        // Resolving inside the mutation loop would let a name that does not exist
        // leave the earlier shapes already moved, which is the partial order this
        // method refuses to produce; two passes keep the refusal atomic.
        //
        // The ownership check is load-bearing and belongs in this same preflight,
        // not in the mutation loop. `Update`, `Delete`, and `ListOwned` all
        // authorise on `CarriesOwnershipTagFor` rather than on the name, so a
        // refresh preserves a user-drawn shape that happens to share a requested
        // primitive ID. Without the check here, `ApplyZOrder` was the one mutation
        // path that would still reorder exactly that shape - a preservation
        // guarantee that holds everywhere else and silently does not hold here.
        // Putting it in the preflight also means an unowned shape refused *later*
        // in the list leaves the whole scene untouched, which is the atomicity the
        // existence check above already provides.
        List<Excel.Shape> resolved = new(backToFront.Count);
        for (int index = 0; index < backToFront.Count; index++)
        {
            string primitiveId = backToFront[index];
            Excel.Shape? shape = FindShapeByName(shapes, NameFor(primitiveId));
            if (shape is null || !CarriesOwnershipTagFor(shape, primitiveId))
            {
                // NotFound rather than a new refusal: from the caller's point of view
                // "no owned shape with this identifier" is one condition, and it is the
                // reason `Update` and `Delete` already report. An unowned shape is
                // indistinguishable from an absent one by design - re-stamping its
                // tag here would "repair" it, which is R9.4's job to report, not this
                // row's.
                return ShapeWriteOutcome.Refused(ShapeWriteRefusal.NotFound);
            }

            resolved.Add(shape);
        }

        // Send each shape to the back in REVERSE order. msoSendToBack is absolute
        // ("make this the backmost shape"), so walking the scene's back-to-front
        // list from the front end and pushing every shape behind everything else
        // leaves the final order equal to backToFront whatever the host's prior
        // order was. Forward order would produce its exact reverse. This is still
        // the single deterministic pass R4.5 D2 requires: no BringToFront, and no
        // per-shape call outside this method.
        //
        // O5 (ADR-0054): only the shapes that are actually out of order are moved. A read
        // costs ~0.03-0.1 ms and a ZOrder command is a write at 0.6-2.7 ms, so reading N
        // positions to skip most of the writes is a large win, and a pass whose order is
        // already correct issues no writes at all.
        //
        // Two facts fix which end can be skipped, and getting either wrong silently leaves
        // the chart mis-stacked rather than merely slower.
        //
        // ONE: `Shape.ZOrderPosition` is the shape's index in the Shapes collection and
        // Shapes(1) is the shape at the BACK of the z-order (Microsoft, Shape.ZOrderPosition
        // remarks). So a correct back-to-front run has strictly INCREASING positions as the
        // index advances.
        //
        // TWO: msoSendToBack is absolute, so every shape this loop touches ends up behind
        // every shape it does not. The untouched shapes therefore always finish in FRONT of
        // the moved ones, which means only the FRONT-most run of the desired order may be
        // left alone: it keeps its relative order and stays above the pass. Skipping a
        // back-most run would push the touched shapes behind it, inverting exactly what was
        // skipped — the defect this comment replaces.
        //
        // So: find the longest front run already in correct relative order, and push
        // everything before it to the back, front-to-back, which is the original full pass
        // applied to a prefix.
        int moveCount = resolved.Count;
        int previousPosition = int.MaxValue;
        for (int index = resolved.Count - 1; index >= 0; index--)
        {
            int position = GetZOrderPosition(resolved[index]);
            if (position <= 0)
            {
                // The host did not report a usable position (a test double, or a shape the
                // host will not describe). Fall back to the full pass rather than trust a
                // partial order: the pass is absolute and always converges.
                moveCount = resolved.Count;
                break;
            }

            if (position >= previousPosition)
            {
                break;
            }

            previousPosition = position;
            moveCount = index;
        }

        for (int index = moveCount - 1; index >= 0; index--)
        {
            resolved[index].ZOrder(MsoZOrderCmd.msoSendToBack);
        }

        return ShapeWriteOutcome.Ok();
    }

    /// <summary>
    /// Resolves the worksheet whose <c>Shapes</c> collection the renderer owns.
    /// </summary>
    /// <param name="refusal">The refusal to return when resolution fails.</param>
    /// <returns>The shapes collection, or <see langword="null"/> with <paramref name="refusal"/> set.</returns>
    /// <remarks>
    /// The protection guard has already run in the entry method, per ADR-0008
    /// D4. The authoritative check against the <em>resolved target</em> runs here,
    /// because the Gantt worksheet is not necessarily the active one: R4.8 D3
    /// forbids the refresh path from activating another sheet, so a guard that
    /// only read the active sheet would authorise a write to an unprotected
    /// active sheet while the real target is protected.
    /// </remarks>
    private Excel.Shapes? ResolveShapes(out ShapeWriteOutcome? refusal)
    {
        // Resolved once per batch: the target cannot change while a refresh runs, because
        // ADR-0020's scope disables events for the whole pass (ADR-0054).
        if (_batch.Shapes is not null)
        {
            refusal = null;
            return _batch.Shapes;
        }

        Excel.Application? application = _application;
        if (application is null)
        {
            refusal = ShapeWriteOutcome.Refused(ShapeWriteRefusal.NoActiveWorkbook);
            return null;
        }

        Excel.Workbook? workbook = application.ActiveWorkbook;
        if (workbook is null)
        {
            refusal = ShapeWriteOutcome.Refused(ShapeWriteRefusal.NoActiveWorkbook);
            return null;
        }

        Excel.Sheets? sheets = workbook.Sheets;
        if (sheets is null)
        {
            refusal = ShapeWriteOutcome.Refused(ShapeWriteRefusal.WorksheetMissing);
            return null;
        }

        Excel._Worksheet? gantt = FindGanttWorksheet(sheets);
        if (gantt is null)
        {
            refusal = ShapeWriteOutcome.Refused(ShapeWriteRefusal.WorksheetMissing);
            return null;
        }

        // The authoritative check is against the resolved target, not the active
        // sheet: the Gantt worksheet is not necessarily the active one (R4.8 D3
        // forbids the refresh path from activating another sheet).
        ProtectionGuardOutcome target = _protectionGuard.QueryTarget(gantt);
        if (target != ProtectionGuardOutcome.NotProtected)
        {
            refusal = ShapeWriteOutcome.Refused(ShapeWriteRefusal.TargetProtected);
            return null;
        }

        Excel.Shapes? shapes = GetShapes(gantt);
        if (shapes is null)
        {
            refusal = ShapeWriteOutcome.Refused(ShapeWriteRefusal.WorksheetMissing);
            return null;
        }

        if (_batch.Active)
        {
            _batch.Shapes = shapes;
        }

        refusal = null;
        return shapes;
    }

    /// <summary>
    /// Validates a request before any host call, so a bad request never reaches Excel.
    /// </summary>
    /// <param name="request">The request to validate.</param>
    /// <param name="refusal">The refusal reason when validation fails.</param>
    /// <returns><see langword="true"/> when the request is well-formed.</returns>
    private static bool Validate(OfficeShapeRequest request, out ShapeWriteRefusal? refusal)
    {
        if (string.IsNullOrWhiteSpace(request.PrimitiveId))
        {
            refusal = ShapeWriteRefusal.BlankIdentifier;
            return false;
        }

        // The ownership scan reads a shape's tag only when its name could be one of ours
        // (ListOwned + CouldBeOwnedName, ADR-0054 O4), so a name outside those forms would
        // be written and then never recognised -- an un-updatable, un-deletable shape. The
        // form is enforced here rather than assumed, at the one place a name reaches the
        // host.
        if (!CouldBeOwnedName(request.PrimitiveId))
        {
            refusal = ShapeWriteRefusal.UnsupportedIdentifierForm;
            return false;
        }

        if (!Enum.IsDefined(request.Kind))
        {
            refusal = ShapeWriteRefusal.InvalidGeometry;
            return false;
        }

        if (!HasRequiredGeometry(request.Kind, request.Geometry))
        {
            refusal = ShapeWriteRefusal.InvalidGeometry;
            return false;
        }

        if (request.LineWidthPt is { } width && (!double.IsFinite(width) || width < 0))
        {
            refusal = ShapeWriteRefusal.InvalidGeometry;
            return false;
        }

        if (request.FontSizePt is { } size && (!double.IsFinite(size) || size <= 0))
        {
            refusal = ShapeWriteRefusal.InvalidGeometry;
            return false;
        }

        refusal = null;
        return true;
    }

    /// <summary>
    /// Determines whether a geometry carries the members its kind requires and
    /// whether every supplied value is finite.
    /// </summary>
    /// <param name="kind">The shape kind.</param>
    /// <param name="geometry">The geometry to check.</param>
    /// <returns><see langword="true"/> when the geometry is usable for the kind.</returns>
    /// <remarks>
    /// The kind-specific requirement is deliberate rather than permissive: a
    /// rectangle with no bounds cannot be positioned, and silently defaulting
    /// them would place the shape at the origin, which is the kind of silent zero
    /// the entity guide forbids.
    /// </remarks>
    private static bool HasRequiredGeometry(OfficeShapeKind kind, OfficeShapeGeometry geometry)
    {
        ArgumentNullException.ThrowIfNull(geometry);

        return kind switch
        {
            OfficeShapeKind.Rectangle or OfficeShapeKind.TextBox => IsUsableBounds(geometry.Bounds),
            OfficeShapeKind.Line => IsUsablePoint(geometry.From) && IsUsablePoint(geometry.To),

            // A diamond needs BOTH: the box places the auto-shape, and the
            // points stay on the request so the translation remains auditable.
            OfficeShapeKind.Diamond =>
                IsUsableBounds(geometry.Bounds)
                && geometry.Points is { Count: >= 3 } points
                && points.All(point => IsUsablePoint(point)),

            _ => false,
        };
    }

    private static bool IsUsableBounds(RectD? bounds) =>
        bounds is { } value
        && double.IsFinite(value.X)
        && double.IsFinite(value.Y)
        && double.IsFinite(value.Width)
        && double.IsFinite(value.Height)
        && value.Width > 0
        && value.Height > 0;

    private static bool IsUsablePoint(PointD? point) =>
        point is { } value && double.IsFinite(value.X) && double.IsFinite(value.Y);

}
