using GanttCreator.Core;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live <see cref="IShapeWritePort"/> over the Excel application object
/// supplied by the host at add-in load.
/// </summary>
/// <param name="application">
/// The Excel application object (for example <c>ExcelDnaUtil.Application</c>), or
/// <see langword="null"/> when the host supplied none (unit tests, non-Excel
/// host). A foreign object fails the interface cast and degrades to the
/// no-active-workbook refusal with no mutation.
/// </param>
/// <param name="protectionGuard">The shared read-only workbook-protection guard.</param>
/// <remarks>
/// <para>
/// ADR-0008 D4: the protection guard is the first read-only check, before any
/// COM mutation. A protected worksheet refuses the whole write.
/// </para>
/// <para>
/// COM ownership: the <c>Application</c>, <c>Workbook</c>, <c>Worksheet</c>,
/// <c>Shapes</c>, and <c>Shape</c> objects reached here are Excel-owned shared
/// roots. This adapter takes no ownership of them, never calls
/// <c>FinalReleaseComObject</c>, and force-releases nothing (the ownership
/// policy of <see cref="ExcelApplicationAdapter"/>). Every proxy is held in a
/// local and used without chained member expressions.
/// </para>
/// <para>
/// The <c>internal virtual</c> accessors isolate the Excel COM parameterised
/// properties (indexers) and the shape-creation calls. Expression trees cannot
/// contain indexed properties (CS0855), so contract tests substitute these
/// seams and every other member through Moq; the real indexer behaviour is
/// exercised by the tagged live-Office integration test.
/// </para>
/// <para>
/// ADR-0019: the adapter writes both ownership members itself. The shape
/// <c>Name</c> is the full primitive identifier and the shape
/// <c>AlternativeText</c> is the bounded <see cref="ShapeOwnershipTag"/> value.
/// A caller can never write an untagged shape through this port, which is what
/// makes R4.8's preservation guarantee structural rather than conventional.
/// </para>
/// </remarks>
public class ExcelShapeWriter(
    object? application,
    IWorksheetProtectionGuard? protectionGuard = null) : IShapeWritePort
{
    private readonly Excel.Application? _application = application as Excel.Application;
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    /// <inheritdoc />
    public ShapeWriteOutcome Create(OfficeShapeRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);

        if (!Validate(request, out ShapeWriteRefusal? invalid))
        {
            return ShapeWriteOutcome.Refused(invalid!.Value);
        }

        ProtectionGuardOutcome protection = _protectionGuard.Query();
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return ShapeWriteOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? ShapeWriteRefusal.NoActiveWorkbook
                    : ShapeWriteRefusal.TargetProtected);
        }

        Excel.Shapes? shapes = ResolveShapes(out ShapeWriteOutcome? refusal);
        if (shapes is null)
        {
            return refusal!;
        }

        if (FindShapeByName(shapes, request.PrimitiveId) is not null)
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.AlreadyExists);
        }

        Excel.Shape? created = AddShape(shapes, request);
        if (created is null)
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.HostRejected);
        }

        // Send the new shape to the back immediately. The host's z-order is a
        // total order over the sheet, so a shape created in the middle of a
        // refresh would otherwise land on top of every earlier primitive and
        // break the entity guide's back-to-front layer table. The authoritative
        // ordering is still applied once for the whole scene by ApplyZOrder; this
        // per-shape call only keeps a partially-built chart sane.
        created.ZOrder(MsoZOrderCmd.msoSendToBack);
        return ApplyOwnership(created, request.PrimitiveId);
    }

    /// <inheritdoc />
    public ShapeWriteOutcome Update(OfficeShapeRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);

        if (!Validate(request, out ShapeWriteRefusal? invalid))
        {
            return ShapeWriteOutcome.Refused(invalid!.Value);
        }

        ProtectionGuardOutcome protection = _protectionGuard.Query();
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return ShapeWriteOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? ShapeWriteRefusal.NoActiveWorkbook
                    : ShapeWriteRefusal.TargetProtected);
        }

        Excel.Shapes? shapes = ResolveShapes(out ShapeWriteOutcome? refusal);
        if (shapes is null)
        {
            return refusal!;
        }

        Excel.Shape? existing = FindShapeByName(shapes, request.PrimitiveId);
        if (existing is null)
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.NotFound);
        }

        if (!ApplyGeometry(existing, request.Kind, request.Geometry))
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.InvalidGeometry);
        }

        // The name and the tag are deliberately not rewritten on update: the name
        // is the reconciliation key R4.7 matched on, and ADR-0019 makes the pair
        // the ownership proof. Re-stamping the tag here would silently "repair" a
        // shape a user edited, which is R9.4's job to report, not R4.7's to hide.
        return ShapeWriteOutcome.Ok();
    }
    /// <inheritdoc />
    public ShapeWriteOutcome Delete(string primitiveId)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(primitiveId);

        ProtectionGuardOutcome protection = _protectionGuard.Query();
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return ShapeWriteOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? ShapeWriteRefusal.NoActiveWorkbook
                    : ShapeWriteRefusal.TargetProtected);
        }

        Excel.Shapes? shapes = ResolveShapes(out ShapeWriteOutcome? refusal);
        if (shapes is null)
        {
            return refusal!;
        }

        Excel.Shape? existing = FindShapeByName(shapes, primitiveId);
        if (existing is null)
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.NotFound);
        }

        // The ownership filter, not the name, authorises the delete. A user
        // shape that happens to share a name is never deleted (R4.8 D1).
        if (!ShapeOwnershipTag.IsOwnedTag(existing.AlternativeText))
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.NotFound);
        }

        existing.Delete();
        return ShapeWriteOutcome.Ok();
    }

    /// <inheritdoc />
    public IReadOnlyList<string> ListOwned()
    {
        Excel.Shapes? shapes = ResolveShapes(out _);
        if (shapes is null)
        {
            return [];
        }

        List<string> owned = [];
        var count = GetShapeCount(shapes);
        for (var index = 1; index <= count; index++)
        {
            Excel.Shape? shape = GetShapeAt(shapes, index);
            if (shape is null)
            {
                continue;
            }

            if (ShapeOwnershipTag.IsOwnedTag(shape.AlternativeText) && !string.IsNullOrWhiteSpace(shape.Name))
            {
                owned.Add(shape.Name);
            }
        }

        // Ordinal sort so the returned order is a property of the shape names
        // rather than of the host's z-order, which R4.5 and R4.7 both mutate.
        owned.Sort(StringComparer.Ordinal);
        return owned;
    }

    /// <inheritdoc />
    public ShapeWriteOutcome ApplyZOrder(IReadOnlyList<string> backToFront)
    {
        ArgumentNullException.ThrowIfNull(backToFront);

        ProtectionGuardOutcome protection = _protectionGuard.Query();
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return ShapeWriteOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? ShapeWriteRefusal.NoActiveWorkbook
                    : ShapeWriteRefusal.TargetProtected);
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

        // Anchor the first shape at the very back, then bring each subsequent
        // shape forward exactly once. The host's ZOrder command is relative
        // (msoSendToBack / msoBringForward) and ZOrderPosition is read-only, so
        // an absolute index cannot be assigned; this sequence is the way to make
        // the host's order match the scene's declared order.
        Excel.Shape? back = FindShapeByName(shapes, backToFront[0]);
        if (back is null)
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.NotFound);
        }

        back.ZOrder(MsoZOrderCmd.msoSendToBack);

        for (var index = 1; index < backToFront.Count; index++)
        {
            Excel.Shape? shape = FindShapeByName(shapes, backToFront[index]);
            if (shape is null)
            {
                return ShapeWriteOutcome.Refused(ShapeWriteRefusal.NotFound);
            }

            shape.ZOrder(MsoZOrderCmd.msoBringForward);
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
            OfficeShapeKind.Polygon => geometry.Points is { Count: >= 3 } points && points.All(point => IsUsablePoint(point)),
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

    /// <summary>
    /// Writes the two ADR-0019 ownership members onto a newly created shape.
    /// </summary>
    /// <param name="shape">The created shape.</param>
    /// <param name="primitiveId">The stable scene primitive identifier.</param>
    /// <returns>The typed result.</returns>
    private static ShapeWriteOutcome ApplyOwnership(Excel.Shape shape, string primitiveId)
    {
        shape.Name = primitiveId;
        shape.AlternativeText = ShapeOwnershipTag.ForPrimitiveId(primitiveId);
        return ShapeWriteOutcome.Ok();
    }

    /// <summary>
    /// Writes a geometry onto an existing shape at the single rounding boundary.
    /// </summary>
    /// <param name="shape">The shape to update.</param>
    /// <param name="kind">The shape kind, which selects which members are written.</param>
    /// <param name="geometry">The scene geometry, in points.</param>
    /// <returns><see langword="true"/> when every member was written.</returns>
    /// <remarks>
    /// <para>
    /// R4.1 D3: this is the <em>only</em> place a scene point becomes a host
    /// <c>Single</c>, via <c>GeometryMath.SnapToDisplayPrecision</c>. A caller
    /// never pre-rounds, and there is no second conversion anywhere on the path.
    /// </para>
    /// <para>
    /// A polygon is not rewritten here. Excel exposes a freeform's vertices as
    /// a read-only <c>Vertices</c> collection and the only way to change them is
    /// to rebuild through <c>BuildFreeform</c>, so an in-place polygon update is
    /// not expressible. R4.7 owns the delete-and-recreate fallback for that case;
    /// returning <see langword="false"/> here makes the limitation explicit rather
    /// than silently ignoring the requested geometry.
    /// </para>
    /// </remarks>
    private static bool ApplyGeometry(Excel.Shape shape, OfficeShapeKind kind, OfficeShapeGeometry geometry)
    {
        ArgumentNullException.ThrowIfNull(geometry);

        if (!HasRequiredGeometry(kind, geometry))
        {
            return false;
        }

        switch (kind)
        {
            case OfficeShapeKind.Rectangle:
            case OfficeShapeKind.TextBox:
                RectD bounds = geometry.Bounds!.Value;
                shape.Left = GeometryMath.SnapToDisplayPrecision(bounds.X);
                shape.Top = GeometryMath.SnapToDisplayPrecision(bounds.Y);
                shape.Width = GeometryMath.SnapToDisplayPrecision(bounds.Width);
                shape.Height = GeometryMath.SnapToDisplayPrecision(bounds.Height);
                return true;

            case OfficeShapeKind.Line:
                // Probed 2026-09-27: a created line carries its absolute endpoints,
                // but the Shape surface exposes only the bounding box, so an update
                // is expressed as left/top/width/height. A negative-width line
                // (right-to-left) is normalised here rather than rejected, because
                // the host's own bounding box cannot represent it.
                PointD from = geometry.From!.Value;
                PointD to = geometry.To!.Value;
                var left = Math.Min(from.X, to.X);
                var top = Math.Min(from.Y, to.Y);
                shape.Left = GeometryMath.SnapToDisplayPrecision(left);
                shape.Top = GeometryMath.SnapToDisplayPrecision(top);
                shape.Width = GeometryMath.SnapToDisplayPrecision(Math.Abs(to.X - from.X));
                shape.Height = GeometryMath.SnapToDisplayPrecision(Math.Abs(to.Y - from.Y));
                return true;

            case OfficeShapeKind.Polygon:
            default:
                // A freeform's vertices are a read-only collection and the only way
                // to change them is to rebuild through BuildFreeform, so an
                // in-place polygon update is not expressible. R4.7 owns the
                // delete-and-recreate fallback; returning false here makes the
                // limitation explicit rather than silently ignoring the geometry.
                return false;
        }
    }

    // ---- Test seams (internal virtual, per the ExcelWorkbookInitialiser pattern) ----

    /// <summary>
    /// Returns the worksheet carrying the Gantt table. Test seam over the COM
    /// parameterised <c>Worksheets.Item</c> property and the table lookup
    /// (CS0855).
    /// </summary>
    /// <param name="sheets">The workbook's sheet collection.</param>
    /// <returns>The Gantt worksheet, or <see langword="null"/> when absent.</returns>
    internal virtual Excel._Worksheet? FindGanttWorksheet(Excel.Sheets sheets)
    {
        ArgumentNullException.ThrowIfNull(sheets);

        var count = GetSheetCount(sheets);
        for (var index = 1; index <= count; index++)
        {
            Excel._Worksheet? sheet = GetSheetAt(sheets, index);
            if (sheet is not null && HasGanttTable(sheet))
            {
                return sheet;
            }
        }

        return null;
    }

    /// <summary>Reads the sheet count. Test seam over <c>Worksheets.Count</c>.</summary>
    /// <param name="sheets">The sheet collection.</param>
    /// <returns>The number of sheets.</returns>
    internal virtual int GetSheetCount(Excel.Sheets sheets) => sheets.Count;

    /// <summary>Reads one sheet by its 1-based index. Test seam over <c>Sheets.Item</c>.</summary>
    /// <param name="sheets">The sheet collection.</param>
    /// <param name="index">The 1-based sheet index.</param>
    /// <returns>The sheet, or <see langword="null"/> when the index does not resolve.</returns>
    internal virtual Excel._Worksheet? GetSheetAt(Excel.Sheets sheets, int index) =>
        sheets[index] as Excel._Worksheet;

    /// <summary>
    /// Determines whether a sheet carries the Gantt table. Test seam over the
    /// <c>ListObjects</c> indexed property and the <c>Name</c> read.
    /// </summary>
    /// <param name="sheet">The candidate worksheet.</param>
    /// <returns><see langword="true"/> when the sheet carries <c>tblGanttData</c>.</returns>
    internal virtual bool HasGanttTable(Excel._Worksheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);

        Excel.ListObjects? objects = sheet.ListObjects;
        if (objects is null)
        {
            return false;
        }

        var count = GetListObjectCount(objects);
        for (var index = 1; index <= count; index++)
        {
            Excel.ListObject? table = GetListObjectAt(objects, index);
            if (table is not null && string.Equals(table.Name, GanttTableSchema.TableName, StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>Reads the table count. Test seam over <c>ListObjects.Count</c>.</summary>
    /// <param name="objects">The table collection.</param>
    /// <returns>The number of tables on the sheet.</returns>
    internal virtual int GetListObjectCount(Excel.ListObjects objects) => objects.Count;

    /// <summary>Reads one table by its 1-based index. Test seam over <c>ListObjects.Item</c>.</summary>
    /// <param name="objects">The table collection.</param>
    /// <param name="index">The 1-based table index.</param>
    /// <returns>The table, or <see langword="null"/> when the index does not resolve.</returns>
    internal virtual Excel.ListObject? GetListObjectAt(Excel.ListObjects objects, int index) =>
        objects[index];

    /// <summary>Reads a worksheet's shapes collection. Test seam over the <c>_Worksheet.Shapes</c> cast.</summary>
    /// <param name="sheet">The worksheet.</param>
    /// <returns>The shapes collection, or <see langword="null"/>.</returns>
    internal virtual Excel.Shapes? GetShapes(Excel._Worksheet sheet)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        return sheet.Shapes;
    }

    /// <summary>Reads the shape count. Test seam over <c>Shapes.Count</c>.</summary>
    /// <param name="shapes">The shapes collection.</param>
    /// <returns>The number of shapes on the sheet.</returns>
    internal virtual int GetShapeCount(Excel.Shapes shapes) => shapes.Count;

    /// <summary>Reads one shape by its 1-based index. Test seam over <c>Shapes.Item</c>.</summary>
    /// <param name="shapes">The shapes collection.</param>
    /// <param name="index">The 1-based shape index.</param>
    /// <returns>The shape, or <see langword="null"/> when the index does not resolve.</returns>
    internal virtual Excel.Shape? GetShapeAt(Excel.Shapes shapes, int index) => shapes.Item(index);

    /// <summary>
    /// Finds a shape by its exact name. Test seam over the name lookup, which
    /// raises a COM error rather than returning null for a missing name; scanning
    /// by index turns that into a null so callers need no try/catch.
    /// </summary>
    /// <param name="shapes">The shapes collection.</param>
    /// <param name="name">The exact shape name.</param>
    /// <returns>The shape, or <see langword="null"/> when no shape carries the name.</returns>
    internal virtual Excel.Shape? FindShapeByName(Excel.Shapes shapes, string name)
    {
        ArgumentNullException.ThrowIfNull(shapes);
        ArgumentException.ThrowIfNullOrWhiteSpace(name);

        var count = GetShapeCount(shapes);
        for (var index = 1; index <= count; index++)
        {
            Excel.Shape? shape = GetShapeAt(shapes, index);
            if (shape is not null && string.Equals(shape.Name, name, StringComparison.Ordinal))
            {
                return shape;
            }
        }

        return null;
    }

    /// <summary>
    /// Creates one shape of the requested kind. Test seam over the three
    /// shape-creation calls and the freeform builder sequence, so contract tests
    /// assert the operation without a live Excel.
    /// </summary>
    /// <param name="shapes">The shapes collection.</param>
    /// <param name="request">The shape to create.</param>
    /// <returns>The created shape, or <see langword="null"/> when the host refused.</returns>
    internal virtual Excel.Shape? AddShape(Excel.Shapes shapes, OfficeShapeRequest request)
    {
        ArgumentNullException.ThrowIfNull(shapes);
        ArgumentNullException.ThrowIfNull(request);

        OfficeShapeGeometry geometry = request.Geometry;

        switch (request.Kind)
        {
            case OfficeShapeKind.Rectangle:
                RectD rect = geometry.Bounds!.Value;
                return shapes.AddShape(
                    MsoAutoShapeType.msoShapeRectangle,
                    GeometryMath.SnapToDisplayPrecision(rect.X),
                    GeometryMath.SnapToDisplayPrecision(rect.Y),
                    GeometryMath.SnapToDisplayPrecision(rect.Width),
                    GeometryMath.SnapToDisplayPrecision(rect.Height));

            case OfficeShapeKind.Line:
                PointD from = geometry.From!.Value;
                PointD to = geometry.To!.Value;
                return shapes.AddLine(
                    GeometryMath.SnapToDisplayPrecision(from.X),
                    GeometryMath.SnapToDisplayPrecision(from.Y),
                    GeometryMath.SnapToDisplayPrecision(to.X),
                    GeometryMath.SnapToDisplayPrecision(to.Y));

            case OfficeShapeKind.TextBox:
                RectD text = geometry.Bounds!.Value;
                Excel.Shape? textBox = shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    GeometryMath.SnapToDisplayPrecision(text.X),
                    GeometryMath.SnapToDisplayPrecision(text.Y),
                    GeometryMath.SnapToDisplayPrecision(text.Width),
                    GeometryMath.SnapToDisplayPrecision(text.Height));

                if (textBox is not null)
                {
                    ApplyText(textBox, request);
                }

                return textBox;

            case OfficeShapeKind.Polygon:
            default:
                // R4.5 owns the freeform point semantics. This row's Step-0 probe
                // established only that the API path is
                // BuildFreeform -> AddNodes -> ConvertToShape, not how the points
                // are interpreted, so the kind is recognised and refused rather
                // than implemented on an unprobed assumption. Returning null makes
                // Create report HostRejected, which is a visible refusal and not a
                // silently misplaced shape.
                return null;
        }
    }

    /// <summary>
    /// Writes the text content, typography, and alignment onto a text shape.
    /// </summary>
    /// <param name="shape">The created text box.</param>
    /// <param name="request">The request whose text members are written.</param>
    /// <remarks>
    /// <para>
    /// R4.4: the renderer <strong>consumes</strong> the scene's resolved text.
    /// The string is written verbatim, including any ellipsis R3.6's overflow
    /// policy already applied, and the resolved <c>TextBounds</c> positions the
    /// shape. Nothing here measures, truncates, re-wraps, or re-selects a label
    /// side; the scene owns all of that.
    /// </para>
    /// <para>
    /// The <c>TextFrame2</c> path is used rather than the legacy
    /// <c>TextFrame.Characters</c> path because the latter's font members are
    /// typed <c>Object</c>, and this repository treats an untyped interop member
    /// as a live hazard - R4.1's <c>Shape.Tag</c> and R4.2's
    /// <c>Application.Calculation</c> were both assumptions about interop member
    /// shapes that did not survive contact. R4.4 D3's original
    /// <c>xlAlignLeft</c>-style constants do not exist in the installed
    /// assembly; the real values are <see cref="MsoParagraphAlignment"/> members.
    /// </para>
    /// <para>
    /// Word wrap is switched off deliberately. A text box that auto-fits would
    /// resize itself away from the bounds the scene resolved, which is the same
    /// class of silent re-layout the no-remeasure rule forbids.
    /// </para>
    /// </remarks>
    private static void ApplyText(Excel.Shape shape, OfficeShapeRequest request)
    {
        ArgumentNullException.ThrowIfNull(shape);
        ArgumentNullException.ThrowIfNull(request);

        Excel.TextFrame2? frame = shape.TextFrame2;
        if (frame is null)
        {
            return;
        }

        frame.WordWrap = MsoTriState.msoFalse;
        frame.AutoSize = MsoAutoSize.msoAutoSizeNone;

        TextRange2? textRange = frame.TextRange;
        if (textRange is null)
        {
            return;
        }

        // Verbatim: the scene already applied the overflow policy.
        textRange.Text = request.Text ?? string.Empty;

        if (request.Alignment is { } alignment)
        {
            textRange.ParagraphFormat.Alignment = MapAlignment(alignment);
        }

        if (request.FontFamily is { } family)
        {
            textRange.Font.Name = family;
        }

        if (request.FontSizePt is { } size)
        {
            textRange.Font.Size = (float)size;
        }

        if (request.Bold is { } bold)
        {
            textRange.Font.Bold = bold ? MsoTriState.msoTrue : MsoTriState.msoFalse;
        }
    }

    /// <summary>
    /// Maps a resolved scene alignment onto the host's paragraph-alignment
    /// constant.
    /// </summary>
    /// <param name="alignment">The alignment the scene resolved.</param>
    /// <returns>The host alignment constant.</returns>
    /// <remarks>
    /// The mapping is closed because <see cref="GanttTextAlignment"/> is. The
    /// default arm is a throw rather than a silent fallback: an alignment added
    /// to the scene enum must fail loudly here, so the guide and this method
    /// cannot drift apart unnoticed.
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when the alignment is not a defined member.</exception>
    private static MsoParagraphAlignment MapAlignment(GanttTextAlignment alignment) =>
        alignment switch
        {
            GanttTextAlignment.Left => MsoParagraphAlignment.msoAlignLeft,
            GanttTextAlignment.Centre => MsoParagraphAlignment.msoAlignCenter,
            GanttTextAlignment.Right => MsoParagraphAlignment.msoAlignRight,
            _ => throw new ArgumentOutOfRangeException(
                nameof(alignment),
                alignment,
                "The scene produced an alignment with no host mapping. Extend MapAlignment and the guide together."),
        };
}
