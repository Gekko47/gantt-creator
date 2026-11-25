using System.Runtime.InteropServices;
using GanttCreator.Core;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

public partial class ExcelShapeWriter
{
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

        int count = GetSheetCount(sheets);
        for (int index = 1; index <= count; index++)
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

        int count = GetListObjectCount(objects);
        for (int index = 1; index <= count; index++)
        {
            Excel.ListObject? table = GetListObjectAt(objects, index);
            if (table is not null && GanttChartContract.IsDataTableNameExact(table.Name))
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
    /// Resolves a set of shape names into one <c>ShapeRange</c>. Test seam over
    /// <c>Shapes.Range</c>, which takes a SAFEARRAY of names.
    /// </summary>
    /// <param name="shapes">The shapes collection.</param>
    /// <param name="names">The exact shape names, already ownership-proven by the caller.</param>
    /// <returns>The range the group style is written through.</returns>
    internal virtual Excel.ShapeRange GetShapeRange(Excel.Shapes shapes, IReadOnlyList<string> names)
    {
        ArgumentNullException.ThrowIfNull(shapes);
        ArgumentNullException.ThrowIfNull(names);

        object[] indices = new object[names.Count];
        for (int index = 0; index < names.Count; index++)
        {
            indices[index] = names[index];
        }

        return shapes.Range[indices];
    }

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

        // Inside a batch the same name is looked up by the geometry check, the update and
        // the z-order pass; one indexed lookup serves all three (ADR-0054).
        if (_batch.Proxies is not null)
        {
            if (_batch.Proxies.TryGetValue(name, out Excel.Shape? cached))
            {
                return cached;
            }

            // A complete index answers ABSENCE without the host call that raises for a
            // missing name. This is the difference between one scan per batch and one
            // COM exception per create (see EnsureBatchNameIndex).
            EnsureBatchNameIndex(shapes);
            return _batch.Proxies.TryGetValue(name, out Excel.Shape? indexed) ? indexed : null;
        }

        // The host's name indexer is the fast path: one lookup instead of a scan of
        // every shape with a COM `Name` read per shape. That scan was O(N) here and
        // this method is called once per create, update and delete, plus once per
        // shape in the ApplyZOrder preflight, so a full refresh cost O(N^2) name
        // reads (~6,400 at 80 shapes, ~10^6 at 1,000).
        //
        // A missing name RAISES (0x800A03EC) rather than returning null, which is why
        // the scan existed. Catching the miss here keeps the "null means absent"
        // contract for every caller, and a miss is rare compared with a hit.
        try
        {
            if (GetShapeByExactName(shapes, name) is { } resolved)
            {
                return Remember(name, resolved);
            }
        }
        catch (Exception exception) when (exception is ArgumentException or COMException or InvalidOperationException)
        {
            // The host reports a missing name by RAISING, and the type depends on the call
            // site. Measured on Excel 16.0 x64, 2026-11-03 (owner session `s6aea0f5a`):
            // `Shapes.Item("row:missing")` raised
            // `System.ArgumentException: The item with the specified name wasn't found.`
            // (0x80070057) — which crashed `Create`'s already-exists check before this was
            // caught. `COMException` and `InvalidOperationException` are the other
            // encodings the interop boundary produces. All three are the documented miss.
            return null;
        }

        // The index resolved nothing without raising. A miss normally raises, so this is
        // the degraded path (a test double, or a host whose indexer behaves differently);
        // the scan keeps the "null means absent" contract for every caller.
        int count = GetShapeCount(shapes);
        for (int index = 1; index <= count; index++)
        {
            Excel.Shape? shape = GetShapeAt(shapes, index);
            if (shape is not null && string.Equals(shape.Name, name, StringComparison.Ordinal))
            {
                return Remember(name, shape);
            }
        }

        return null;
    }

    /// <summary>
    /// Resolves a shape through the host's name indexer. Test seam over the
    /// parameterised COM property, which cannot be expressed in a Moq expression
    /// (CS0855); overrides model both the hit and the raising miss.
    /// </summary>
    /// <param name="shapes">The worksheet's shapes collection.</param>
    /// <param name="name">The exact shape name.</param>
    /// <returns>The shape when the host found it.</returns>
    /// <exception cref="COMException">Raised by the host when no shape carries the name.</exception>
    internal virtual Excel.Shape? GetShapeByExactName(Excel.Shapes shapes, string name)
    {
        ArgumentNullException.ThrowIfNull(shapes);
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        return shapes.Item(name);
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

                if (textBox is not null && !ApplyTextOrDiscard(textBox, request))
                {
                    return null;
                }

                return textBox;

            case OfficeShapeKind.Diamond:
                return AddDiamond(shapes, request);

            default:
                // An undefined kind is refused rather than placed. Returning null
                // makes Create report HostRejected, which is a visible refusal and
                // not a silently misplaced shape.
                return null;
        }
    }

    /// <summary>
    /// Creates the milestone diamond from the request's bounding box.
    /// </summary>
    /// <param name="shapes">The shapes collection.</param>
    /// <param name="request">The request whose bounds place the diamond.</param>
    /// <returns>The created diamond, or <see langword="null"/> when the host refused.</returns>
    /// <remarks>
    /// <para>
    /// R4.5 D1, decided by measurement rather than preference. The row was first
    /// implemented as a host freeform, and the Step-0 probe then established
    /// that <c>BuildFreeform</c> stores each vertex on a whole-EMU grid: the
    /// resulting span is <strong>+1 EMU</strong> (1 EMU = 1/12700 pt) off the
    /// scene value, measured at +1.005, +0.993, +1.017, +0.969 and +0.969 EMU
    /// across exact spans of 10, 20, 37.5, 100 and 253 pt. That is a fixed
    /// one-unit quantisation, not drift: it is constant across a 25x size range.
    /// </para>
    /// <para>
    /// A diamond auto-shape is <em>derived</em> from its
    /// <c>Left</c>/<c>Top</c>/<c>Width</c>/<c>Height</c> box rather than storing
    /// absolute vertices, and measured <strong>0.000 EMU at every span</strong>.
    /// Using it therefore removes the only reason this project had ever needed a
    /// second geometry tolerance, so <see cref="GeometryMath.Epsilon"/> governs
    /// every shape family again.
    /// </para>
    /// <para>
    /// The shape is placed from <see cref="OfficeShapeGeometry.Bounds"/> - the
    /// same member <see cref="ApplyGeometry"/> uses on the update path - rather
    /// than recomputed here from the points. The scene still models a milestone
    /// as four points and they stay on the request so the translation remains
    /// auditable, but <see cref="SceneShapeRenderer"/> has already refused any
    /// polygon that is not a symmetric axis-aligned diamond, so the two boxes are
    /// equal by construction. Recomputing them here created a second place where
    /// the box could be derived, which is exactly the dual-source risk R4.1's
    /// single rounding boundary exists to prevent.
    /// </para>
    /// </remarks>
    private static Excel.Shape? AddDiamond(Excel.Shapes shapes, OfficeShapeRequest request)
    {
        ArgumentNullException.ThrowIfNull(shapes);
        ArgumentNullException.ThrowIfNull(request);

        if (request.Geometry.Bounds is not { } bounds)
        {
            // Validate has already refused a bounds-less diamond before any host
            // call, so this arm is unreachable through Create. It is the guard
            // that keeps this method total for any future caller.
            return null;
        }

        return shapes.AddShape(
            MsoAutoShapeType.msoShapeDiamond,
            GeometryMath.SnapToDisplayPrecision(bounds.X),
            GeometryMath.SnapToDisplayPrecision(bounds.Y),
            GeometryMath.SnapToDisplayPrecision(bounds.Width),
            GeometryMath.SnapToDisplayPrecision(bounds.Height));
    }

}
