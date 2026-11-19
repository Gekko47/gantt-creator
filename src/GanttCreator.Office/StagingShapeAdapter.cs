using System.Runtime.InteropServices;
using GanttCreator.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Shared target resolution and staging ownership proof for the editable-export adapters
/// (R6.4 grouping, R6.6 copy).
/// </summary>
/// <remarks>
/// <para>
/// <b>One authority for the two questions every staging adapter asks.</b> The Gantt worksheet
/// is the one carrying <c>tblGanttData</c> (the user may rename the sheet, the table name is
/// the schema's identity), the protection guard is queried against that resolved target -
/// never the active sheet (ADR-0008 D4, R4.8 D3) - and a staged shape is owned only when its
/// alternative text equals <see cref="ShapeStagingTag.ForStagedName"/> recomputed from its
/// live name. Both adapters consume these members; neither restates the rules.
/// </para>
/// <para>
/// <b>Why a base class rather than a composed helper.</b> The adapters' host access is
/// exposed as <c>internal virtual</c> test seams (the <see cref="ExcelShapeWriter"/>
/// pattern), and the contract tests subclass the adapters to substitute those seams. A
/// composed locator would be constructed inside the adapter and would bypass a subclass's
/// overrides; inheritance keeps every seam virtual on the instance the test subclasses, so
/// the R6.4 adapter tests keep working with no edits - the proof this extraction is pure.
/// </para>
/// <para>
/// <b>COM ownership.</b> The <c>Application</c>, <c>Workbook</c>, <c>Worksheet</c>,
/// <c>Shapes</c> and <c>Shape</c> objects reached here are Excel-owned shared roots. This
/// class takes no ownership of them, never calls <c>FinalReleaseComObject</c>, and every
/// proxy is held in a local and used without chained member expressions.
/// </para>
/// </remarks>
/// <param name="application">
/// The Excel application object (for example <c>ExcelDnaUtil.Application</c>), or
/// <see langword="null"/> when the host supplied none (unit tests, non-Excel host). A
/// foreign object fails the interface cast and degrades to the no-active-workbook refusal
/// with no mutation.
/// </param>
/// <param name="protectionGuard">The shared read-only workbook-protection guard.</param>
public abstract class StagingShapeAdapter(
    object? application,
    IWorksheetProtectionGuard? protectionGuard = null)
{
    private readonly Excel.Application? _application = application as Excel.Application;
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    /// <summary>
    /// ADR-0008 D4's first host check: queries the protection guard before any workbook
    /// access. The authoritative per-target check runs again in <see cref="ResolveShapes"/>
    /// against the resolved worksheet.
    /// </summary>
    /// <returns>
    /// <see cref="ShapeWriteRefusal.TargetProtected"/> when the active workbook is protected,
    /// otherwise <see langword="null"/>.
    /// </returns>
    internal ShapeWriteRefusal? CheckEntryGuard() =>
        _protectionGuard.Query() == ProtectionGuardOutcome.NotProtected
            ? null
            : ShapeWriteRefusal.TargetProtected;

    /// <summary>
    /// Resolves the shapes collection of the worksheet carrying <c>tblGanttData</c>, after
    /// the protection guard has cleared the resolved target.
    /// </summary>
    /// <param name="refusal">The refusal to return when resolution fails.</param>
    /// <returns>The shapes collection, or <see langword="null"/> with <paramref name="refusal"/> set.</returns>
    internal Excel.Shapes? ResolveShapes(out ShapeWriteRefusal? refusal)
    {
        refusal = null;

        Excel.Application? application = _application;
        if (application is null)
        {
            refusal = ShapeWriteRefusal.NoActiveWorkbook;
            return null;
        }

        Excel.Workbook? workbook = GetActiveWorkbook(application);
        if (workbook is null)
        {
            refusal = ShapeWriteRefusal.NoActiveWorkbook;
            return null;
        }

        Excel.Sheets? sheets = GetSheets(workbook);
        if (sheets is null)
        {
            refusal = ShapeWriteRefusal.WorksheetMissing;
            return null;
        }

        Excel._Worksheet? gantt = FindGanttWorksheet(sheets);
        if (gantt is null)
        {
            refusal = ShapeWriteRefusal.WorksheetMissing;
            return null;
        }

        // The authoritative check is against the resolved target, not the active sheet: the
        // Gantt worksheet is not necessarily the active one (R4.8 D3).
        ProtectionGuardOutcome target = _protectionGuard.QueryTarget(gantt);
        if (target != ProtectionGuardOutcome.NotProtected)
        {
            refusal = ShapeWriteRefusal.TargetProtected;
            return null;
        }

        Excel.Shapes? shapes = GetShapes(gantt);
        if (shapes is null)
        {
            refusal = ShapeWriteRefusal.WorksheetMissing;
            return null;
        }

        return shapes;
    }

    /// <summary>
    /// Finds a shape by exact name and proves it staging-owned by recomputing the tag from
    /// the live name (R6.2 D4).
    /// </summary>
    /// <param name="shapes">The resolved shapes collection.</param>
    /// <param name="name">The exact staged shape name.</param>
    /// <param name="refusal">
    /// <see cref="ShapeWriteRefusal.NotFound"/> when the shape is absent or does not carry the
    /// staging tag for its live name; otherwise <see langword="null"/>.
    /// </param>
    /// <returns>The owned shape, or <see langword="null"/> with <paramref name="refusal"/> set.</returns>
    /// <remarks>
    /// Ownership is never the name: a user's shape that happens to share the name is refused
    /// and never touched, exactly as every write operation proves ownership.
    /// </remarks>
    internal Excel.Shape? FindOwnedStagingShape(
        Excel.Shapes shapes,
        string name,
        out ShapeWriteRefusal? refusal)
    {
        ArgumentNullException.ThrowIfNull(shapes);
        ArgumentException.ThrowIfNullOrWhiteSpace(name);

        refusal = null;

        Excel.Shape? shape = FindShapeByName(shapes, name);
        if (shape is null)
        {
            refusal = ShapeWriteRefusal.NotFound;
            return null;
        }

        if (!string.Equals(
                ReadAlternativeText(shape),
                ShapeStagingTag.ForStagedName(name),
                StringComparison.Ordinal))
        {
            refusal = ShapeWriteRefusal.NotFound;
            return null;
        }

        return shape;
    }

    // ---- Shared host-read seams (internal virtual, per the ExcelShapeWriter pattern) ----

    /// <summary>Reads the active workbook. Test seam over <c>Application.ActiveWorkbook</c>.</summary>
    /// <param name="application">The application object.</param>
    /// <returns>The workbook, or <see langword="null"/>.</returns>
    internal virtual Excel.Workbook? GetActiveWorkbook(Excel.Application application) =>
        application.ActiveWorkbook;

    /// <summary>Reads a workbook's sheets. Test seam over <c>Workbook.Sheets</c>.</summary>
    /// <param name="workbook">The workbook.</param>
    /// <returns>The sheet collection, or <see langword="null"/>.</returns>
    internal virtual Excel.Sheets? GetSheets(Excel.Workbook workbook) => workbook.Sheets;

    /// <summary>Reads the sheet count. Test seam over <c>Sheets.Count</c>.</summary>
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
    /// Returns the worksheet carrying <c>tblGanttData</c>. Test seam over the
    /// <c>ListObjects</c> indexed property and the <c>Name</c> read (CS0855).
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

    /// <summary>Whether a sheet carries <c>tblGanttData</c>. Test seam over the table collection.</summary>
    /// <param name="sheet">The candidate worksheet.</param>
    /// <returns><see langword="true"/> when the sheet carries the Gantt table.</returns>
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

    /// <summary>Reads a worksheet's shapes collection. Test seam over <c>_Worksheet.Shapes</c>.</summary>
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

    /// <summary>Reads a shape's alternative text. Test seam over <c>Shape.AlternativeText</c>.</summary>
    /// <param name="shape">The shape.</param>
    /// <returns>The alternative text, or <see langword="null"/>.</returns>
    internal virtual string? ReadAlternativeText(Excel.Shape shape) => shape.AlternativeText;

    /// <summary>
    /// Finds a shape by its exact name, preferring the host's name indexer so an absent
    /// name is a <see langword="null"/> rather than a scan of every shape.
    /// </summary>
    /// <param name="shapes">The shapes collection.</param>
    /// <param name="name">The exact shape name.</param>
    /// <returns>The shape, or <see langword="null"/> when absent.</returns>
    /// <remarks>
    /// <para>
    /// <b>Measured cost of the old scan.</b> This used to read every shape's <c>Name</c>
    /// until it found a match, on the argument that staging runs once over a bounded set.
    /// The grouping preflight calls it once per staged primitive, so it was O(N) per call
    /// and O(N²) per copy. Measured 2026-11-19 on a 20-activity workbook (Excel 16.0
    /// x64): 340 staged shapes against a 438-shape collection, 340 lookups,
    /// <b>29.1 s</b> of the operation's 39.4 s — about 149,000 COM name reads.
    /// </para>
    /// <para>
    /// The host name indexer is the fast path the live writer already uses (ADR-0054 O4),
    /// and a missing name RAISES rather than returning null. Catching the miss keeps the
    /// "null means absent" contract for every caller without the scan. The measured miss
    /// types are the same three the writer catches (<c>ArgumentException</c> on Excel
    /// 16.0 x64, plus <c>COMException</c> and <c>InvalidOperationException</c> from other
    /// call sites). The scan remains only as the degraded path for an indexer that
    /// resolves nothing without raising (a test double or a different host).
    /// </para>
    /// </remarks>
    internal virtual Excel.Shape? FindShapeByName(Excel.Shapes shapes, string name)
    {
        ArgumentNullException.ThrowIfNull(shapes);
        ArgumentException.ThrowIfNullOrWhiteSpace(name);

        try
        {
            if (GetShapeByExactName(shapes, name) is { } resolved)
            {
                return resolved;
            }
        }
        catch (Exception exception) when (exception is ArgumentException or COMException or InvalidOperationException)
        {
            return null;
        }

        int count = GetShapeCount(shapes);
        for (int index = 1; index <= count; index++)
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
    /// Resolves a shape through the host's name indexer. Test seam over the parameterised
    /// COM property, which cannot be expressed in a Moq expression (CS0855); overrides
    /// model the hit, the raising miss, and the degraded miss alike.
    /// </summary>
    /// <param name="shapes">The worksheet's shapes collection.</param>
    /// <param name="name">The exact shape name.</param>
    /// <returns>The shape when the host found it.</returns>
    /// <exception cref="ArgumentException">Raised by Excel 16.0 x64 when no shape carries the name.</exception>
    /// <exception cref="COMException">The other encoding the interop boundary produces for the miss.</exception>
    internal virtual Excel.Shape? GetShapeByExactName(Excel.Shapes shapes, string name)
    {
        ArgumentNullException.ThrowIfNull(shapes);
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        return shapes.Item(name);
    }
}
