using System.Runtime.InteropServices;
using GanttCreator.Core;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live <see cref="IShapeGroupingPort"/> over the Excel application object supplied by the
/// host: groups the staged shape set into one shape group, stamps the staging identity on
/// it, and reports the child order and envelope the host holds (R6.4).
/// </summary>
/// <param name="application">
/// The Excel application object (for example <c>ExcelDnaUtil.Application</c>), or
/// <see langword="null"/> when the host supplied none (unit tests, non-Excel host). A
/// foreign object fails the interface cast and degrades to the no-active-workbook refusal
/// with no mutation.
/// </param>
/// <param name="protectionGuard">The shared read-only workbook-protection guard.</param>
/// <remarks>
/// <para>
/// <b>Target resolution mirrors <see cref="ExcelShapeWriter"/>.</b> The Gantt worksheet is
/// the one carrying <c>tblGanttData</c> (the user may rename the sheet, the table name is
/// the schema's identity), and the protection guard is queried against that resolved
/// target, never the active sheet (ADR-0008 D4, R4.8 D3).
/// </para>
/// <para>
/// <b>COM ownership.</b> The <c>Application</c>, <c>Workbook</c>, <c>Worksheet</c>,
/// <c>Shapes</c>, <c>ShapeRange</c>, <c>Shape</c> and <c>GroupShapes</c> objects reached
/// here are Excel-owned shared roots. This adapter takes no ownership of them, never calls
/// <c>FinalReleaseComObject</c>, and force-releases nothing (the ownership policy of
/// <see cref="ExcelApplicationAdapter"/>). Every proxy is held in a local and used without
/// chained member expressions.
/// </para>
/// <para>
/// <b>Test seams.</b> The <c>internal virtual</c> accessors isolate the Excel COM
/// parameterised properties (indexers), the group call, and the read-backs. Expression
/// trees cannot contain indexed properties (CS0855), so contract tests substitute these
/// seams; the real indexer behaviour is exercised by the tagged live-Office integration
/// test.
/// </para>
/// <para>
/// <b>A single-shape range is not a group.</b> Excel's <c>ShapeRange.Group</c> returns the
/// shape itself when the range holds one shape. That is reported as
/// <see cref="ShapeWriteRefusal.HostRejected"/> rather than passed off as a group: the
/// grouping stage's membership check would otherwise be comparing against an empty
/// <c>GroupItems</c> collection with no explanation.
/// </para>
/// </remarks>
public class ExcelShapeGrouper(
    object? application,
    IWorksheetProtectionGuard? protectionGuard = null) : IShapeGroupingPort
{
    private readonly Excel.Application? _application = application as Excel.Application;
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    /// <inheritdoc />
    public ShapeGroupReadback GroupStaged(IReadOnlyList<string> shapeNames)
    {
        if (shapeNames is null)
        {
            return Refused(ShapeWriteRefusal.NullRequest);
        }

        if (shapeNames.Count == 0)
        {
            return Refused(ShapeWriteRefusal.BlankIdentifier);
        }

        foreach (string name in shapeNames)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return Refused(ShapeWriteRefusal.BlankIdentifier);
            }
        }

        // ADR-0008 D4's first host check, before any workbook access; the authoritative
        // per-target check runs again in ResolveShapes against the resolved worksheet.
        if (_protectionGuard.Query() != ProtectionGuardOutcome.NotProtected)
        {
            return Refused(ShapeWriteRefusal.TargetProtected);
        }

        Excel.Shapes? shapes = ResolveShapes(out ShapeWriteRefusal? resolutionRefusal);
        if (shapes is null)
        {
            return Refused(resolutionRefusal ?? ShapeWriteRefusal.WorksheetMissing);
        }

        // Preflight, then commit: every staged name is resolved before the first mutation,
        // because Shapes.Range raises for a missing name rather than returning an empty
        // range. One absent shape must refuse the whole grouping with nothing changed.
        foreach (string name in shapeNames)
        {
            if (FindShapeByName(shapes, name) is null)
            {
                return Refused(ShapeWriteRefusal.NotFound);
            }
        }

        Excel.Shape? group = null;
        try
        {
            Excel.ShapeRange range = GetShapeRange(shapes, shapeNames);
            group = GroupRange(range);

            // Excel returns the shape itself for a one-shape range; that is not a group.
            if (group is null || ReadShapeType(group) != MsoShapeType.msoGroup)
            {
                DiscardGroup(group);
                return Refused(ShapeWriteRefusal.HostRejected);
            }

            ApplyGroupStamp(group);

            string[] children = ReadChildNames(group);
            RectD? bounds = ReadBounds(group);
            if (bounds is null)
            {
                // A host that reports non-finite geometry has failed the operation; restore
                // the children rather than report a group the stage cannot verify.
                DiscardGroup(group);
                return Refused(ShapeWriteRefusal.HostRejected);
            }

            return new ShapeGroupReadback(group.Name, children, bounds, null);
        }
        catch (COMException)
        {
            DiscardGroup(group);
            return Refused(ShapeWriteRefusal.HostRejected);
        }
        catch (InvalidOperationException)
        {
            DiscardGroup(group);
            return Refused(ShapeWriteRefusal.HostRejected);
        }
        catch (ArgumentException)
        {
            // Shapes.Range raises ArgumentException for an index the host rejects even
            // though the preflight proved the names resolve (the interop boundary reports
            // the miss this way); it is a host refusal here, not a caller error.
            DiscardGroup(group);
            return Refused(ShapeWriteRefusal.HostRejected);
        }
    }

    /// <inheritdoc />
    public ShapeWriteOutcome UngroupStaged(string groupName)
    {
        if (string.IsNullOrWhiteSpace(groupName))
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.BlankIdentifier);
        }

        // ADR-0008 D4's first host check, before any workbook access; the authoritative
        // per-target check runs again in ResolveShapes against the resolved worksheet.
        if (_protectionGuard.Query() != ProtectionGuardOutcome.NotProtected)
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.TargetProtected);
        }

        Excel.Shapes? shapes = ResolveShapes(out ShapeWriteRefusal? resolutionRefusal);
        if (shapes is null)
        {
            return ShapeWriteOutcome.Refused(resolutionRefusal ?? ShapeWriteRefusal.WorksheetMissing);
        }

        Excel.Shape? group = FindShapeByName(shapes, groupName);
        if (group is null)
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.NotFound);
        }

        // Ownership is proved by recomputation from the LIVE name, exactly as every write
        // operation proves it: a user's shape that happens to share the name is refused.
        if (!string.Equals(
                ReadAlternativeText(group),
                ShapeStagingTag.ForStagedName(groupName),
                StringComparison.Ordinal))
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.NotFound);
        }

        try
        {
            UngroupShape(group);
            return ShapeWriteOutcome.Ok();
        }
        catch (COMException)
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.HostRejected);
        }
        catch (InvalidOperationException)
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.HostRejected);
        }
        catch (ArgumentException)
        {
            return ShapeWriteOutcome.Refused(ShapeWriteRefusal.HostRejected);
        }
    }

    /// <summary>
    /// Resolves the shapes collection of the worksheet carrying <c>tblGanttData</c>, after
    /// the protection guard has cleared the resolved target.
    /// </summary>
    /// <param name="refusal">The refusal to return when resolution fails.</param>
    /// <returns>The shapes collection, or <see langword="null"/> with <paramref name="refusal"/> set.</returns>
    private Excel.Shapes? ResolveShapes(out ShapeWriteRefusal? refusal)
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
    /// Deletes nothing: restores a group's children to the sheet so a failed grouping
    /// leaves the staged set exactly as it was.
    /// </summary>
    /// <param name="group">The group to ungroup, or <see langword="null"/> when none was created.</param>
    /// <remarks>
    /// Best effort, on the same reasoning as the writer's unowned-shape discard: the caller
    /// is already being told about the original refusal, and a cleanup failure that
    /// replaced it with an exception would hide the cause. A child left inside an
    /// unstamped group is still staging-tagged, so R6.5's sweep can reach it.
    /// </remarks>
    private void DiscardGroup(Excel.Shape? group)
    {
        if (group is null)
        {
            return;
        }

        // CA1031: a cleanup failure must never replace the refusal being reported.
#pragma warning disable CA1031
        try
        {
            if (ReadShapeType(group) == MsoShapeType.msoGroup)
            {
                UngroupShape(group);
            }
        }
        catch (Exception)
        {
            // Intentionally empty: the original refusal is the report.
        }
#pragma warning restore CA1031
    }

    private static ShapeGroupReadback Refused(ShapeWriteRefusal refusal) =>
        new(null, [], null, refusal);

    /// <summary>Reads the child names in the host's <c>GroupItems</c> order.</summary>
    /// <param name="group">The group shape.</param>
    /// <returns>The child names, 1..Count.</returns>
    private string[] ReadChildNames(Excel.Shape group)
    {
        Excel.GroupShapes? items = GetGroupItems(group);
        if (items is null)
        {
            return [];
        }

        int count = GetGroupItemCount(items);
        string[] names = new string[count];
        for (int index = 1; index <= count; index++)
        {
            names[index - 1] = GetGroupItemName(items, index);
        }

        return names;
    }

    /// <summary>Reads the group's envelope, or <see langword="null"/> for non-finite members.</summary>
    /// <param name="group">The group shape.</param>
    /// <returns>The envelope in scene points, or <see langword="null"/>.</returns>
    private RectD? ReadBounds(Excel.Shape group)
    {
        double left = ReadLeft(group);
        double top = ReadTop(group);
        double width = ReadWidth(group);
        double height = ReadHeight(group);

        return double.IsFinite(left) && double.IsFinite(top)
            && double.IsFinite(width) && double.IsFinite(height)
            ? new RectD(left, top, width, height)
            : null;
    }

    // ---- Test seams (internal virtual, per the ExcelShapeWriter pattern) ----

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

    /// <summary>
    /// Builds a shape range from exact names. Test seam over the parameterised
    /// <c>Shapes.Range</c> property, which cannot appear in a Moq expression (CS0855).
    /// </summary>
    /// <param name="shapes">The worksheet's shapes collection.</param>
    /// <param name="shapeNames">The exact names to select, in scene order.</param>
    /// <returns>The selected range.</returns>
    internal virtual Excel.ShapeRange GetShapeRange(Excel.Shapes shapes, IReadOnlyList<string> shapeNames)
    {
        ArgumentNullException.ThrowIfNull(shapes);
        ArgumentNullException.ThrowIfNull(shapeNames);

        object[] index = [.. shapeNames];
        return shapes.Range[index];
    }

    /// <summary>Groups a shape range. Test seam over <c>ShapeRange.Group</c>.</summary>
    /// <param name="range">The selected range.</param>
    /// <returns>The created shape, or <see langword="null"/> when the host refused.</returns>
    internal virtual Excel.Shape? GroupRange(Excel.ShapeRange range)
    {
        ArgumentNullException.ThrowIfNull(range);
        return range.Group();
    }

    /// <summary>Reads a shape's type. Test seam over <c>Shape.Type</c>.</summary>
    /// <param name="shape">The shape.</param>
    /// <returns>The host shape type.</returns>
    internal virtual MsoShapeType ReadShapeType(Excel.Shape shape) => shape.Type;

    /// <summary>Reads a group's child collection. Test seam over <c>Shape.GroupItems</c>.</summary>
    /// <param name="group">The group shape.</param>
    /// <returns>The child collection, or <see langword="null"/>.</returns>
    internal virtual Excel.GroupShapes? GetGroupItems(Excel.Shape group) => group.GroupItems;

    /// <summary>Reads the child count. Test seam over <c>GroupShapes.Count</c>.</summary>
    /// <param name="items">The child collection.</param>
    /// <returns>The number of children.</returns>
    internal virtual int GetGroupItemCount(Excel.GroupShapes items) => items.Count;

    /// <summary>Reads one child's name. Test seam over <c>GroupShapes.Item</c>.</summary>
    /// <param name="items">The child collection.</param>
    /// <param name="index">The 1-based child index.</param>
    /// <returns>The child shape name.</returns>
    internal virtual string GetGroupItemName(Excel.GroupShapes items, int index)
    {
        ArgumentNullException.ThrowIfNull(items);
        Excel.Shape child = items.Item(index);
        return child.Name;
    }

    /// <summary>Reads a shape's left edge. Test seam over <c>Shape.Left</c>.</summary>
    /// <param name="shape">The shape.</param>
    /// <returns>The left edge in points.</returns>
    internal virtual double ReadLeft(Excel.Shape shape) => shape.Left;

    /// <summary>Reads a shape's top edge. Test seam over <c>Shape.Top</c>.</summary>
    /// <param name="shape">The shape.</param>
    /// <returns>The top edge in points.</returns>
    internal virtual double ReadTop(Excel.Shape shape) => shape.Top;

    /// <summary>Reads a shape's width. Test seam over <c>Shape.Width</c>.</summary>
    /// <param name="shape">The shape.</param>
    /// <returns>The width in points.</returns>
    internal virtual double ReadWidth(Excel.Shape shape) => shape.Width;

    /// <summary>Reads a shape's height. Test seam over <c>Shape.Height</c>.</summary>
    /// <param name="shape">The shape.</param>
    /// <returns>The height in points.</returns>
    internal virtual double ReadHeight(Excel.Shape shape) => shape.Height;

    /// <summary>Reads a shape's alternative text. Test seam over <c>Shape.AlternativeText</c>.</summary>
    /// <param name="shape">The shape.</param>
    /// <returns>The alternative text, or <see langword="null"/>.</returns>
    internal virtual string? ReadAlternativeText(Excel.Shape shape) => shape.AlternativeText;

    /// <summary>
    /// Writes the staging identity and the temporary-shape placement class onto the group.
    /// </summary>
    /// <param name="group">The created group.</param>
    /// <remarks>
    /// <c>internal virtual</c> is a test seam, per the <see cref="ExcelShapeWriter"/> pattern:
    /// the stamp writes four COM members, and a test needs to assert it was reached - not
    /// merely that no error surfaced - to pin the rule that a group is never left unstamped.
    /// <c>xlMove</c> and the locked flag match every staged child (R6.3 D1, ADR-0055): the
    /// group is temporary and must not stretch with a user's row or column resize.
    /// </remarks>
    internal virtual void ApplyGroupStamp(Excel.Shape group)
    {
        ArgumentNullException.ThrowIfNull(group);

        group.Name = ShapeStagingTag.GroupName;
        group.AlternativeText = ShapeStagingTag.ForStagedName(ShapeStagingTag.GroupName);
        group.Placement = Excel.XlPlacement.xlMove;
        group.Locked = true;
    }

    /// <summary>Ungroups a group. Test seam over <c>Shape.Ungroup</c>.</summary>
    /// <param name="group">The group shape.</param>
    internal virtual void UngroupShape(Excel.Shape group)
    {
        ArgumentNullException.ThrowIfNull(group);
        _ = group.Ungroup();
    }

    /// <summary>
    /// Finds a shape by its exact name by scanning the collection, so an absent name is a
    /// <see langword="null"/> rather than a raised COM miss.
    /// </summary>
    /// <param name="shapes">The shapes collection.</param>
    /// <param name="name">The exact shape name.</param>
    /// <returns>The shape, or <see langword="null"/> when absent.</returns>
    /// <remarks>
    /// A scan rather than the host's raising name indexer, deliberately: grouping runs once
    /// per export over a bounded set, so the per-shape fast path the live writer needs
    /// (ADR-0054 O4) would buy nothing here and the scan keeps the "null means absent"
    /// contract without an exception per miss.
    /// </remarks>
    internal virtual Excel.Shape? FindShapeByName(Excel.Shapes shapes, string name)
    {
        ArgumentNullException.ThrowIfNull(shapes);
        ArgumentException.ThrowIfNullOrWhiteSpace(name);

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
}
