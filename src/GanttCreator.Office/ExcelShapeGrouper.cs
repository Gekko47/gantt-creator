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
/// <b>Target resolution is inherited from <see cref="StagingShapeAdapter"/>.</b> The Gantt
/// worksheet is the one carrying <c>tblGanttData</c> (the user may rename the sheet, the
/// table name is the schema's identity), and the protection guard is queried against that
/// resolved target, never the active sheet (ADR-0008 D4, R4.8 D3). The R6.6 copy adapter
/// shares the same resolution and ownership proof rather than restating them.
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
    IWorksheetProtectionGuard? protectionGuard = null)
    : StagingShapeAdapter(application, protectionGuard), IShapeGroupingPort
{

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
        if (CheckEntryGuard() is { } entryRefusal)
        {
            return Refused(entryRefusal);
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
        if (CheckEntryGuard() is { } entryRefusal)
        {
            return ShapeWriteOutcome.Refused(entryRefusal);
        }

        Excel.Shapes? shapes = ResolveShapes(out ShapeWriteRefusal? resolutionRefusal);
        if (shapes is null)
        {
            return ShapeWriteOutcome.Refused(resolutionRefusal ?? ShapeWriteRefusal.WorksheetMissing);
        }

        Excel.Shape? group = FindOwnedStagingShape(shapes, groupName, out ShapeWriteRefusal? ownershipRefusal);
        if (group is null)
        {
            return ShapeWriteOutcome.Refused(ownershipRefusal ?? ShapeWriteRefusal.NotFound);
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

}
