using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live <see cref="IShapeCopyPort"/> over the Excel application object supplied by the
/// host: copies the verified staged group to the clipboard (R6.6).
/// </summary>
/// <remarks>
/// <para>
/// <b>Target resolution and ownership are inherited.</b> <see cref="StagingShapeAdapter"/>
/// resolves the worksheet carrying <c>tblGanttData</c>, queries the protection guard against
/// that resolved target, and proves staging ownership by recomputing the tag from the live
/// name. This adapter adds only the copy itself.
/// </para>
/// <para>
/// <b>COM ownership.</b> The <c>Application</c>, <c>Workbook</c>, <c>Worksheet</c>,
/// <c>Shapes</c> and <c>Shape</c> objects reached here are Excel-owned shared roots. This
/// adapter takes no ownership of them, never calls <c>FinalReleaseComObject</c>, and every
/// proxy is held in a local and used without chained member expressions.
/// </para>
/// <para>
/// <b>Test seams.</b> <see cref="CopyShape"/> is an <c>internal virtual</c> seam, so a
/// contract test can observe that the copy was attempted on the proven shape - and force a
/// host refusal - without owning a clipboard. The real call is exercised by the tagged live
/// integration test.
/// </para>
/// <para>
/// <b>Nothing is exposed through a command yet.</b> R6.6 fills the copy seam; R6.8 owns the
/// click.
/// </para>
/// </remarks>
/// <param name="application">
/// The Excel application object (for example <c>ExcelDnaUtil.Application</c>), or
/// <see langword="null"/> when the host supplied none. A foreign object fails the interface
/// cast and degrades to the no-active-workbook refusal with no mutation.
/// </param>
/// <param name="protectionGuard">The shared read-only workbook-protection guard.</param>
public class ExcelShapeCopier(
    object? application,
    IWorksheetProtectionGuard? protectionGuard = null)
    : StagingShapeAdapter(application, protectionGuard), IShapeCopyPort
{
    /// <inheritdoc />
    public ShapeWriteOutcome CopyStaged(string groupName)
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

        // Ownership is proved by recomputation from the LIVE name, exactly as every write
        // operation proves it: a user's shape that happens to share the name is refused.
        Excel.Shape? group = FindOwnedStagingShape(shapes, groupName, out ShapeWriteRefusal? ownershipRefusal);
        if (group is null)
        {
            return ShapeWriteOutcome.Refused(ownershipRefusal ?? ShapeWriteRefusal.NotFound);
        }

        try
        {
            CopyShape(group);
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

    /// <summary>Copies one shape to the clipboard. Test seam over <c>Shape.Copy</c>.</summary>
    /// <param name="shape">The proven, staging-owned group shape.</param>
    internal virtual void CopyShape(Excel.Shape shape)
    {
        ArgumentNullException.ThrowIfNull(shape);
        shape.Copy();
    }
}
