using System.Runtime.InteropServices;
// Namespace alias rather than a wholesale using: Microsoft.Office.Interop.Excel declares its own
// Action type, which would collide with System.Action under a plain using.
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live <see cref="ISelectedShapePort"/> over <c>Application.Selection</c> (ADR-0055 clause 6).
/// </summary>
/// <remarks>
/// <para>
/// <b>Never throws.</b> The port is read by a Ribbon <c>getEnabled</c> callback and by the Unlock
/// command, and a propagating COM failure in either place surfaces as a host error dialog. Every
/// refusal is typed instead.
/// </para>
/// <para>
/// One shape only. A cell range, a chart, or a multi-shape selection all report
/// <see cref="SelectedShapeRefusalReason.SelectionIsNotASingleShape"/>: unlocking "whatever is
/// selected" would make the command's effect depend on a selection the user can change after reading
/// the button's tooltip.
/// </para>
/// </remarks>
public class ExcelSelectedShapeReader(object? application) : ISelectedShapePort
{
    private readonly Excel.Application? _application = application as Excel.Application;

    /// <inheritdoc />
    public SelectedShapeOutcome Read()
    {
        Excel.Application? application = _application;
        if (application is null)
        {
            return SelectedShapeOutcome.Refused(SelectedShapeRefusalReason.NoApplication);
        }

        // CA1031: a Ribbon getEnabled callback and a command entry point both call this, and the
        // host can refuse a property read on a chart sheet, a grouped shape, or a protected sheet. A
        // refusal is the useful answer; an exception would become a host error dialog.
#pragma warning disable CA1031
        try
        {
            if (GetSelection(application) is not Excel.Shape shape)
            {
                return SelectedShapeOutcome.Refused(SelectedShapeRefusalReason.SelectionIsNotASingleShape);
            }

            string name = shape.Name;
            if (string.IsNullOrWhiteSpace(name))
            {
                return SelectedShapeOutcome.Refused(SelectedShapeRefusalReason.SelectionIsNotASingleShape);
            }

            // The tag recomputed from the name authorises the unlock, exactly as it authorises a
            // delete or an update: a user's shape that happens to share a name is refused.
            string? tag = shape.AlternativeText;
            return string.Equals(tag, ShapeOwnershipTag.ForPrimitiveId(name), StringComparison.Ordinal)
                ? SelectedShapeOutcome.Ok(name)
                : SelectedShapeOutcome.Refused(SelectedShapeRefusalReason.ShapeNotOwned);
        }
        catch (COMException)
        {
            return SelectedShapeOutcome.Refused(SelectedShapeRefusalReason.HostRejected);
        }
        catch (InvalidCastException)
        {
            return SelectedShapeOutcome.Refused(SelectedShapeRefusalReason.SelectionIsNotASingleShape);
        }
#pragma warning restore CA1031
    }

    /// <summary>Reads <c>Application.Selection</c>. Test seam over the COM property.</summary>
    /// <param name="application">The Excel application.</param>
    /// <returns>The current selection, which the host types as <c>Object</c>.</returns>
    internal virtual object? GetSelection(Excel.Application application) => application.Selection;
}
