using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;

namespace GanttCreator.AddIn;

/// <summary>
/// Excel-DNA RibbonX entry point. Returns a minimal, valid RibbonX document
/// that declares a single Gantt Creator tab.
/// </summary>
/// <remarks>
/// Excel-DNA auto-discovers and registers any non-abstract
/// <see cref="ExcelRibbon"/> descendant (see
/// <c>AssemblyLoader.IsRibbonType</c>), so no .dna file change is required.
/// <see cref="GetCustomUI"/> does not call the base implementation: the base
/// reads <c>DnaLibrary.CustomUIs</c>, which is empty for this add-in and null
/// when not hosted in Excel.
/// </remarks>
[ComVisible(true)]
public class GanttRibbon : ExcelRibbon
{
    /// <summary>
    /// Excel's onLoad callback. No-op: this add-in has no dynamic load-time
    /// ribbon state.
    /// </summary>
    /// <param name="ribbon">The RibbonUI handle supplied by Excel.</param>
#pragma warning disable IDE0060 // param required by the RibbonX onLoad contract
    public void OnLoad(IRibbonUI ribbon)
    {
        // Intentionally empty: no load-time ribbon state.
    }
#pragma warning restore IDE0060

    /// <summary>
    /// Returns the RibbonX document for the Excel workbook RibbonID, else null.
    /// The document is stored in the RibbonResources.resx resource (a
    /// ResXFileRef to <c>RibbonResources\Ribbon.xml</c>) and read through the
    /// committed <c>RibbonResources.Designer.cs</c> accessor, so the XML never
    /// appears as a C# string literal.
    /// </summary>
    /// <param name="RibbonID">The Ribbon identifier supplied by Excel.</param>
    /// <returns>
    /// The RibbonX document, or null when the RibbonID is not the workbook.
    /// </returns>
    public override string GetCustomUI(string RibbonID) => RibbonID != "Microsoft.Excel.Workbook" ? null! : RibbonResources.Ribbon;

    /// <summary>
    /// Called when the user clicks the Diagnostics button. Delegates to the
    /// project-wide <see cref="DiagnosticsService"/> singleton.
    /// </summary>
    /// <param name="control">The ribbon control that raised the event.</param>
#pragma warning disable IDE0060 // control is required by the RibbonX onAction contract.
    public void OnDiagnosticsClick(IRibbonControl control)
#pragma warning restore IDE0060 // control is required by the RibbonX onAction contract.
    {
        // CA1031: The ribbon callback must never propagate an exception into
        // Excel — an unhandled onAction exception surfaces as a host error.
        // Diagnostics failures degrade to a missing dialog rather than a crash.
#pragma warning disable CA1031
        try
        {
            DiagnosticsService.Instance.ShowDiagnostics();
        }
        catch
#pragma warning restore CA1031
        {
            // Intentionally empty: no propagation.
        }
    }
}
