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
    // The RibbonX document. The Office 2009/07 namespace targets Excel 2010+,
    // the supported baseline.
#pragma warning disable IDE0055
    // IDE0055: dotnet format (the repo formatting gate) accepts this wrapping,
    // but EnforceCodeStyleInBuild flags the multiline binary expression. The
    // formatter and the in-build analyzer disagree here; suppressing for this
    // declaration only so the documented XML content is not churned further.
    private const string _ribbonXml =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
        + "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\" onLoad=\"OnLoad\">"
        + "<ribbon><tabs>"
        + "<tab id=\"tabGanttCreator\" label=\"Gantt Creator\">"
        + "<group id=\"grpPlaceholder\" label=\"Gantt Creator\"></group>"
        + "</tab></tabs></ribbon>"
        + "</customUI>";
#pragma warning restore IDE0055

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
    /// </summary>
    /// <param name="RibbonID">The Ribbon identifier supplied by Excel.</param>
    /// <returns>
    /// The RibbonX document, or null when the RibbonID is not the workbook.
    /// </returns>
    public override string GetCustomUI(string RibbonID)
    {
        return RibbonID != "Microsoft.Excel.Workbook" ? null! : _ribbonXml;
    }
}
