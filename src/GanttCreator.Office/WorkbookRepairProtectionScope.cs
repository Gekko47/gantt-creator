// Namespace alias rather than a wholesale using: Microsoft.Office.Interop.Excel declares its own
// Action type, which would collide with System.Action under a plain using.
using GanttCreator.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Releases the protection the explicit Fix command has to get past, for the whole repair
/// sequence, and puts it back on every exit path.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why the scope is workbook-wide rather than per adapter.</b> Every stage of Fix refuses a
/// protected target instead of writing through it: <see cref="ExcelWorkbookRecovery.Recover"/>
/// and <see cref="ExcelColumnVisibilityPort"/> both return a typed refusal, and the catalogue
/// writer checks the configuration sheet before it writes. Releasing protection inside one
/// adapter therefore only moves the wall: the sequence stops at the next stage that meets it,
/// and the user has to press Fix once per stage (owner report 2026-11-27: "unlock the cells,
/// unhide the columns, then click Fix two or three times").
/// </para>
/// <para>
/// <b>Which sheets.</b> The visible worksheet carrying the Gantt table (whose shapes, columns
/// and rows the repair touches) and <c>_GanttCreatorConfig</c> (which the catalogue writer
/// rewrites). A workbook whose configuration sheet is protected is refused by the writer even
/// after the visible sheet has been released, so both must be inside one window.
/// </para>
/// <para>
/// <b>Restoration.</b> The window re-applies protection only where it released protection it did
/// not arm, which is the ADR-0011 amendment's rule for Fix. A protection the host refuses to
/// release stays on, the adapters' own checks refuse the write, and the user sees that refusal
/// rather than a silent half-repair.
/// </para>
/// </remarks>
public sealed class WorkbookRepairProtectionScope : IDisposable
{
    private readonly List<IDisposable> _windows;

    private WorkbookRepairProtectionScope(List<IDisposable> windows)
    {
        _windows = windows;
    }

    /// <summary>
    /// Opens the repair window over the Gantt worksheet and the configuration worksheet.
    /// </summary>
    /// <param name="application">The Excel application object.</param>
    /// <param name="trace">Optional technical-record sink for the window's own records.</param>
    /// <param name="preferences">The user's protection toggles, or null for the shared source.</param>
    /// <returns>
    /// The scope that restores protection. Never null: a workbook with nothing protected costs
    /// one bool read per sheet and releases nothing.
    /// </returns>
    public static WorkbookRepairProtectionScope Open(
        object? application,
        Action<string>? trace = null,
        IProtectionPreferenceSource? preferences = null)
    {
        // One arming state across both windows: the repair is not the add-in's own arming
        // (a reopened workbook's protection is a previous session's), so both windows release
        // foreign protection and both restore what they released.
        ProtectionArmingState arming = new();
        List<IDisposable> windows = [];

        WorksheetProtectionWindow gantt = new(
            new ExcelWorksheetProtectionPort(application, trace, preferences),
            arming,
            trace);
        windows.Add(gantt.Open(repairCells: true, releaseForeignProtection: true));

        if (FindConfigWorksheet(application) is { } config)
        {
            WorksheetProtectionWindow configuration = new(
                new ExcelWorksheetProtectionPort(application, trace, preferences, config),
                arming,
                trace);
            windows.Add(configuration.Open(repairCells: true, releaseForeignProtection: true));
        }

        return new WorkbookRepairProtectionScope(windows);
    }

    /// <inheritdoc />
    public void Dispose()
    {
        // Reverse order, so the configuration sheet is re-protected before the visible sheet:
        // the visible sheet is the one the user is looking at, and a throw while protecting the
        // configuration sheet must not leave the chart sheet open.
        for (int index = _windows.Count - 1; index >= 0; index--)
        {
            _windows[index].Dispose();
        }

        _windows.Clear();
    }

    /// <summary>
    /// Finds <c>_GanttCreatorConfig</c> by contract name in the active workbook.
    /// </summary>
    /// <param name="application">The Excel application object.</param>
    /// <returns>The configuration worksheet, or <see langword="null"/>.</returns>
    private static Excel._Worksheet? FindConfigWorksheet(object? application)
    {
        if (application is not Excel.Application host)
        {
            return null;
        }

        Excel.Workbook? workbook = host.ActiveWorkbook;
        if (workbook is null)
        {
            return null;
        }

        Excel.Sheets sheets = workbook.Sheets;
        int count = sheets.Count;
        for (int index = 1; index <= count; index++)
        {
            if (sheets[index] is Excel._Worksheet worksheet
                && GanttWorkbookContract.IsConfigSheetName(worksheet.Name))
            {
                return worksheet;
            }
        }

        return null;
    }
}
