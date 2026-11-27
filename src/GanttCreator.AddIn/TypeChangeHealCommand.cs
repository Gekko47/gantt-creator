using GanttCreator.Core;
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// Heals a row's type-dependent cells when its Type cell changes (R5.23), on the change
/// path and independently of whether the chart is rebuilt.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why it is not part of the pass.</b> The pass's step 5b normalises every row, but it
/// only runs when the render is allowed. With the automatic refresh off or the chart
/// unlocked a Type edit was never healed, and the ribbon — which now always reflects the
/// sheet — showed exactly that unhealed row.
/// </para>
/// <para>
/// <b>Old type from the cache.</b> The sheet no longer holds the previous Type. The rows
/// the last pass or projection read are the pre-edit baseline (a selection change does not
/// move that cache), and they are matched by engine Id so a reorder cannot borrow the
/// wrong row's type.
/// </para>
/// <para>
/// <b>The caller republishes.</b> This command only writes. The scheduler runs the render
/// or the projection straight after it, and that action owns the ribbon republish, so a
/// Type burst still costs one capture.
/// </para>
/// <para>
/// <b>The release window is opened even though only cells are written.</b>
/// <see cref="ExcelTypeChangeNormaliser"/> refuses a protected target before it writes, so
/// without the window an edit that reached a protected sheet would never heal. The window
/// is the same nested scope the refresh uses, so it releases once and re-protects once.
/// </para>
/// </remarks>
internal static class TypeChangeHealCommand
{
    internal static void RunForExcel()
    {
        // CA1031: this runs inside the scheduler's queued macro. A failure must degrade to
        // the next pass, which applies the same Core rules, not escape into Excel.
#pragma warning disable CA1031
        try
        {
            object? application = ExcelDna.Integration.ExcelDnaUtil.Application;
            if (application is null)
            {
                return;
            }

            Dictionary<GanttRowId, GanttEntityType> oldTypes = [];
            foreach (GanttRowDto row in RibbonStateService.Instance.GetLastTableRows())
            {
                if (GanttRowId.TryParse(row.Id, out GanttRowId? id)
                    && id is not null
                    && EntityTypeCatalog.TryParse(row.TypeText, out GanttEntityType type))
                {
                    oldTypes[id] = type;
                }
            }

            ConfigReadOutcome config = new ExcelConfigCatalogueReader(
                application, CommandBoundary.Instance.WriteTechnicalRecord).Read();

            GanttTypeChangeNormaliseOutcome outcome;
            using IDisposable window = RefreshSheetCommand.OpenProtectionWindow();
            using ExcelApplicationStateScope scope = new(
                application, CommandBoundary.Instance.WriteTechnicalRecord);
            scope.SuppressScreenUpdating();
            scope.SuppressEvents();
            scope.SuppressAlerts();
            scope.SuppressStatusBar();

            outcome = new ExcelTypeChangeNormaliser(
                application,
                trace: CommandBoundary.Instance.WriteTechnicalRecord).Normalise(
                config.Succeeded ? config.Styles : null,
                id => oldTypes.TryGetValue(id, out GanttEntityType type) ? type : null);

            CommandBoundary.Instance.WriteTechnicalRecord(
                string.Create(
                    System.Globalization.CultureInfo.InvariantCulture,
                    $"gantt.types heal cached={oldTypes.Count} cells={outcome.CellsWritten}"));
        }
        catch (Exception ex)
        {
            // The reason, not just the fact: a protected target, a failed config read and a
            // defect in the transition are indistinguishable from "refused" alone.
            CommandBoundary.Instance.WriteTechnicalRecord(
                "gantt.types heal refused type=" + ex.GetType().Name + " message=" + ex.Message);
        }
#pragma warning restore CA1031
    }
}
