using System.Globalization;
using GanttCreator.Core;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live <see cref="IWorkbookInitialiser"/> over the Excel application object
/// supplied by the host at add-in load.
/// </summary>
/// <param name="application">
/// The Excel application object (for example <c>ExcelDnaUtil.Application</c>), or
/// <see langword="null"/> when the host supplied none (unit tests, non-Excel
/// host). A foreign object fails the interface cast and degrades to the
/// no-active-workbook refusal with no mutation.
/// </param>
/// <param name="catalogueWriter">
/// The catalogue writer that materialises the <c>_GanttCreatorConfig</c>
/// tables (R2.7); a production default is created when <see langword="null"/>.
/// Tests pass a stub to isolate the sheet contract from the catalogue
/// contract.
/// </param>
/// <param name="protectionGuard">The shared read-only workbook-protection guard.</param>
/// <param name="typeOptionsMaterialiser">The TypeOptions name and validation materialiser.</param>
/// <remarks>
/// <para>
/// COM ownership: the <c>Application</c>, <c>Workbook</c>, <c>Worksheet</c>,
/// <c>Range</c>, <c>ListObject</c>, and <c>Names</c> objects reached here are
/// Excel-owned shared roots. This adapter takes no ownership of them, never
/// calls <c>FinalReleaseComObject</c>, and force-releases nothing (the
/// ownership policy of <see cref="ExcelApplicationAdapter"/>). Every proxy is
/// held in a local and used without chained member expressions; sheets are
/// reached by index through <see cref="GetSheetAt"/> rather than a COM
/// enumerator.
/// </para>
/// <para>
/// Mutation order (work item R2.2 decision D4, extended by R2.7): all
/// read-only checks first, then rename, header row, table (with the R2.7
/// appearance settings), configuration sheet, catalogue tables, defined
/// name. A failure before the rename mutates nothing; a failure after it
/// leaves at most a renamed blank worksheet (cosmetic) and propagates to
/// the command boundary, which translates it to one safe message. The
/// catalogues live only on the configuration sheet, so the configuration
/// rollback removes them with the sheet; a typed catalogue refusal rolls
/// back every mutation above and returns the matching initialise refusal.
/// Every rollback is best-effort: a host refusal during a rollback is
/// caught there, so it neither masks the exception that triggered it --
/// the command boundary logs that original exception -- nor aborts the
/// remaining rollbacks that restore the user's sheet.
/// </para>
/// <para>
/// The three <c>internal virtual</c> accessors (<see cref="GetSheetAt"/>,
/// <see cref="GetTableAt"/>, <see cref="GetHeaderRange"/>) isolate the Excel
/// COM parameterised properties (indexers). Expression trees cannot contain
/// indexed properties (CS0855), so contract tests substitute these seams and
/// every other member through Moq; the real indexer behaviour is exercised by
/// the tagged live-Office integration test.
/// </para>
/// </remarks>
public class ExcelWorkbookInitialiser(
    object? application,
    IConfigCatalogueWriter? catalogueWriter = null,
    IWorksheetProtectionGuard? protectionGuard = null,
    ITypeOptionsMaterialiser? typeOptionsMaterialiser = null) : IWorkbookInitialiser
{
    private readonly Application? _application = application as Application;
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    private readonly IConfigCatalogueWriter _catalogueWriter =
        catalogueWriter ?? new ExcelConfigCatalogueWriter(application);

    private readonly ITypeOptionsMaterialiser _typeOptionsMaterialiser =
        typeOptionsMaterialiser ?? new ExcelTypeOptionsMaterialiser(application);

    /// <inheritdoc />
    public WorkbookInitialiseOutcome Initialise()
    {
        Application? application = _application;
        if (application is null)
        {
            return WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.NoActiveWorkbook);
        }

        // ADR-0008 D4: the shared protection guard is the first read-only check
        // for every mutating adapter. The target-specific checks below remain
        // authoritative for the sheet this command actually writes.
        ProtectionGuardOutcome protection = _protectionGuard.Query();
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return WorkbookInitialiseOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? InitialiseRefusalReason.NoActiveWorkbook
                    : InitialiseRefusalReason.TargetProtected);
        }

        // Resolve the actual mutation target before the authoritative protection
        // check. The active-sheet query remains the early host/workbook preflight.
        // One proxy per local: no chained `app.ActiveWorkbook.Worksheets[…]`
        // member expressions (docs/02-ARCHITECTURE.md COM ownership).
        Workbook? workbook = application.ActiveWorkbook;
        if (workbook is null)
        {
            return WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.NoActiveWorkbook);
        }

        Sheets sheets = workbook.Sheets;

        // Read-only check 1: the workbook must not already carry the
        // configuration sheet — creating a second helper sheet is a product
        // invariant violation, so this check runs before any mutation.
        if (NameTakenByOtherSheet(sheets, GanttWorkbookContract.ConfigSheetName, excludeSheetName: null))
        {
            return WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.ConfigSheetExists);
        }

        // Read-only check 2: scan every worksheet for a table named
        // GanttTableSchema.TableName before any sheets or headers are created.
        // This catches the table on any worksheet, not only the active one.
        if (AnyWorksheetContainsTable(sheets))
        {
            return WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.TableExists);
        }

        // Read-only check 3: select the target. The active worksheet is
        // adopted only when it is pristine; a chart sheet or any cell/non-cell
        // user state causes a fresh sheet to be created. A chart sheet is not
        // an Excel.Worksheet, so the cast selects the create path for it.
        var activeWorksheet = workbook.ActiveSheet as Worksheet;
        var adopt = activeWorksheet is not null
            && IsPristy(activeWorksheet, workbook, application);

        // Read-only check 3: resolve the label against every other sheet
        // (OrdinalIgnoreCase, matching Excel's own case-insensitive sheet-name
        // uniqueness) before any mutation. On the adopt path the target's own
        // current name is excluded — it is about to be renamed.
        var label = ResolveAvailableLabel(sheets, adopt ? activeWorksheet!.Name : null);

        // Read-only check 4: verify worksheet and workbook-structure
        // protection before any mutation. On the create path, both the
        // worksheet-level protection (which would block headerRange.Value2
        // writes) and the workbook-structure protection (which would block
        // sheet creation or rename) are checked here on the active worksheet
        // so the create path does not add a sheet before validation.
        if (!adopt && activeWorksheet is not null)
        {
            if (IsWorkbookStructureProtected(workbook))
            {
                return WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.TargetProtected);
            }
        }
        else if (adopt)
        {
            // The adopt path writes directly onto the active worksheet, so its
            // protection is authoritative for the later headerRange.Value2
            // write.
            if (IsWorksheetProtected(activeWorksheet!))
            {
                return WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.TargetProtected);
            }
        }

        Worksheet target = adopt ? activeWorksheet! : CreateTargetSheet(sheets, activeWorksheet);
        ProtectionGuardOutcome targetProtection = _protectionGuard.QueryTarget(target);
        if (targetProtection != ProtectionGuardOutcome.NotProtected)
        {
            if (!adopt)
            {
                RollBackCreatedSheet(target);
            }

            return WorkbookInitialiseOutcome.Refused(
                targetProtection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? InitialiseRefusalReason.NoActiveWorkbook
                    : InitialiseRefusalReason.TargetProtected);
        }

        try
        {
            if (!string.Equals(target.Name, label, StringComparison.OrdinalIgnoreCase))
            {
                target.Name = label;
            }
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // Expected user situation: the target (or the workbook
            // structure) is protected, so the rename — the first mutation —
            // cannot proceed. Nothing else has been written yet.
            return WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.TargetProtected);
        }

        // Post-creation protection re-check: if a sheet was created on the
        // create path, validate that the new worksheet and workbook structure
        // still permit writes before attempting headerRange.Value2. If
        // protection would block later writes, roll back the sheet creation.
        if (!adopt)
        {
            if (IsWorksheetProtected(target) || IsWorkbookStructureProtected(workbook))
            {
                // Roll back the sheet creation: delete the sheet we just
                // added, then refuse. Nothing else has been written yet.
                RollBackCreatedSheet(target);
                return WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.TargetProtected);
            }
        }

        var wroteTitle = false;
        var wroteHeader = false;
        var createdTable = false;
        ListObject? createdDataTable = null;
        var createdConfigSheet = false;
        var wrotePlotAnchor = false;

        try
        {
            // The reserved row is written FIRST (ADR-0030 D4), so the title cell
            // exists before the header range is resolved from row 2 and before the
            // table is created from it. Writing rather than inserting is deliberate:
            // the adopt path only runs on a pristine sheet, so row 1 is already empty
            // and there is no structural mutation for a refusal to leave behind.
            WriteTitleRow(
                target,
                GanttCatalogues.SettingDefault("ChartTitle") ?? GanttWorkbookContract.DefaultChartTitle);
            wroteTitle = true;
            WriteHeaderRow(target);
            wroteHeader = true;

            // The table is split into "add" and "present" so `createdTable` can be set
            // BETWEEN them. Everything after `ListObjects.Add` can throw -- the host
            // refuses a rename or a presentation write on a locked or otherwise hostile
            // target -- and each of those throws while `createdTable` was still false,
            // so the catch block rolled back only the header row and left a real
            // `tblGanttData` on the sheet behind a refusal that promised no mutation.
            // Recording the creation immediately after the Add makes every one of those
            // paths roll the table back too.
            //
            // The reference itself is recorded as well, because a table whose rename
            // never landed is not findable by name and the rollback must still remove
            // it -- see RollBackDataTable.
            createdDataTable = AddDataTable(target);
            createdTable = true;
            PresentDataTable(createdDataTable);
            ApplyColumnPresentation(createdDataTable);

            CreateConfigurationSheet(sheets, target);
            createdConfigSheet = true;
            ConfigWriteOutcome catalogueOutcome = _catalogueWriter.Write();
            if (!catalogueOutcome.Succeeded)
            {
                RollBackForCatalogueRefusal(
                    target,
                    sheets,
                    catalogueOutcome,
                    wroteHeader,
                    wroteTitle,
                    createdTable,
                    createdConfigSheet,
                    createdDataTable);

                // Map each reason explicitly. A catch-all "everything else is
                // TargetProtected" would tell the user their sheet is protected
                // when the actual cause was a missing config sheet, an
                // unpreservable catalogue, or a host write failure.
                InitialiseRefusalReason reason = catalogueOutcome.Refusal switch
                {
                    ConfigWriteRefusalReason.NoActiveWorkbook
                        => InitialiseRefusalReason.NoActiveWorkbook,
                    ConfigWriteRefusalReason.TargetProtected
                        => InitialiseRefusalReason.TargetProtected,
                    ConfigWriteRefusalReason.ConfigSheetMissing
                        => InitialiseRefusalReason.ConfigSheetExists,
                    ConfigWriteRefusalReason.CataloguePreservationInvalid
                        => InitialiseRefusalReason.CatalogueDrift,
                    ConfigWriteRefusalReason.HostRejected
                        => InitialiseRefusalReason.CatalogueWriteFailed,
                    null => InitialiseRefusalReason.CatalogueWriteFailed,
                    _ => InitialiseRefusalReason.CatalogueWriteFailed,
                };
                return WorkbookInitialiseOutcome.Refused(reason);
            }

            WritePlotAnchorName(target);
            wrotePlotAnchor = true;
            TypeOptionsMaterialiseOutcome typeOptionsOutcome = _typeOptionsMaterialiser.Materialise();
            if (!typeOptionsOutcome.Succeeded)
            {
                RollBackPlotAnchorName(target);
                RollBackConfigurationSheet(sheets);
                RollBackDataTable(target, createdDataTable);
                RollBackHeaderRow(target);
                RollBackTitleRow(target);
                return typeOptionsOutcome.Refusal == TypeOptionsRefusalReason.NoActiveWorkbook
                    ? WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.NoActiveWorkbook)
                    : typeOptionsOutcome.Refusal == TypeOptionsRefusalReason.TargetProtected
                        ? WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.TargetProtected)
                        : WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.CatalogueDrift);
            }
        }
        catch
        {
            if (wrotePlotAnchor)
            {
                RollBackPlotAnchorName(target);
            }

            if (createdConfigSheet)
            {
                RollBackConfigurationSheet(sheets);
            }

            if (createdTable)
            {
                RollBackDataTable(target, createdDataTable);
            }

            if (wroteHeader)
            {
                RollBackHeaderRow(target);
            }

            if (wroteTitle)
            {
                RollBackTitleRow(target);
            }

            throw;
        }

        return adopt
            ? WorkbookInitialiseOutcome.Adopted(label)
            : WorkbookInitialiseOutcome.CreatedNew(label);
    }

    /// <summary>
    /// Rolls back every mutation above when the catalogue writer returns a
    /// typed refusal: the configuration sheet (with any part-written
    /// catalogues), the data table, and the header row. The plot anchor was
    /// not written yet, so there is nothing to roll back there.
    /// </summary>
    /// <param name="target">The Gantt worksheet.</param>
    /// <param name="sheets">The workbook's sheets.</param>
    /// <param name="catalogueOutcome">The refusing catalogue outcome (for parity, not surfaced).</param>
    /// <param name="wroteHeader">Whether the header row was written.</param>
    /// <param name="wroteTitle">Whether the reserved title row was written.</param>
    /// <param name="createdTable">Whether the data table was created.</param>
    /// <param name="createdConfigSheet">Whether the configuration sheet was created.</param>
    /// <param name="createdDataTable">
    /// The table returned by <see cref="AddDataTable"/>, so a table that never received
    /// its name is still removed rather than looked up and missed.
    /// </param>
    private void RollBackForCatalogueRefusal(
        Worksheet target,
        Sheets sheets,
        ConfigWriteOutcome catalogueOutcome,
        bool wroteHeader,
        bool wroteTitle,
        bool createdTable,
        bool createdConfigSheet,
        ListObject? createdDataTable)
    {
        // The refusal is translated by the caller; the parameter keeps the
        // refusal's evidence attached to the rollback for debugging.
        _ = catalogueOutcome;
        if (createdConfigSheet)
        {
            RollBackConfigurationSheet(sheets);
        }

        if (createdTable)
        {
            RollBackDataTable(target, createdDataTable);
        }

        if (wroteHeader)
        {
            RollBackHeaderRow(target);
        }

        if (wroteTitle)
        {
            RollBackTitleRow(target);
        }
    }

    /// <summary>
    /// Determines whether the worksheet is pristine enough for GanttCreator to
    /// take over in place. The cell checks use the underlying used range, not
    /// formatted display text. Every collection is read once through a local
    /// COM proxy; the predicate never mutates workbook state.
    /// </summary>
    /// <param name="worksheet">The candidate target worksheet.</param>
    /// <param name="workbook">The workbook that owns the candidate.</param>
    /// <param name="application">The Excel application used for <c>CountA</c>.</param>
    /// <returns><see langword="true"/> when the worksheet is pristine.</returns>
    private bool IsPristy(Worksheet worksheet, Workbook workbook, Application application)
    {
        Excel.Range? usedRange = worksheet.UsedRange;
        if (usedRange is not null)
        {
            if (!string.Equals(GetUsedRangeAddress(usedRange), "$A$1", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            WorksheetFunction functions = application.WorksheetFunction;
            if (functions.CountA(usedRange) != 0)
            {
                return false;
            }
        }

        Shapes shapes = worksheet.Shapes;
        Comments comments = worksheet.Comments;
        CommentsThreaded threadedComments = worksheet.CommentsThreaded;
        Names sheetNames = worksheet.Names;
        ListObjects listObjects = worksheet.ListObjects;
        QueryTables queryTables = worksheet.QueryTables;
        Hyperlinks hyperlinks = worksheet.Hyperlinks;
        Names workbookNames = workbook.Names;

        return shapes.Count == 0
            && comments.Count == 0
            && threadedComments.Count == 0
            && sheetNames.Count == 0
            && workbookNames.Count == 0
            && listObjects.Count == 0
            && GetPivotTableCount(worksheet) == 0
            && queryTables.Count == 0
            && hyperlinks.Count == 0;
    }

    /// <summary>
    /// Determines whether any worksheet in the workbook already contains a
    /// list object named <c>tblGanttData</c>. This runs before any sheets or
    /// headers are created and inspects every worksheet.
    /// </summary>
    /// <param name="sheets">The workbook's sheets.</param>
    /// <returns><see langword="true"/> when the table name is already taken on any worksheet.</returns>
    private bool AnyWorksheetContainsTable(Sheets sheets)
    {
        var count = sheets.Count;
        for (var index = 1; index <= count; index++)
        {
            // Only worksheets carry ListObjects; chart sheets are skipped
            // because the cast to Excel.Worksheet fails for them.
            var sheet = GetSheetAt(sheets, index);
            if (sheet is Worksheet worksheet)
            {
                ListObjects listObjects = worksheet.ListObjects;
                var tableCount = listObjects.Count;
                for (var tableIndex = 1; tableIndex <= tableCount; tableIndex++)
                {
                    ListObject table = GetTableAt(listObjects, tableIndex);
                    if (string.Equals(
                        table.Name,
                        GanttTableSchema.TableName,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }
        }

        return false;
    }

    /// <summary>
    /// Resolves the deterministic Gantt sheet label: the contract label when
    /// no other sheet uses it, otherwise Excel's own <c>" (n)"</c> suffix
    /// convention with the first free <c>n</c>.
    /// </summary>
    /// <param name="sheets">The workbook's sheets.</param>
    /// <param name="excludeSheetName">
    /// The candidate target's own current name during an adopt (it is about
    /// to be renamed), or <see langword="null"/> on the create path.
    /// </param>
    /// <returns>The available, culture-invariant label.</returns>
    private string ResolveAvailableLabel(Sheets sheets, string? excludeSheetName)
    {
        var baseLabel = GanttWorkbookContract.GanttSheetLabel;
        if (!NameTakenByOtherSheet(sheets, baseLabel, excludeSheetName))
        {
            return baseLabel;
        }

        for (var suffix = 2; ; suffix++)
        {
            var candidate = string.Create(
                CultureInfo.InvariantCulture, $"{baseLabel} ({suffix})");
            if (!NameTakenByOtherSheet(sheets, candidate, excludeSheetName))
            {
                return candidate;
            }
        }
    }

    /// <summary>
    /// Determines whether <paramref name="name"/> is used by any sheet other
    /// than the one whose name equals <paramref name="excludeSheetName"/>.
    /// Workbook.Sheets includes chart sheets; the Sheets collection exposes
    /// Name on every sheet type, so chart-sheet names participate in
    /// availability validation.
    /// </summary>
    /// <param name="sheets">The workbook's sheets.</param>
    /// <param name="name">The candidate name.</param>
    /// <param name="excludeSheetName">The name to ignore, or <see langword="null"/>.</param>
    /// <returns><see langword="true"/> when the name is taken.</returns>
    private bool NameTakenByOtherSheet(Sheets sheets, string name, string? excludeSheetName)
    {
        var count = sheets.Count;
        for (var index = 1; index <= count; index++)
        {
            // Every sheet in the collection exposes Name, including chart
            // sheets (which are not Excel.Worksheet). Use the unsealed cast
            // chain via GetSheetAt to keep the seam consistent with the test
            // seam pattern, then widen to object only for Name access on
            // chart sheets that are not Excel.Worksheet.
            var sheet = GetSheetAt(sheets, index);
            var sheetName = sheet switch
            {
                Worksheet worksheet => worksheet.Name,
                _ => (string)sheet.GetType().InvokeMember(
                    "Name",
                    System.Reflection.BindingFlags.GetProperty,
                    null,
                    sheet,
                    null,
                    CultureInfo.InvariantCulture)!,
            };
            if (excludeSheetName is not null
                && string.Equals(sheetName, excludeSheetName, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (string.Equals(sheetName, name, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Creates the new target worksheet after the active sheet.
    /// </summary>
    /// <param name="sheets">The workbook's sheets.</param>
    /// <param name="activeWorksheet">The current active worksheet, or <see langword="null"/>.</param>
    /// <returns>The created worksheet.</returns>
    private static Worksheet CreateTargetSheet(Sheets sheets, Worksheet? activeWorksheet)
    {
        Worksheet created = activeWorksheet is null
            ? (Worksheet)sheets.Add()
            : (Worksheet)sheets.Add(After: activeWorksheet);
        return created;
    }

    /// <summary>
    /// Deletes a newly-created worksheet that must not persist because
    /// protection would block later writes. Called only on the create path
    /// after the sheet has been added but before any content is written.
    /// </summary>
    /// <param name="target">The worksheet to remove.</param>
    private void RollBackCreatedSheet(Worksheet target)
    {
        if (_application is null)
        {
            return;
        }
        var original = _application.DisplayAlerts;
        try
        {
            _application.DisplayAlerts = false;
            target.Delete();
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // Best-effort, like every rollback in this class: a host
            // that refuses the deletion must not replace the typed
            // refusal this rollback is part of with a raw COM error.
        }
        finally
        {
            _application.DisplayAlerts = original;
        }
    }

    /// <summary>
    /// Removes the plot-anchor defined name that <see cref="WritePlotAnchorName"/>
    /// created on <paramref name="target"/>. No-op when the name does not exist.
    /// </summary>
    /// <param name="target">The Gantt worksheet.</param>
    private static void RollBackPlotAnchorName(Worksheet target)
    {
        try
        {
            Names names = target.Names;
            Name name = names.Item(
                GanttWorkbookContract.PlotAnchorDefinedName);
            name.Delete();
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // The defined name may not exist; nothing to roll back.
        }
    }

    /// <summary>
    /// Deletes the configuration worksheet named
    /// <c>GanttWorkbookContract.ConfigSheetName</c>, if it exists. Used to roll
    /// back <see cref="CreateConfigurationSheet"/> when a later mutation fails.
    /// </summary>
    /// <param name="sheets">The workbook's sheets.</param>
    private void RollBackConfigurationSheet(Sheets sheets)
    {
        if (_application is null)
        {
            return;
        }
        var original = _application.DisplayAlerts;
        try
        {
            _application.DisplayAlerts = false;
            var count = sheets.Count;
            for (var index = 1; index <= count; index++)
            {
                var sheet = GetSheetAt(sheets, index);
                var sheetName = sheet switch
                {
                    Worksheet worksheet => worksheet.Name,
                    _ => (string)sheet.GetType().InvokeMember(
                        "Name",
                        System.Reflection.BindingFlags.GetProperty,
                        null,
                        sheet,
                        null,
                        CultureInfo.InvariantCulture)!,
                };
                if (string.Equals(
                    sheetName,
                    GanttWorkbookContract.ConfigSheetName,
                    StringComparison.OrdinalIgnoreCase))
                {
                    if (sheet is Worksheet configSheet)
                    {
                        configSheet.Delete();
                    }

                    break;
                }
            }
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // Best-effort, like RollBackDataTable: a host that refuses
            // the configuration sheet's deletion must not mask the
            // exception that triggered the rollback -- the `throw;` in
            // Initialise's catch block carries that original exception
            // to the command boundary, which logs it -- and must not
            // abort the table, header, and title rollbacks after it.
        }
        finally
        {
            _application.DisplayAlerts = original;
        }
    }

    /// <summary>
    /// Deletes the add-in's own table on <paramref name="target"/>. Used to roll back
    /// <see cref="AddDataTable"/> when a later mutation fails.
    /// </summary>
    /// <param name="target">The Gantt worksheet.</param>
    /// <param name="created">
    /// The table returned by <see cref="AddDataTable"/>, or <see langword="null"/> to
    /// look the table up by name.
    /// </param>
    /// <remarks>
    /// <para>
    /// The reference is what makes a PARTIAL table removable. The lookup matches on
    /// <c>GanttTableSchema.TableName</c>, so a table whose rename never landed -- still
    /// carrying a host-assigned name -- is invisible to it, and the rollback would
    /// silently keep it. That is the exact state a refused
    /// <c>table.Name</c> write leaves behind, and it is what makes a retry then meet a
    /// table the user never asked for.
    /// </para>
    /// <para>
    /// The lookup remains the fallback for the callers that reach this without a
    /// reference, so the behaviour is unchanged where the table is fully named.
    /// </para>
    /// </remarks>
    private void RollBackDataTable(Worksheet target, ListObject? created = null)
    {
        if (_application is null)
        {
            return;
        }
        var original = _application.DisplayAlerts;
        try
        {
            _application.DisplayAlerts = false;

            // Unhide before deleting: deleting the table does not unhide the
            // worksheet column behind it, so restoring afterwards would find no
            // table to read the columns from and would leave the user's engine
            // columns hidden after a refused Initialise.
            if ((created ?? FindDataTable(target)) is { } owned)
            {
                RestoreColumnVisibility(owned);
                owned.Delete();
            }
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // The table may already have been removed; nothing to roll back.
        }
        finally
        {
            _application.DisplayAlerts = original;
        }
    }

    /// <summary>
    /// Clears the title cell written by <see cref="WriteTitleRow"/> in the reserved
    /// row, so a refused Initialise leaves no add-in-authored content behind.
    /// </summary>
    /// <param name="target">The Gantt worksheet.</param>
    /// <remarks>
    /// <b>Clearing a cell, not deleting a row.</b> The reserved row was written into,
    /// never structurally inserted, so there is no row to remove and no shifted table
    /// to restore. That is the whole reason the layout is produced by writing rather
    /// than by <c>Rows(1).Insert</c>: an inserted row would need a matching delete on
    /// every refusal path, and a missed one would leave the sheet altered after a
    /// command that reported no change.
    /// </remarks>
    private void RollBackTitleRow(Worksheet target)
    {
        // The whole SPAN is cleared, not just its top-left cell. Clearing one cell
        // would leave a title behind if a write had partially succeeded across the
        // range, and the rollback exists precisely for the case where the sheet was
        // already touched (ADR-0032 D1).
        try
        {
            Excel.Range titleRange = GetTitleRange(target);
            titleRange.ClearContents();
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // Best-effort: a host that refuses the clear must not
            // mask the original exception nor abort the remaining
            // rollbacks.
        }
    }

    /// <summary>
    /// Clears the header row content written by <see cref="WriteHeaderRow"/> on
    /// <paramref name="target"/> and unhides the columns it hid. Used to roll back
    /// the header row when a later mutation fails.
    /// </summary>
    /// <param name="target">The Gantt worksheet.</param>
    /// <remarks>
    /// The hidden state is restored because hiding a column is a <em>worksheet</em>
    /// change that outlives the table: deleting the table (see
    /// <see cref="RollBackDataTable"/>) does not unhide the worksheet column behind it.
    /// Rolling back the table and the header while leaving eleven engine columns hidden
    /// would leave the user's sheet altered after a refused Initialise, which is
    /// exactly the zero-mutation guarantee the refusals promise.
    /// </remarks>
    private void RollBackHeaderRow(Worksheet target)
    {
        try
        {
            Excel.Range headerRange = GetHeaderRange(
                target,
                GanttTableSchema.Default.Columns.Count);
            headerRange.ClearContents();
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // Best-effort: a host that refuses the clear must not
            // mask the original exception nor abort the remaining
            // rollbacks.
        }

        // Best-effort: the table may already have been deleted by the preceding
        // rollback, in which case RollBackDataTable restored the visibility already.
        if (FindDataTable(target) is { } remaining)
        {
            RestoreColumnVisibility(remaining);
        }
    }

    /// <summary>
    /// Unhides every column of <paramref name="table"/>, undoing
    /// <see cref="ApplyColumnPresentation"/>. Every column is restored to visible
    /// because a schema column that is visible to the user must be visible, and one
    /// hidden by the failed attempt is an artefact of that attempt rather than
    /// something the user chose.
    /// </summary>
    /// <remarks>
    /// Takes <paramref name="table"/> rather than re-finding it: on the
    /// <see cref="RollBackDataTable"/> path the table is deleted immediately after
    /// this runs, so a second lookup would have nothing to find, and on the
    /// <see cref="RollBackHeaderRow"/> path the table has already been deleted.
    /// </remarks>
    private void RestoreColumnVisibility(ListObject table)
    {
        // The table may already have been removed, or the host may refuse the read.
        // Either way there is no add-in-owned visibility left to restore, and a
        // rollback must not throw over a best-effort restore.
        try
        {
            ListColumns listColumns = GetTableColumns(table);
            var count = listColumns.Count;
            for (var columnIndex = 1; columnIndex <= count; columnIndex++)
            {
                ListColumn column = GetColumnAt(listColumns, columnIndex);
                Excel.Range columnRange = column.Range;
                Excel.Range entireColumn = GetEntireColumn(columnRange);
                if (ReadHiddenFlag(entireColumn))
                {
                    entireColumn.Hidden = false;
                }
            }
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // The table is gone or the host refused the read; there is no
            // add-in-owned visibility left to restore.
        }
    }

    /// <summary>
    /// Finds the add-in's own table on <paramref name="target"/>, or
    /// <see langword="null"/> when it is absent or already deleted.
    /// </summary>
    /// <param name="target">The Gantt worksheet.</param>
    /// <returns>The named table, or <see langword="null"/>.</returns>
    private ListObject? FindDataTable(Worksheet target)
    {
        try
        {
            // One proxy in a local, used for BOTH the count and the indexed lookup.
            // Reaching `target.ListObjects` twice re-enters the COM property getter and
            // yields two RCWs for one host collection, which is exactly the chained
            // access the COM-ownership rule forbids.
            ListObjects listObjects = target.ListObjects;
            var tableCount = listObjects.Count;
            for (var index = 1; index <= tableCount; index++)
            {
                ListObject table = GetTableAt(listObjects, index);
                if (string.Equals(
                    table.Name,
                    GanttTableSchema.TableName,
                    StringComparison.OrdinalIgnoreCase))
                {
                    return table;
                }
            }
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // The worksheet is gone; there is no table to find.
        }

        return null;
    }

    /// <summary>
    /// Determines whether the worksheet is protected against edits.
    /// </summary>
    /// <param name="worksheet">The worksheet to inspect.</param>
    /// <returns><see langword="true"/> when the worksheet is protected.</returns>
    private static bool IsWorksheetProtected(Worksheet worksheet) =>
        // The ProtectContents flag indicates cell-level protection is active.
        // A protected worksheet blocks headerRange.Value2 writes.
        worksheet.ProtectContents;

    /// <summary>
    /// Determines whether the workbook structure is protected, which blocks
    /// sheet creation, deletion, rename, and move operations.
    /// </summary>
    /// <param name="workbook">The workbook to inspect.</param>
    /// <returns><see langword="true"/> when the workbook structure is protected.</returns>
    private static bool IsWorkbookStructureProtected(Workbook workbook) =>
        // The ProtectStructure flag indicates workbook-structure protection is
        // active, blocking sheet-level structural changes.
        workbook.ProtectStructure;

    /// <summary>
    /// Writes the <c>tblGanttData</c> header row from the Core schema, in the
    /// exact contract order, as one array assignment (one COM call).
    /// </summary>
    /// <param name="target">The Gantt worksheet.</param>
    private void WriteHeaderRow(Worksheet target)
    {
        IReadOnlyList<GanttTableColumn> columns =
            GanttTableSchema.Default.Columns;
        // CA1814: Excel's Range.Value2 accepts only a rectangular object
        // array (a COM SAFEARRAY of VARIANT); a jagged array does not marshal
        // to it. The multidimensional form is the requirement, not a style
        // choice.
#pragma warning disable CA1814
        var values = new object[1, columns.Count];
        for (var index = 0; index < columns.Count; index++)
        {
            values[0, index] = columns[index].Name;
        }

        Excel.Range headerRange = GetHeaderRange(target, columns.Count);
        headerRange.Value2 = values;
#pragma warning restore CA1814
    }

    /// <summary>
    /// Adds the <c>tblGanttData</c> table over the header row, and returns it
    /// WITHOUT naming or presenting it.
    /// </summary>
    /// <param name="target">The Gantt worksheet.</param>
    /// <returns>The added table.</returns>
    /// <remarks>
    /// <para>
    /// Deliberately stops at the Add. Everything after it can throw, and the caller
    /// must be able to record the creation before any of it runs -- otherwise a
    /// refused rename leaves a real table on a sheet whose rollback promised no
    /// mutation, and the retry then finds a table the user never asked for.
    /// </para>
    /// <para>
    /// Splitting here is what makes the partial state <em>findable</em> as well:
    /// <see cref="RollBackDataTable(Worksheet, ListObject?)"/> takes the returned
    /// reference rather than searching by name, so a table whose rename never landed
    /// is still removed.
    /// </para>
    /// </remarks>
    private ListObject AddDataTable(Worksheet target)
    {
        var columnCount = GanttTableSchema.Default.Columns.Count;
        Excel.Range tableRange = GetHeaderRange(target, columnCount);
        ListObjects listObjects = target.ListObjects;
        return listObjects.Add(
            XlListObjectSourceType.xlSrcRange,
            tableRange,
            Type.Missing,
            XlYesNoGuess.xlYes,
            Type.Missing);
    }

    /// <summary>
    /// Names the table from the Core schema constant and applies the R2.7 appearance
    /// settings (no autofilter dropdowns, no banded rows) that keep the data panel
    /// visually neutral for the live Gantt (ADR-0007 D8).
    /// </summary>
    /// <param name="table">The added table.</param>
    private static void PresentDataTable(ListObject table)
    {
        table.Name = GanttTableSchema.TableName;
        table.ShowAutoFilter = false;
        table.ShowTableStyleRowStripes = false;
        table.ShowTableStyleColumnStripes = false;
    }

    /// <summary>
    /// Applies each column's <see cref="GanttColumnAccess"/> classification to the
    /// live table: engine columns are hidden and every non-authoring column's cells
    /// carry the locked format (R4.7C D1/D2, ADR-0029 D7/D8).
    /// </summary>
    /// <remarks>
    /// <para>
    /// This runs on Initialise only. The acceptance test "un-hiding one and
    /// refreshing restores it" is a <em>repair</em>, and repair on Refresh is
    /// R4.8A's orchestration; wiring it here would give this adapter a second
    /// trigger it does not own. What this row guarantees is that a freshly
    /// initialised workbook is correct, and that the classification driving it is
    /// the code-owned schema rather than a hand-written column list.
    /// </para>
    /// <para>
    /// <b>Hidden, not just locked.</b> A cell's <c>Locked</c> flag has no effect
    /// until the sheet is protected, and the add-in refuses a protected target
    /// rather than protecting one (ADR-0008 D4). So <c>Locked</c> alone would
    /// leave the engine columns fully visible on an ordinary sheet; hiding them
    /// is what actually presents the table the product describes.
    /// </para>
    /// <para>
    /// <b>Why the whole worksheet column is hidden.</b> Excel has no per-table
    /// hidden flag: a ListObject column is hidden by hiding the worksheet column
    /// behind it, which is what the Excel UI's own Hide command does. The
    /// alternative, zero column width, leaves a visible sliver and breaks
    /// print layout, so it is rejected.
    /// </para>
    /// </remarks>
    private void ApplyColumnPresentation(ListObject table)
    {
        ListColumns columns = GetTableColumns(table);
        IReadOnlyList<GanttTableColumn> schema = GanttTableSchema.Default.Columns;
        for (var index = 1; index <= schema.Count; index++)
        {
            ListColumn column = GetColumnAt(columns, index);
            Excel.Range columnRange = column.Range;
            columnRange.Locked = schema[index - 1].IsLocked;
            Excel.Range entireColumn = GetEntireColumn(columnRange);

            // `Range.Hidden` is declared `object` in this PIA (verified by reflection
            // over the installed Microsoft.Office.Interop.Excel 14.0.1), so comparing it
            // to a `bool` with `!=` boxed the right-hand side and compared references:
            // always true, whatever the host actually reported. Every column was
            // therefore rewritten on every Initialise, which marks the workbook dirty
            // for a sheet that is already in the intended state. The value is unwrapped
            // to a real `bool` first, so a column already carrying the schema's
            // visibility is left alone.
            var wanted = schema[index - 1].IsHidden;
            if (ReadHiddenFlag(entireColumn) != wanted)
            {
                entireColumn.Hidden = wanted;
            }
        }
    }

    /// <summary>
    /// Reads <c>Range.Hidden</c> as a <see cref="bool"/>.
    /// </summary>
    /// <param name="range">The worksheet column behind a table column.</param>
    /// <returns>
    /// The column's hidden state. A value the host does not report as a boolean reads
    /// as <see langword="false"/>, which matches the ordinary visible-column default
    /// and therefore causes the write that makes the state correct.
    /// </returns>
    /// <remarks>
    /// The PIA types <c>Hidden</c> as <see cref="object"/>, so the value arrives boxed
    /// and must be unwrapped before any comparison. An unrecognised or absent value
    /// falls back to "not hidden" rather than being propagated: a column whose real
    /// state is unknown is written, and a write is the safe direction because it makes
    /// the column match the schema instead of leaving a possibly-wrong state alone.
    /// </remarks>
    private static bool ReadHiddenFlag(Excel.Range range) => range.Hidden switch
    {
        bool flag => flag,
        null => false,
        _ => bool.TryParse(
            Convert.ToString(range.Hidden, CultureInfo.InvariantCulture),
            out bool parsed) && parsed,
    };

    /// <summary>
    /// Creates the configuration worksheet directly after the Gantt sheet,
    /// names it from the Core contract, and sets <c>xlSheetVeryHidden</c>.
    /// </summary>
    /// <param name="sheets">The workbook's sheets.</param>
    /// <param name="target">The Gantt worksheet the config sheet follows.</param>
    private static void CreateConfigurationSheet(Sheets sheets, Worksheet target)
    {
        var config = (Worksheet)sheets.Add(After: target);
        config.Name = GanttWorkbookContract.ConfigSheetName;
        config.Visible = XlSheetVisibility.xlSheetVeryHidden;
    }

    /// <summary>
    /// Writes the table title into the reserved row above the header
    /// (ADR-0030 D4/D6, R4.7I slice 1).
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>The title is a cell value, not a drawn shape.</b> It is editable, printable,
    /// and survives save/reopen without a shape. <c>SceneBuilder</c> emits no title
    /// primitive under the <c>LiveExcel</c> profile (D6); the export profile still
    /// emits one.
    /// </para>
    /// <para>
    /// <b>No row is structurally inserted.</b> The adopt path only runs on a pristine
    /// sheet, so row 1 is already empty and writing into it produces exactly the
    /// reserved layout the plan's Option A diagram shows. That is strictly safer than
    /// <c>Rows(1).Insert</c>: there is no structural mutation to roll back, so no
    /// refusal path can leave a stray row and a shifted table behind it. The header
    /// range then starts at row 2 and the table is created from it.
    /// </para>
    /// <para>
    /// The value comes from the stored <c>ChartTitle</c> setting rather than a local
    /// literal, so the cell and the export composition cannot name different charts.
    /// </para>
    /// </remarks>
    private void WriteTitleRow(Worksheet target, string chartTitle)
    {
        // ONE cell, above Description.
        //
        // This previously wrote to a 1-by-5 `Resize` range on the stated belief that
        // "Excel resolves a range assignment to that range's top-left cell, so D2:H2
        // receives the title in D2". That premise is FALSE: assigning `Value2` to a
        // multi-cell range writes the value into EVERY cell, which is why the live
        // sheet showed "Gantt Chart" once per visible column. The fix is not to reach
        // into the range's one-based cell indexer - which ADR-0032 D1 forbids as
        // hardcoded addressing - but to make the RANGE itself one cell wide, derived
        // from `GanttSheetLayout` like every other piece of sheet geometry.
        Excel.Range titleRange = GetTitleRange(target);
        titleRange.Value2 = chartTitle;
    }

    /// <summary>
    /// Writes the sheet-scoped plot-anchor defined name: the cell one column
    /// right of the table's last column, on the header row.
    /// </summary>
    /// <param name="target">The Gantt worksheet.</param>
    private static void WritePlotAnchorName(Worksheet target)
    {
        Names names = target.Names;
        _ = names.Add(
            GanttWorkbookContract.PlotAnchorDefinedName,
            GanttSheetLayout.BuildPlotAnchorRefersTo(target.Name));
    }

    /// <summary>
    /// Returns the sheet at the one-based index. Test seam over the COM
    /// parameterised <c>Sheets.Item</c> property (see the type remarks).
    /// </summary>
    /// <param name="sheets">The workbook's sheets.</param>
    /// <param name="index">The one-based sheet index.</param>
    /// <returns>The worksheet at the index.</returns>
    internal virtual object GetSheetAt(Sheets sheets, int index)
        => sheets[index];

    /// <summary>
    /// Returns the used-range address. This indexed COM property cannot be
    /// used from a Moq expression tree, so contract tests substitute this
    /// seam and live Office proves the real call.
    /// </summary>
    /// <param name="usedRange">The used range to inspect.</param>
    /// <returns>The Excel A1-style address.</returns>
    internal virtual string GetUsedRangeAddress(Excel.Range usedRange) => usedRange.Address;

    /// <summary>
    /// Returns the number of pivot tables on the worksheet. This indexed COM
    /// method cannot be used from a Moq expression tree, so contract tests
    /// substitute this seam and live Office proves the real call.
    /// </summary>
    /// <param name="worksheet">The worksheet to inspect.</param>
    /// <returns>The pivot-table count.</returns>
    internal virtual int GetPivotTableCount(Worksheet worksheet) => worksheet.PivotTables().Count;

    /// <summary>
    /// Returns the list object at the one-based index. Test seam over the COM
    /// parameterised <c>ListObjects.Item</c> property (see the type remarks).
    /// </summary>
    /// <param name="listObjects">The worksheet's list objects.</param>
    /// <param name="index">The one-based table index.</param>
    /// <returns>The list object at the index.</returns>
    internal virtual ListObject GetTableAt(ListObjects listObjects, int index)
        => listObjects[index];

    /// <summary>
    /// Returns the header row range over the requested column count. Test seam
    /// over the COM parameterised <c>Range.Item</c> and <c>Range.Resize</c>
    /// properties (see the type remarks).
    /// </summary>
    /// <param name="target">The Gantt worksheet.</param>
    /// <param name="columnCount">The header column count.</param>
    /// <returns>The one-row range spanning the header columns.</returns>
    internal virtual Excel.Range GetHeaderRange(Worksheet target, int columnCount)
        => target.Cells[GanttSheetLayout.HeaderRowIndex, 1].Resize[1, columnCount];

    /// <summary>
    /// Returns the SINGLE cell holding the table title, in the reserved row directly
    /// above the header and above the <c>Description</c> column. Test seam over the
    /// COM parameterised <c>Range.Resize</c> property, mirroring
    /// <see cref="GetHeaderRange"/>.
    /// </summary>
    /// <param name="target">The Gantt worksheet.</param>
    /// <returns>The one-cell range for the title.</returns>
    /// <remarks>
    /// <para>
    /// The span is <b>one cell</b>, and that is the fix rather than a detail. It was
    /// five, so the title was assigned to D2:H2 and Excel wrote the value into every
    /// one of those cells - the live sheet showed the title duplicated across the
    /// table. A range assignment resolves to each cell in the range, not to the
    /// range's top-left cell, which is what the previous comment here asserted.
    /// </para>
    /// <para>
    /// The owner asked for the title directly above <c>Description</c>, the column the
    /// table is read by. The row and the column both come from
    /// <see cref="GanttSheetLayout"/>, so this seam cannot disagree with the layout
    /// authority about where the title lives. A title wider than its column is a
    /// merge or a left-aligned overflow, not a value written into each cell.
    /// </para>
    /// </remarks>
    internal virtual Excel.Range GetTitleRange(Worksheet target) =>
        target
            .Cells[GanttSheetLayout.TitleRowIndex, GanttSheetLayout.TitleColumnIndex]
            .Resize[1, GanttSheetLayout.TitleColumnSpan];

    /// <summary>
    /// Returns the table's list-column collection. Test seam over the COM
    /// parameterised <c>ListObject.ListColumns</c> collection, so contract tests
    /// can supply columns without a host.
    /// </summary>
    /// <param name="table">The created data table.</param>
    /// <returns>The table's list columns.</returns>
    internal virtual ListColumns GetTableColumns(ListObject table) => table.ListColumns;

    /// <summary>
    /// Returns the list column at the one-based index. Test seam over the COM
    /// parameterised <c>ListColumns.Item</c> property (see the type remarks).
    /// </summary>
    /// <param name="columns">The table's list columns.</param>
    /// <param name="index">The one-based column index.</param>
    /// <returns>The list column at the index.</returns>
    internal virtual ListColumn GetColumnAt(ListColumns columns, int index)
        => columns[index];

    /// <summary>
    /// Returns the whole worksheet column behind a range. Test seam over the
    /// COM parameterised <c>Range.EntireColumn</c> property, which is where the
    /// per-column hidden state actually lives.
    /// </summary>
    /// <param name="range">The column's range.</param>
    /// <returns>The entire worksheet column.</returns>
    internal virtual Excel.Range GetEntireColumn(Excel.Range range) => range.EntireColumn;
}
