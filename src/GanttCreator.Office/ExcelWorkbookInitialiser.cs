using System.Diagnostics;
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
/// <param name="trace">
/// Optional technical-record sink. When supplied, one <c>gantt.initialise steps</c>
/// record splits the mutation sequence — header row, table, the 17-column presentation
/// pass, configuration sheet plus catalogue write, plot anchor, and the TypeOptions
/// validation — because "Initialise feels slow" cannot be answered by the refresh
/// timings, which begin only after all of this has run. Absent in tests, where the
/// timings are not a product behaviour.
/// </param>
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
    ITypeOptionsMaterialiser? typeOptionsMaterialiser = null,
    Action<string>? trace = null) : IWorkbookInitialiser
{
    /// <summary>
    /// Optional technical-record sink. When supplied, one
    /// <c>gantt.initialise steps</c> record splits the mutation sequence, which is the
    /// half of "Initialise feels slow" the refresh timings cannot describe: the table,
    /// the 17-column presentation pass and the catalogue write all happen here.
    /// </summary>
    private readonly Action<string>? _trace = trace;

    private static double Ms(long ticks) => ticks * 1000d / Stopwatch.Frequency;

    /// <summary>The four border indices that form a table's outside edge.</summary>
    private static readonly XlBordersIndex[] _outsideBorderEdges =
    [
        XlBordersIndex.xlEdgeTop,
        XlBordersIndex.xlEdgeBottom,
        XlBordersIndex.xlEdgeLeft,
        XlBordersIndex.xlEdgeRight,
    ];

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
        Worksheet? activeWorksheet = workbook.ActiveSheet as Worksheet;
        bool adopt = activeWorksheet is not null
            && IsPristy(activeWorksheet, workbook, application);

        // Read-only check 3: resolve the label against every other sheet
        // (OrdinalIgnoreCase, matching Excel's own case-insensitive sheet-name
        // uniqueness) before any mutation. On the adopt path the target's own
        // current name is excluded — it is about to be renamed.
        string label = ResolveAvailableLabel(sheets, adopt ? activeWorksheet!.Name : null);

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

        bool wroteHeader = false;
        bool createdTable = false;
        ListObject? createdDataTable = null;
        bool createdConfigSheet = false;
        bool wrotePlotAnchor = false;

        long stepStart = _trace is null ? 0L : Stopwatch.GetTimestamp();
        long headerTicks = 0;
        long tableTicks = 0;
        long columnsTicks = 0;
        long configTicks = 0;
        long anchorTicks = 0;
        long typeTicks = 0;

        try
        {
            WriteHeaderRow(target);
            wroteHeader = true;
            if (_trace is not null)
            {
                long now = Stopwatch.GetTimestamp();
                headerTicks = now - stepStart;
                stepStart = now;
            }

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
            if (_trace is not null)
            {
                long now = Stopwatch.GetTimestamp();
                tableTicks = now - stepStart;
                stepStart = now;
            }

            ApplyColumnPresentation(createdDataTable);
            ApplyLeadingGutterWidth(createdDataTable);
            if (_trace is not null)
            {
                long now = Stopwatch.GetTimestamp();
                columnsTicks = now - stepStart;
                stepStart = now;
            }

            CreateConfigurationSheet(sheets, target);
            createdConfigSheet = true;
            ConfigWriteOutcome catalogueOutcome = _catalogueWriter.Write();
            if (_trace is not null)
            {
                long now = Stopwatch.GetTimestamp();
                configTicks = now - stepStart;
                stepStart = now;
            }
            if (!catalogueOutcome.Succeeded)
            {
                RollBackForCatalogueRefusal(
                    target,
                    sheets,
                    catalogueOutcome,
                    wroteHeader,
                    createdTable,
                    createdConfigSheet,
                    createdDataTable);

                // Map each reason explicitly. A catch-all "everything else is
                // TargetProtected" would tell the user their sheet is protected
                // when the actual cause was a missing config sheet, an
                // unpreservable catalogue, or a host write failure.
                InitialiseRefusalReason reason = catalogueOutcome.Refusal switch
                {
                    ConfigWriteRefusalReason.NoActiveWorkbook => InitialiseRefusalReason.NoActiveWorkbook,
                    ConfigWriteRefusalReason.TargetProtected => InitialiseRefusalReason.TargetProtected,
                    ConfigWriteRefusalReason.ConfigSheetMissing => InitialiseRefusalReason.ConfigSheetExists,
                    ConfigWriteRefusalReason.CataloguePreservationInvalid => InitialiseRefusalReason.CatalogueDrift,
                    ConfigWriteRefusalReason.HostRejected or null => InitialiseRefusalReason.CatalogueWriteFailed,
                    ConfigWriteRefusalReason.ThemeValidationFailed => InitialiseRefusalReason.CatalogueWriteFailed,
                    _ => InitialiseRefusalReason.CatalogueWriteFailed,
                };
                return WorkbookInitialiseOutcome.Refused(reason);
            }

            WritePlotAnchorName(target);
            wrotePlotAnchor = true;
            if (_trace is not null)
            {
                long now = Stopwatch.GetTimestamp();
                anchorTicks = now - stepStart;
                stepStart = now;
            }

            TypeOptionsMaterialiseOutcome typeOptionsOutcome = _typeOptionsMaterialiser.Materialise();
            if (_trace is not null)
            {
                typeTicks = Stopwatch.GetTimestamp() - stepStart;
            }
            if (!typeOptionsOutcome.Succeeded)
            {
                RollBackPlotAnchorName(target);
                RollBackConfigurationSheet(sheets);
                RollBackDataTable(target, createdDataTable);
                RollBackHeaderRow(target);
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

            throw;
        }

        _trace?.Invoke(string.Create(
            CultureInfo.InvariantCulture,
            $"gantt.initialise steps header={Ms(headerTicks):F1} table={Ms(tableTicks):F1}"
            + $" columns={Ms(columnsTicks):F1} config={Ms(configTicks):F1}"
            + $" anchor={Ms(anchorTicks):F1} types={Ms(typeTicks):F1}"));

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
        int count = sheets.Count;
        for (int index = 1; index <= count; index++)
        {
            // Only worksheets carry ListObjects; chart sheets are skipped
            // because the cast to Excel.Worksheet fails for them.
            object sheet = GetSheetAt(sheets, index);
            if (sheet is Worksheet worksheet)
            {
                ListObjects listObjects = worksheet.ListObjects;
                int tableCount = listObjects.Count;
                for (int tableIndex = 1; tableIndex <= tableCount; tableIndex++)
                {
                    ListObject table = GetTableAt(listObjects, tableIndex);
                    if (GanttChartContract.IsDataTableName(table.Name))
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
        string baseLabel = GanttWorkbookContract.GanttSheetLabel;
        if (!NameTakenByOtherSheet(sheets, baseLabel, excludeSheetName))
        {
            return baseLabel;
        }

        for (int suffix = 2; ; suffix++)
        {
            string candidate = string.Create(
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
        int count = sheets.Count;
        for (int index = 1; index <= count; index++)
        {
            // Every sheet in the collection exposes Name, including chart
            // sheets (which are not Excel.Worksheet). Use the unsealed cast
            // chain via GetSheetAt to keep the seam consistent with the test
            // seam pattern, then widen to object only for Name access on
            // chart sheets that are not Excel.Worksheet.
            object sheet = GetSheetAt(sheets, index);
            string sheetName = sheet switch
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
        bool original = _application.DisplayAlerts;
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
    /// <remarks>
    /// <para>
    /// <b>Why this is not <c>static</c> any more.</b> The name lookup is the
    /// <c>Names.Item</c> parameterised property — the same COM shape as
    /// <c>Sheets.Item</c> and <c>ListObjects.Item</c>, which this class already exposes
    /// as <c>internal virtual</c> seams for exactly this reason. It was the one such
    /// call in the rollback path that was made inline, so the TypeOptions rollback could
    /// not be exercised at all: the refusing path threw a <see cref="NullReferenceException"/>
    /// under a Moq proxy, which is neither the production outcome nor the one the
    /// <c>COMException</c> catch models.
    /// </para>
    /// <para>
    /// The lookup is also null-tolerant now. Excel throws for a name that does not exist
    /// (caught below), but the documented contract is "no-op when the name does not
    /// exist", and a null answer is that case rather than an exception.
    /// </para>
    /// </remarks>
    private void RollBackPlotAnchorName(Worksheet target)
    {
        try
        {
            Names names = target.Names;
            Name? name = FindName(names, GanttWorkbookContract.PlotAnchorDefinedName);
            name?.Delete();
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // The defined name may not exist; nothing to roll back.
        }
    }

    /// <summary>
    /// Returns the sheet-scoped defined name with the given text, or
    /// <see langword="null"/> when the sheet carries none. Test seam over the COM
    /// parameterised <c>Names.Item</c> property.
    /// </summary>
    /// <param name="names">The worksheet's sheet-scoped names.</param>
    /// <param name="nameText">The defined name to resolve.</param>
    /// <returns>The resolved name, or <see langword="null"/>.</returns>
    internal virtual Name? FindName(Names names, string nameText)
    {
        ArgumentNullException.ThrowIfNull(names);
        return names.Item(nameText);
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
        bool original = _application.DisplayAlerts;
        try
        {
            _application.DisplayAlerts = false;
            int count = sheets.Count;
            for (int index = 1; index <= count; index++)
            {
                object sheet = GetSheetAt(sheets, index);
                string sheetName = sheet switch
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
                if (GanttWorkbookContract.IsConfigSheetName(sheetName))
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
        bool original = _application.DisplayAlerts;
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
            int count = listColumns.Count;
            for (int columnIndex = 1; columnIndex <= count; columnIndex++)
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
            int tableCount = listObjects.Count;
            for (int index = 1; index <= tableCount; index++)
            {
                ListObject table = GetTableAt(listObjects, index);
                if (GanttChartContract.IsDataTableName(table.Name))
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
        object[,] values = new object[1, columns.Count];
        for (int index = 0; index < columns.Count; index++)
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
        int columnCount = GanttTableSchema.Default.Columns.Count;
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
    /// The tolerance within which a measured gutter width counts as already correct.
    /// </summary>
    /// <remarks>
    /// Half a point: Excel stores a column width in characters and reports the width in
    /// points, so a round trip through one unit and back is not exact. Half a point is
    /// below what any border or grid line can show, so a column inside it needs no write
    /// and the workbook stays clean.
    /// </remarks>
    private const double _gutterWidthTolerancePt = 0.5;

    /// <summary>
    /// Makes column A the chart's left margin by setting its width to the framing rows'
    /// padding height (ADR-0059, R5.18k).
    /// </summary>
    /// <param name="table">The Gantt data table, whose worksheet owns the gutter column.</param>
    /// <remarks>
    /// <para>
    /// <b>Why one provisional write and one correction rather than a formula.</b> Excel
    /// expresses a column width in CHARACTERS of the Normal font and reports it in
    /// POINTS, and the conversion depends on the workbook's font. The existing width and
    /// its measured points give that ratio for THIS workbook, so one correction lands on
    /// the target without guessing a characters-per-point constant. There is no second
    /// correction: a value still outside the tolerance afterwards is left alone rather
    /// than iterated, because the scene measures whatever is really there.
    /// </para>
    /// <para>
    /// <b>Zero writes on a correct sheet.</b> A column already at the target is not
    /// rewritten, for the same reason the date format and the hidden flags are not: a
    /// no-op pass must not mark the workbook dirty.
    /// </para>
    /// </remarks>
    private void ApplyLeadingGutterWidth(ListObject table)
    {
        if (table.Parent is not Worksheet worksheet
            || GetColumnRange(worksheet) is not { } column)
        {
            return;
        }

        if (ReadColumnWidthCharacters(column) is not { } characters
            || characters <= 0
            || ReadColumnWidthPoints(column) is not { } points
            || points <= 0)
        {
            return;
        }

        double targetPt = GanttCatalogues.MetricDefault("ChartPaddingRowHeightPt");
        if (CorrectedColumnWidthCharacters(characters, points, targetPt) is { } corrected)
        {
            SetColumnWidthCharacters(column, corrected);
        }
    }

    /// <summary>
    /// Converts a column's current character width and measured point width into the
    /// character width that lands on <paramref name="targetPt"/>.
    /// </summary>
    /// <param name="characters">The width the column carries, in Excel's character unit.</param>
    /// <param name="points">That same width, as the host measured it, in points.</param>
    /// <param name="targetPt">The width the gutter needs, in points.</param>
    /// <returns>
    /// The character width to write, or <see langword="null"/> when nothing should be
    /// written.
    /// </returns>
    /// <remarks>
    /// <para>
    /// <b>Why this is a pure function.</b> The ratio between the two units belongs to the
    /// workbook's Normal font, so it must come from the sheet rather than a constant; the
    /// arithmetic that applies it is the part a test can pin without a host, and it is
    /// the part that would otherwise be wrong silently.
    /// </para>
    /// <para>
    /// A width already inside <see cref="_gutterWidthTolerancePt"/> returns
    /// <see langword="null"/> — the no-write path — and so does any input that is not a
    /// usable pair of positive, finite numbers: a correction computed from an unusable
    /// measurement would write a width nobody measured.
    /// </para>
    /// </remarks>
    internal static double? CorrectedColumnWidthCharacters(double characters, double points, double targetPt)
    {
        // One expression, so the three reasons for "no write" cannot drift apart: an
        // unusable pair, a target that is not a width, and a width already inside the
        // tolerance are all the same answer to the caller.
        bool usable = double.IsFinite(characters) && characters > 0
            && double.IsFinite(points) && points > 0
            && double.IsFinite(targetPt) && targetPt > 0;

        return usable && Math.Abs(points - targetPt) > _gutterWidthTolerancePt
            ? characters * targetPt / points
            : null;
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

        // The host's own table style is the BASE the theme's cell formatting sits on. A
        // hide/show re-applies it, so a neutral built-in style is what keeps the table's
        // lines stable when the user toggles a column (owner ruling 2026-11-14). Read
        // first: a table already carrying the style is not rewritten, so a converged
        // initialise still leaves the workbook clean.
        string? liveStyle = ReadTableStyleName(table);
        if (!string.Equals(liveStyle, GanttTableAppearance.TableStyleName, StringComparison.Ordinal))
        {
            SetTableStyle(table, GanttTableAppearance.TableStyleName);
        }

        // The theme's body font and vertical centring are the SAME for every column, so
        // they are written once over the table's whole columns rather than once per
        // column. Measured 2026-11-14 on the live host: the presentation pass cost
        // 252.3 ms for 17 columns (~15 ms each) and was 55% of Initialise, while the
        // columns' font and vertical alignment were being rewritten 17 times each with
        // an identical value. Two writes replace thirty-four.
        //
        // The whole column is still the target - not the table's body range - so a row
        // that does not exist yet already carries the theme instead of inheriting
        // whatever Excel decides when the user types.
        Excel.Range allColumns = GetEntireColumn(table.Range);
        ApplyBodyFont(allColumns);
        ApplyVerticalCentre(allColumns);

        for (int index = 1; index <= schema.Count; index++)
        {
            ListColumn column = GetColumnAt(columns, index);
            Excel.Range columnRange = column.Range;

            // The comparison looks at the BODY when the table has one, because the header
            // cell is locked separately after the loop (see below). Reading the whole list
            // column would report a MIXED range -- a locked header over an unlocked
            // authoring body -- as null, which the adapter treats as drift, so the write
            // would rerun on every Initialise and dirty the workbook for a sheet that is
            // already in the intended state.
            bool wantedLocked = schema[index - 1].IsLocked;
            Excel.Range? columnBody = column.DataBodyRange;
            bool? liveLocked = columnBody is null ? null : ReadLockedFlag(columnBody);
            if (liveLocked is not { } current || current != wantedLocked)
            {
                // The whole column, so a row that does not exist yet already carries the
                // class's own lock flag. The header cell is re-locked once below.
                columnRange.Locked = wantedLocked;
            }

            Excel.Range entireColumn = GetEntireColumn(columnRange);

            // `Range.Hidden` is declared `object` in this PIA (verified by reflection
            // over the installed Microsoft.Office.Interop.Excel 14.0.1), so comparing it
            // to a `bool` with `!=` boxed the right-hand side and compared references:
            // always true, whatever the host actually reported. Every column was
            // therefore rewritten on every Initialise, which marks the workbook dirty
            // for a sheet that is already in the intended state. The value is unwrapped
            // to a real `bool` first, so a column already carrying the schema's
            // visibility is left alone.
            bool wanted = schema[index - 1].IsHidden;
            if (ReadHiddenFlag(entireColumn) != wanted)
            {
                entireColumn.Hidden = wanted;
            }

            // A date column left on General shows its serial number, so the format is
            // part of the presentation contract, not an optional nicety.
            if (GanttTableSchema.IsDateColumn(schema[index - 1].Name))
            {
                ApplyDateColumnFormat(entireColumn);
            }

            // Only this column's alignment varies (R5.18d); the font and the vertical
            // centring are written once for the whole table above.
            ApplyHorizontalAlignment(entireColumn, GanttTableAppearance.AlignmentFor(schema[index - 1].Name));
        }

        // The header row is locked for EVERY column (R5.18k follow-up). `ListColumn.Range`
        // includes the header cell, so the per-column write above left the headers of the
        // authoring columns unlocked -- and once the visible columns led the table, that
        // put an unlocked `Type` header directly above a protected body. A user could then
        // rename a schema column, which breaks the name-keyed column map every adapter
        // builds. Written only when the host does not already report it locked, so a
        // refresh does not dirty the workbook.
        Excel.Range? headerRow = table.HeaderRowRange;
        if (headerRow is not null && headerRow.Locked is not true)
        {
            headerRow.Locked = true;
        }

        // Table-wide, so they run once rather than per column.
        ApplyHeaderFormat(table);
        ApplyBorderColour(table);
    }

    /// <summary>Applies the theme's header fill, text colour and header type to the table's header row.</summary>
    /// <param name="table">The visible data table.</param>
    /// <remarks>
    /// <para>
    /// The header's fill, type and text colour are the chart theme's; Excel's table style
    /// keeps every border it draws inside the table (owner ruling 2026-11-13). Both proxies
    /// are held locally rather than chained.
    /// </para>
    /// <para>
    /// <b>The text colour is not optional.</b> Live evidence 2026-11-14: writing the white
    /// <see cref="GanttTableAppearance.HeaderFill"/> without it left the table style's white
    /// header font in place, so every schema name was present and invisible. The fill and
    /// the text colour are written together, from the same authority.
    /// </para>
    /// </remarks>
    internal virtual void ApplyHeaderFormat(ListObject table)
    {
        Excel.Range header = table.HeaderRowRange;
        Font font = header.Font;
        ApplyHeaderType(font);
        font.Color = GanttTableAppearance.HeaderTextColour.ARGB;
        Interior interior = header.Interior;
        interior.Color = GanttTableAppearance.HeaderFill.ARGB;
    }

    /// <summary>Writes the theme's header type onto a header row's font.</summary>
    /// <param name="font">The header row's font.</param>
    /// <remarks>
    /// The bold flag is written rather than left alone. Excel's table style bolds a header
    /// row, so a regular header only exists if the adapter says so — and the ruling
    /// (2026-11-14) is that the header matches the chart's period band, which is regular.
    /// One seam, so the initialiser and the recovery cannot disagree about its type.
    /// </remarks>
    internal virtual void ApplyHeaderType(Font font)
    {
        font.Name = GanttTableAppearance.FontFamily;
        font.Size = GanttTableAppearance.HeaderFontSizePt;
        font.Bold = GanttTableAppearance.HeaderIsBold;
    }

    /// <summary>
    /// Recolours every border of the table and gives its outside edge the frame's weight.
    /// </summary>
    /// <param name="table">The visible data table.</param>
    /// <remarks>
    /// <c>Borders.Color</c> writes the colour of every border the range contains and leaves
    /// each border's line style and weight alone, which is the split the 2026-11-13 ruling
    /// describes: the table style owns the cell grid, the theme owns the colour. The
    /// 2026-11-14 ruling adds one exception — the table's <em>outside</em> edge takes the
    /// chart frame's weight, so the two boundaries meet at the same 1pt instead of one
    /// stepping down to the other. Only the four edge indices are written; the interior
    /// lines keep whatever weight the table style gives them.
    /// </remarks>
    internal virtual void ApplyBorderColour(ListObject table)
    {
        Excel.Range range = table.Range;
        Borders borders = range.Borders;
        borders.Color = GanttTableAppearance.BorderColour.ARGB;
        foreach (XlBordersIndex edge in _outsideBorderEdges)
        {
            // Excel's cell-border weights are a closed set, not a point width: xlThin is
            // the 1pt weight, which is the frame token's default (MajorBoundaryPt).
            SetBorderWeight(borders, edge, XlBorderWeight.xlThin);
        }
    }

    /// <summary>Writes one border's weight.</summary>
    /// <param name="borders">The table range's border collection.</param>
    /// <param name="edge">Which border to write.</param>
    /// <param name="weight">The weight to write.</param>
    /// <remarks>
    /// <c>Borders.Item</c> is an indexed COM property, so it is a documented seam like
    /// <see cref="GetTableAt"/>: expression trees cannot contain it (CS0855) and the
    /// contract tests substitute this member. The one proxy is held in a local rather than
    /// chained, per the COM ownership rule.
    /// </remarks>
    internal virtual void SetBorderWeight(Borders borders, XlBordersIndex edge, XlBorderWeight weight)
    {
        Border border = borders[edge];
        border.Weight = weight;
    }

    /// <summary>Reads the table's own Excel table-style name.</summary>
    /// <param name="table">The visible data table.</param>
    /// <returns>The style's name, or <see langword="null"/> when the host reports none.</returns>
    /// <remarks>
    /// <c>ListObject.TableStyle</c> is <c>object</c>-typed and reports either a
    /// <c>TableStyle</c> proxy or the name it was assigned by. The proxy's <c>Name</c> is
    /// read through a local rather than chained, per the COM ownership rule; the
    /// comparison itself lives at the call site so this stays a seam.
    /// </remarks>
    internal virtual string? ReadTableStyleName(ListObject table)
    {
        object style = table.TableStyle;
        return style switch
        {
            TableStyle named => named.Name,
            string text => text,
            _ => null,
        };
    }

    /// <summary>Assigns a built-in table style by name.</summary>
    /// <param name="table">The visible data table.</param>
    /// <param name="styleName">The built-in style's name.</param>
    internal virtual void SetTableStyle(ListObject table, string styleName) =>
        table.TableStyle = styleName;

    /// <summary>Applies the theme's body font family and size to a whole worksheet column.</summary>
    /// <param name="entireColumn">The worksheet column behind the table column.</param>
    /// <remarks>
    /// One proxy is held locally rather than chained through <c>range.Font.Name</c>: the COM
    /// rule in ADR-0011 is that no proxy expression is chained, and a chained set here would
    /// also leak an un-released <c>Font</c> proxy on every column of every initialise.
    /// </remarks>
    internal virtual void ApplyBodyFont(Excel.Range entireColumn)
    {
        Font font = entireColumn.Font;
        font.Name = GanttTableAppearance.FontFamily;
        font.Size = GanttTableAppearance.BodyFontSizePt;
    }

    /// <summary>Middle-aligns a whole worksheet column.</summary>
    /// <param name="entireColumn">The worksheet column behind the table column.</param>
    /// <remarks>
    /// A table row is one worksheet row, so a top-aligned cell sits against the row above as
    /// soon as that row grows — which is exactly what the row-height normaliser does on a
    /// group expand. Middle is therefore the only alignment that stays correct.
    /// </remarks>
    internal virtual void ApplyVerticalCentre(Excel.Range entireColumn) =>
        entireColumn.VerticalAlignment = XlVAlign.xlVAlignCenter;

    /// <summary>Applies the column's horizontal alignment to a whole worksheet column.</summary>
    /// <param name="entireColumn">The worksheet column behind the table column.</param>
    /// <param name="alignment">The column's alignment, from the Core appearance authority.</param>
    internal virtual void ApplyHorizontalAlignment(Excel.Range entireColumn, GanttCellAlignment alignment) =>
        entireColumn.HorizontalAlignment = alignment == GanttCellAlignment.Centre
            ? XlHAlign.xlHAlignCenter
            : XlHAlign.xlHAlignLeft;

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
    /// Reads <c>Range.Locked</c>.
    /// </summary>
    /// <param name="range">The range to read.</param>
    /// <returns>
    /// The flag, or <see langword="null"/> when the host reports a mixed range (a locked
    /// header over an unlocked body) or a value that is not a boolean.
    /// </returns>
    /// <remarks>
    /// <c>Locked</c> is declared <see cref="object"/> in this PIA too, so it is unboxed
    /// before comparison for the same reason <c>Hidden</c> is. <see langword="null"/>
    /// deliberately differs from BOTH desired states: a mixed range is drift, and treating
    /// it as "already correct" would report success over a range the adapter never fixed.
    /// </remarks>
    private static bool? ReadLockedFlag(Excel.Range range) => range.Locked switch
    {
        bool flag => flag,
        null => null,
        _ => bool.TryParse(
            Convert.ToString(range.Locked, CultureInfo.InvariantCulture),
            out bool parsed)
            ? parsed
            : null,
    };

    /// <summary>
    /// Writes the date format onto a date column's whole worksheet column, unless it
    /// already carries it.
    /// </summary>
    /// <param name="entireColumn">The worksheet column behind the table column.</param>
    /// <remarks>
    /// <para>
    /// The WHOLE column, not the table's body. A date typed into an empty cell, and a
    /// date written into a row that does not exist yet, both need the format already
    /// in place; formatting only the body leaves every future row on General, which is
    /// exactly how a date reached the sheet as its serial number.
    /// </para>
    /// <para>
    /// The read-before-write is the same discipline the hidden flag follows: rewriting
    /// an unchanged format marks the workbook dirty for a sheet that is already in the
    /// intended state. <c>NumberFormat</c> is <see cref="object"/>-typed in this PIA, so
    /// a value that is not a string reads as "not the wanted format" and is written.
    /// </para>
    /// <para>
    /// The pattern comes from the single Core authority
    /// (<see cref="GanttDateFormatting.WorksheetPatternFor"/>) and is
    /// <see cref="GanttDateDisplayFormat.DdMMMyy"/>, <c>dd-mmm-yy</c>, which is what the
    /// column is styled with (owner ruling 2026-11-14). It was <c>dd/mm/yyyy</c>: the
    /// adapter hardcoded the other approved member, so the columns were styled with a
    /// pattern nobody had chosen. A workbook whose stored setting names another member is
    /// the settings-driven presentation pass, not a second spelling of the pattern here.
    /// </para>
    /// </remarks>
    private void ApplyDateColumnFormat(Excel.Range entireColumn)
    {
        string pattern = GanttDateFormatting.WorksheetPatternFor(GanttDateDisplayFormat.DdMMMyy);
        if (string.Equals(ReadNumberFormat(entireColumn), pattern, StringComparison.Ordinal))
        {
            return;
        }

        SetNumberFormat(entireColumn, pattern);
    }

    /// <summary>
    /// Creates the configuration worksheet directly after the Gantt sheet,
    /// names it from the Core contract, and sets <c>xlSheetVeryHidden</c>.
    /// </summary>
    /// <param name="sheets">The workbook's sheets.</param>
    /// <param name="target">The Gantt worksheet the config sheet follows.</param>
    private static void CreateConfigurationSheet(Sheets sheets, Worksheet target)
    {
        Worksheet config = (Worksheet)sheets.Add(After: target);
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
        => target.Cells[GanttSheetLayout.HeaderRowIndex, GanttSheetLayout.TableOriginColumnIndex]
            .Resize[1, columnCount];

    /// <summary>Returns the leading gutter column. Test seam over the <c>Columns</c> indexer.</summary>
    /// <param name="worksheet">The Gantt worksheet.</param>
    /// <returns>The whole-column range, or <see langword="null"/> when the host reports none.</returns>
    internal virtual Excel.Range? GetColumnRange(Worksheet worksheet)
    {
        ArgumentNullException.ThrowIfNull(worksheet);

        return worksheet.Columns[GanttSheetLayout.LeadingGutterColumnIndex];
    }

    /// <summary>Reads a column's width in Excel's character unit. Test seam over <c>ColumnWidth</c>.</summary>
    /// <param name="column">The whole-column range.</param>
    /// <returns>The width in characters, or <see langword="null"/> when the host reports none.</returns>
    internal virtual double? ReadColumnWidthCharacters(Excel.Range column)
    {
        ArgumentNullException.ThrowIfNull(column);

        return column.ColumnWidth is double width ? width : null;
    }

    /// <summary>Reads a column's measured width in points. Test seam over <c>Width</c>.</summary>
    /// <param name="column">The whole-column range.</param>
    /// <returns>The width in points, or <see langword="null"/> when the host reports none.</returns>
    internal virtual double? ReadColumnWidthPoints(Excel.Range column)
    {
        ArgumentNullException.ThrowIfNull(column);

        return column.Width is double width ? width : null;
    }

    /// <summary>Writes a column's width in Excel's character unit. Test seam over <c>ColumnWidth</c>.</summary>
    /// <param name="column">The whole-column range.</param>
    /// <param name="width">The width in characters.</param>
    internal virtual void SetColumnWidthCharacters(Excel.Range column, double width)
    {
        ArgumentNullException.ThrowIfNull(column);

        column.ColumnWidth = width;
    }

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

    /// <summary>
    /// Reads <c>Range.NumberFormat</c>, which is <see cref="object"/>-typed in this PIA.
    /// </summary>
    /// <param name="range">The range to read.</param>
    /// <returns>The format code, or <see langword="null"/> when the host did not report one.</returns>
    /// <remarks>
    /// A range whose cells carry different formats reports no single code, so null is
    /// the honest answer and the caller writes rather than assuming the format matches.
    /// </remarks>
    internal virtual string? ReadNumberFormat(Excel.Range range) =>
        range.NumberFormat is string format ? format : null;

    /// <summary>Writes a format code onto a range.</summary>
    /// <param name="range">The range to format.</param>
    /// <param name="pattern">The host format code.</param>
    internal virtual void SetNumberFormat(Excel.Range range, string pattern) =>
        range.NumberFormat = pattern;
}
