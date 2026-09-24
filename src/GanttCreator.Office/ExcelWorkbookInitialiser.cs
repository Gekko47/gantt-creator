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

        var wroteHeader = false;
        var createdTable = false;
        var createdConfigSheet = false;
        var wrotePlotAnchor = false;

        try
        {
            WriteHeaderRow(target);
            wroteHeader = true;
            CreateDataTable(target);
            createdTable = true;
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
                    createdTable,
                    createdConfigSheet);
                return catalogueOutcome.Refusal == ConfigWriteRefusalReason.NoActiveWorkbook
                    ? WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.NoActiveWorkbook)
                    : WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.TargetProtected);
            }

            WritePlotAnchorName(target);
            wrotePlotAnchor = true;
            TypeOptionsMaterialiseOutcome typeOptionsOutcome = _typeOptionsMaterialiser.Materialise();
            if (!typeOptionsOutcome.Succeeded)
            {
                RollBackPlotAnchorName(target);
                RollBackConfigurationSheet(sheets);
                RollBackDataTable(target);
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
                RollBackDataTable(target);
            }

            if (wroteHeader)
            {
                RollBackHeaderRow(target);
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
    /// <param name="createdTable">Whether the data table was created.</param>
    /// <param name="createdConfigSheet">Whether the configuration sheet was created.</param>
    private void RollBackForCatalogueRefusal(
        Worksheet target,
        Sheets sheets,
        ConfigWriteOutcome catalogueOutcome,
        bool wroteHeader,
        bool createdTable,
        bool createdConfigSheet)
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
            RollBackDataTable(target);
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
        finally
        {
            _application.DisplayAlerts = original;
        }
    }

    /// <summary>
    /// Deletes the table named <c>GanttTableSchema.TableName</c> on
    /// <paramref name="target"/>, if it exists. Used to roll back
    /// <see cref="CreateDataTable"/> when a later mutation fails.
    /// </summary>
    /// <param name="target">The Gantt worksheet.</param>
    private void RollBackDataTable(Worksheet target)
    {
        if (_application is null)
        {
            return;
        }
        var original = _application.DisplayAlerts;
        try
        {
            _application.DisplayAlerts = false;
            var tableCount = target.ListObjects.Count;
            for (var index = 1; index <= tableCount; index++)
            {
                ListObject table = GetTableAt(target.ListObjects, index);
                if (string.Equals(
                    table.Name,
                    GanttTableSchema.TableName,
                    StringComparison.OrdinalIgnoreCase))
                {
                    table.Delete();
                    break;
                }
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
    /// <paramref name="target"/>. Used to roll back the header row when a later
    /// mutation fails.
    /// </summary>
    /// <param name="target">The Gantt worksheet.</param>
    private void RollBackHeaderRow(Worksheet target)
    {
        Excel.Range headerRange = GetHeaderRange(
            target,
            GanttTableSchema.Default.Columns.Count);
        headerRange.ClearContents();
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
    /// Creates the <c>tblGanttData</c> table over the header row with
    /// header-name behaviour, then names it from the Core schema constant.
    /// The R2.7 appearance settings (no autofilter dropdowns, no banded
    /// rows) keep the data panel visually neutral for the live Gantt
    /// (ADR-0007 D8).
    /// </summary>
    /// <param name="target">The Gantt worksheet.</param>
    private void CreateDataTable(Worksheet target)
    {
        var columnCount = GanttTableSchema.Default.Columns.Count;
        Excel.Range tableRange = GetHeaderRange(target, columnCount);
        ListObjects listObjects = target.ListObjects;
        ListObject table = listObjects.Add(
            XlListObjectSourceType.xlSrcRange,
            tableRange,
            Type.Missing,
            XlYesNoGuess.xlYes,
            Type.Missing);
        table.Name = GanttTableSchema.TableName;
        table.ShowAutoFilter = false;
        table.ShowTableStyleRowStripes = false;
        table.ShowTableStyleColumnStripes = false;
    }

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
    /// Writes the sheet-scoped plot-anchor defined name: the cell one column
    /// right of the table's last column, on the header row.
    /// </summary>
    /// <param name="target">The Gantt worksheet.</param>
    private static void WritePlotAnchorName(Worksheet target)
    {
        var anchorColumnIndex = GanttTableSchema.Default.Columns.Count + 1;
        Names names = target.Names;
        _ = names.Add(
            GanttWorkbookContract.PlotAnchorDefinedName,
            BuildAnchorRefersTo(target.Name, anchorColumnIndex));
    }

    /// <summary>
    /// Builds the <c>refersTo</c> string for the plot anchor:
    /// <c>='&lt;escaped sheet name&gt;'!$&lt;column&gt;$1</c> in invariant
    /// culture. Sheet names may contain apostrophes, so each is doubled
    /// inside the quoted reference.
    /// </summary>
    /// <param name="sheetName">The final label of the Gantt worksheet.</param>
    /// <param name="anchorColumnIndex">The one-based anchor column index.</param>
    /// <returns>The <c>refersTo</c> string.</returns>
    private static string BuildAnchorRefersTo(string sheetName, int anchorColumnIndex)
    {
        var escaped = sheetName.Replace("'", "''", StringComparison.Ordinal);
        return string.Create(
            CultureInfo.InvariantCulture,
            $"='{escaped}'!${ToA1Column(anchorColumnIndex)}$1");
    }

    /// <summary>
    /// Converts a one-based column index to its A1-style letter sequence
    /// (1 → A, 26 → Z, 27 → AA), culture-invariant.
    /// </summary>
    /// <param name="columnIndex">The one-based column index.</param>
    /// <returns>The A1-style column letters.</returns>
    private static string ToA1Column(int columnIndex)
    {
        System.Text.StringBuilder builder = new();
        var remaining = columnIndex;
        while (remaining > 0)
        {
            var digit = (remaining - 1) % 26;
            _ = builder.Insert(0, (char)('A' + digit));
            remaining = (remaining - 1) / 26;
        }

        return builder.ToString();
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
        => target.Cells[1, 1].Resize[1, columnCount];
}
