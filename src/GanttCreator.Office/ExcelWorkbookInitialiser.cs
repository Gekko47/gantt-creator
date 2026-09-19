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
/// Mutation order (work item R2.2 decision D4): all read-only checks first,
/// then rename, header row, table, configuration sheet, defined name. A
/// failure before the rename mutates nothing; a failure after it leaves at
/// most a renamed blank worksheet (cosmetic) and propagates to the command
/// boundary, which translates it to one safe message.
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
public class ExcelWorkbookInitialiser(object? application) : IWorkbookInitialiser
{
    private readonly Application? _application = application as Application;

    /// <inheritdoc />
    public WorkbookInitialiseOutcome Initialise()
    {
        Application? application = _application;
        if (application is null)
        {
            return WorkbookInitialiseOutcome.Refused(InitialiseRefusalReason.NoActiveWorkbook);
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
        // adopted only when it is blank; a chart sheet or any non-empty
        // worksheet causes a fresh sheet to be created. A chart sheet is not
        // an Excel.Worksheet, so the cast selects the create path for it.
        var activeWorksheet = workbook.ActiveSheet as Worksheet;
        var adopt = false;
        if (activeWorksheet is not null && IsBlank(activeWorksheet, application))
        {
            adopt = true;
        }

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
            WritePlotAnchorName(target);
            wrotePlotAnchor = true;
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
    /// Determines whether the worksheet is blank: no cell in its used range
    /// holds a value.
    /// </summary>
    /// <param name="worksheet">The candidate target worksheet.</param>
    /// <param name="application">The Excel application (for the worksheet function).</param>
    /// <returns><see langword="true"/> when the used range contains no values.</returns>
    private static bool IsBlank(Worksheet worksheet, Application application)
    {
        Excel.Range? usedRange = worksheet.UsedRange;
        if (usedRange is null)
        {
            return true;
        }

        WorksheetFunction functions = application.WorksheetFunction;
        return functions.CountA(usedRange) == 0;
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
                object sheet = sheets[index];
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
