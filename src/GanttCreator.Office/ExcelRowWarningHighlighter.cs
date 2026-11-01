using GanttCreator.Core;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live <see cref="IRowWarningHighlighter"/> over the Excel application object
/// supplied by the host at add-in load. Fills row or cell interiors with
/// pastel red/pink (<see cref="GanttValidationSeverity.Error"/>) or pastel
/// amber/yellow (<see cref="GanttValidationSeverity.Warning"/>), and clears
/// them with <c>xlColorIndexNone</c>.
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
/// <c>Range</c>, <c>ListObject</c>, and <c>ListColumns</c> objects reached here
/// are Excel-owned shared roots. This adapter takes no ownership of them,
/// never calls <c>FinalReleaseComObject</c>, and force-releases nothing (the
/// ownership policy of <see cref="ExcelApplicationAdapter"/>). Every proxy is
/// held in a local and used without chained member expressions; sheets are
/// reached by index through <see cref="GetSheetAt"/> and tables through
/// <c>GetTableAt</c>.
/// </para>
/// <para>
/// Row scope fills every visible column of the <c>tblGanttData</c> body row
/// (all columns in <see cref="GanttTableSchema.Default"/>), so a blocked
/// row is visually distinct at a glance. Cell scope fills a single
/// user-visible data cell.
/// </para>
/// <para>
/// Colours are pastel (not the vivid standard ColorIndex 3/46) so the fills
/// are visible without being visually aggressive:
/// <list type="table">
/// <listheader><term>Severity</term><description>RGB colour</description></listheader>
/// <item><term><see cref="GanttValidationSeverity.Error"/></term><description>RGB(255, 200, 200) — pastel red/pink</description></item>
/// <item><term><see cref="GanttValidationSeverity.Warning"/></term><description>RGB(255, 235, 170) — pastel amber/yellow</description></item>
/// </list>
/// </para>
/// <para>
/// The instance tracks every address it highlights so that
/// <see cref="ReconcileHighlights"/> can perform only the minimal set of COM
/// writes needed to bring the sheet in line with the current targets.
/// Row fills are tracked per-cell (every column of the row), so a row clear
/// removes every cell that was previously filled for that row.
/// </para>
/// <para>
/// Row numbers are bounds-checked against
/// <see cref="Excel.Range.Rows"/>/<c>Count</c> of the table body, so a row
/// number beyond the table's current extent never writes outside the table
/// bounds (for example when a new row is being added at the end of the table).
/// </para>
/// <para>
/// Row fills are batched: setting <c>Interior.Color</c> on the entire row
/// range in one COM call instead of one call per cell, reducing COM
/// round-trips from 18+ to 1 per row.
/// </para>
/// <para>
/// The <c>internal virtual</c> accessors (<see cref="GetSheetAt"/>,
/// <c>GetTableAt</c> (both <c>ListObjects</c> and <c>ListColumns</c> overloads),
/// <see cref="GetCellRange"/>) isolate the Excel COM parameterised properties
/// (indexers). Expression trees cannot contain indexed properties (CS0855), so
/// contract tests substitute these seams and every other member through Moq;
/// the real indexer behaviour is exercised by the tagged live-Office
/// integration test.
/// </para>
/// </remarks>
/// <param name="protectionGuard">The shared read-only workbook-protection guard; defaults to a live guard over <paramref name="application"/>.</param>
public class ExcelRowWarningHighlighter(
    object? application,
    IWorksheetProtectionGuard? protectionGuard = null) : IRowWarningHighlighter
{
    private readonly Application? _application = application as Application;
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    /// <summary>Pastel red/pink = RGB(255, 200, 200) as OLE_COLOR (0x00C8C8FF).</summary>
    private const int _olePastelRed = 0x00C8C8FF;

    /// <summary>Pastel amber/yellow = RGB(255, 235, 170) as OLE_COLOR (0x00AAEBFF).</summary>
    private const int _olePastelAmber = 0x00AAEBFF;

    /// <summary>Workbook CustomDocumentProperties key under which highlighted addresses are persisted so they survive workbook close/reopen.</summary>
    private const string _highlightsPropertyName = "GanttCreator.RowHighlights";

    /// <summary>
    /// Every address previously highlighted by this instance, as
    /// <c>(Row, Column)</c> where <c>Column</c> is <see langword="null"/>
    /// for a whole-row fill. Used by <see cref="ReconcileHighlights"/> to
    /// compute the minimal diff against the current targets.
    /// </summary>
    private readonly HashSet<(int Row, int? Column)> _previousHighlights = [];

    /// <summary>Set once <see cref="LoadHighlights"/> has run for the current workbook.</summary>
    private bool _highlightsLoaded;

    /// <summary>Fills or clears the entire visible row of <c>tblGanttData</c> body cells.</summary>
    public GanttValidationReportOutcome HighlightRow(int rowNumber, GanttValidationSeverity severity)
    {
        ProtectionGuardOutcome activeProtection = _protectionGuard.Query();
        if (activeProtection != ProtectionGuardOutcome.NotProtected)
        {
            return ProtectionRefusal(activeProtection);
        }

        HighlightSetup? setup = ResolveSetup();
        if (setup is null)
        {
            return GanttValidationReportOutcome.Refused(GanttValidationReportRefusalReason.NoActiveWorkbook);
        }
        if (setup.RefusalReason is { } refusal)
        {
            return GanttValidationReportOutcome.Refused(refusal);
        }

        // Bounds check: never write outside the table body's current extent.
        if (setup.Body is null || rowNumber < 1 || rowNumber > setup.Body.Rows.Count)
        {
            return GanttValidationReportOutcome.Ok(0);
        }

        int color = ResolveOleColor(severity);
        (int Row, int? Column) addressKey = (rowNumber, null);

#pragma warning disable IDE0045 // If/else can't be simplified: Remove vs Add are different methods
#pragma warning disable IDE0011 // Braces required by project style
        if (color == (int)XlColorIndex.xlColorIndexNone)
        {
            _ = _previousHighlights.Remove(addressKey);
        }
        else
        {
            _ = _previousHighlights.Add(addressKey);
        }
#pragma warning restore IDE0011
#pragma warning restore IDE0045

        // Batch write: set Interior.Color on the entire row range in one COM call
        // instead of one call per cell (18+ → 1).
        Excel.Range rowRange = setup.Body.Rows[rowNumber];
        rowRange.Interior.Color = color;

        return GanttValidationReportOutcome.Ok(setup.Body.Columns.Count);
    }

    /// <summary>Fills or clears a single user-visible data cell.</summary>
    public GanttValidationReportOutcome HighlightCell(int rowNumber, int column, GanttValidationSeverity severity)
    {
        ProtectionGuardOutcome activeProtection = _protectionGuard.Query();
        if (activeProtection != ProtectionGuardOutcome.NotProtected)
        {
            return ProtectionRefusal(activeProtection);
        }

        HighlightSetup? setup = ResolveSetup();
        if (setup is null)
        {
            return GanttValidationReportOutcome.Refused(GanttValidationReportRefusalReason.NoActiveWorkbook);
        }
        if (setup.RefusalReason is { } refusal)
        {
            return GanttValidationReportOutcome.Refused(refusal);
        }

        // Bounds check: never write outside the table body's current extent.
        if (setup.Body is null || rowNumber < 1 || rowNumber > setup.Body.Rows.Count ||
            column < 1 || column > setup.Body.Columns.Count)
        {
            return GanttValidationReportOutcome.Ok(0);
        }

        int color = ResolveOleColor(severity);
        (int rowNumber, int column) addressKey = (rowNumber, column);

        _ = color == (int)XlColorIndex.xlColorIndexNone ? _previousHighlights.Remove(addressKey) : _previousHighlights.Add(addressKey);

        Excel.Range? cell = GetCellRange(setup.Body, rowNumber, column);
        if (cell is null)
        {
            return GanttValidationReportOutcome.Ok(0);
        }

        cell.Interior.Color = color;
        return GanttValidationReportOutcome.Ok(1);
    }

    /// <summary>Clears the highlight fill from a row, or from a single cell.</summary>
    public GanttValidationReportOutcome ClearHighlight(int rowNumber, int? column = null)
    {
        ProtectionGuardOutcome activeProtection = _protectionGuard.Query();
        if (activeProtection != ProtectionGuardOutcome.NotProtected)
        {
            return ProtectionRefusal(activeProtection);
        }

        HighlightSetup? setup = ResolveSetup();
        if (setup is null)
        {
            return GanttValidationReportOutcome.Refused(GanttValidationReportRefusalReason.NoActiveWorkbook);
        }
        if (setup.RefusalReason is { } refusal)
        {
            return GanttValidationReportOutcome.Refused(refusal);
        }

        int cleared = 0;
        if (column is not null)
        {
            (int rowNumber, int Value) addressKey = (rowNumber, column.Value);
            _ = _previousHighlights.Remove(addressKey);

            if (setup.Body is not null && rowNumber >= 1 && rowNumber <= setup.Body.Rows.Count &&
                column.Value >= 1 && column.Value <= setup.Body.Columns.Count)
            {
                Excel.Range? cell = GetCellRange(setup.Body, rowNumber, column.Value);
                if (cell is not null)
                {
                    cell.Interior.Color = (int)XlColorIndex.xlColorIndexNone;
                    cleared = 1;
                }
            }
        }
        else
        {
            _ = _previousHighlights.RemoveWhere(addr => addr.Row == rowNumber);

            if (setup.Body is not null && rowNumber >= 1 && rowNumber <= setup.Body.Rows.Count)
            {
                // Batch clear: set Interior.Color on the entire row range in one COM call.
                Excel.Range rowRange = setup.Body.Rows[rowNumber];
                rowRange.Interior.Color = (int)XlColorIndex.xlColorIndexNone;
                cleared = setup.Body.Columns.Count;
            }
        }

        return GanttValidationReportOutcome.Ok(cleared);
    }

    /// <summary>
    /// Reconciles the highlight fills against the supplied targets, performing
    /// only the minimal set of COM writes needed.
    /// </summary>
    public GanttValidationReportOutcome ReconcileHighlights(
        IReadOnlyList<HighlightTarget> currentTargets,
        IReadOnlyList<int>? rowsToReconcile = null)
    {
        ArgumentNullException.ThrowIfNull(currentTargets);
        ProtectionGuardOutcome activeProtection = _protectionGuard.Query();
        if (activeProtection != ProtectionGuardOutcome.NotProtected)
        {
            return ProtectionRefusal(activeProtection);
        }

        HighlightSetup? setup = ResolveSetup();
        if (setup is null)
        {
            return GanttValidationReportOutcome.Refused(GanttValidationReportRefusalReason.NoActiveWorkbook);
        }
        if (setup.RefusalReason is { } refusal)
        {
            return GanttValidationReportOutcome.Refused(refusal);
        }

        if (setup.Body is null)
        {
            return GanttValidationReportOutcome.Ok(0);
        }

        // When a row scope is supplied, build a fast lookup set so we only
        // reconcile highlights within those rows. Entries outside the scope
        // belong to unchanged rows whose validation state was not re-checked
        // on this refresh and must be left intact.
        HashSet<int>? scope = rowsToReconcile is not null
            ? new(rowsToReconcile)
            : null;

        // Build the set of current targets for fast lookup.
        HashSet<(int Row, int? Column)> currentSet = [];
        foreach (HighlightTarget target in currentTargets)
        {
            _ = currentSet.Add((target.Row, target.Column));
        }

        // Bounds-check all targets up front; skip anything outside the table body.
        List<HighlightTarget> validTargets = [];
        foreach (HighlightTarget target in currentTargets)
        {
            if (target.Row < 1 || target.Row > setup.Body.Rows.Count)
            {
                continue;
            }
            if (target.Column is { } col &&
                (col < 1 || col > setup.Body.Columns.Count))
            {
                continue;
            }
            validTargets.Add(target);
        }

        int written = 0;

        // Clear addresses that were highlighted before but are no longer needed.
        // When a row scope is active, only clear entries whose row is in scope;
        // entries for unchanged rows are left intact on the sheet and in tracking.
        foreach ((int row, int? column) in _previousHighlights.ToArray())
        {
            if (scope is not null && !scope.Contains(row))
            {
                continue;
            }
            if (!currentSet.Contains((row, column)))
            {
                if (column is null)
                {
                    // Whole-row clear: batch write on the row range.
                    if (row >= 1 && row <= setup.Body.Rows.Count)
                    {
                        Excel.Range rowRange = setup.Body.Rows[row];
                        rowRange.Interior.Color = (int)XlColorIndex.xlColorIndexNone;
                        _ = _previousHighlights.Remove((row, null));
                        written += setup.Body.Columns.Count;
                    }
                }
                else
                {
                    // Single-cell clear.
                    if (row >= 1 && row <= setup.Body.Rows.Count &&
                        column >= 1 && column <= setup.Body.Columns.Count)
                    {
                        Excel.Range? cell = GetCellRange(setup.Body, row, column.Value);
                        if (cell is not null)
                        {
                            cell.Interior.Color = (int)XlColorIndex.xlColorIndexNone;
                            _ = _previousHighlights.Remove((row, column.Value));
                            written++;
                        }
                    }
                }
            }
        }

        // Fill addresses that are newly needed (not previously highlighted).
        foreach (HighlightTarget target in validTargets)
        {
            if (!_previousHighlights.Contains((target.Row, target.Column)))
            {
                int color = ResolveOleColor(target.Severity);
                if (target.Column is null)
                {
                    // Whole-row fill: batch write on the row range.
                    Excel.Range rowRange = setup.Body.Rows[target.Row];
                    rowRange.Interior.Color = color;
                    _ = _previousHighlights.Add((target.Row, null));
                    written += setup.Body.Columns.Count;
                }
                else
                {
                    // Single-cell fill.
                    Excel.Range? cell = GetCellRange(setup.Body, target.Row, target.Column.Value);
                    if (cell is not null)
                    {
                        cell.Interior.Color = color;
                        _ = _previousHighlights.Add((target.Row, target.Column.Value));
                        written++;
                    }
                }
            }
            // Addresses in both sets are skipped — no COM call needed.
        }

        // Persist the updated tracking set so highlights survive workbook close
        // and can be reconciled (cleared) after the workbook reopens.
#pragma warning disable CA1031
        if (setup.Worksheet is { } ws)
        {
            try
            {
                SaveHighlights((Workbook)ws.Parent);
            }
            catch
            {
                // Best-effort: persistence failure must not block highlighting.
            }
        }
#pragma warning restore CA1031

        return GanttValidationReportOutcome.Ok(written);
    }

    private static int ResolveOleColor(GanttValidationSeverity severity) =>
        severity switch
        {
            GanttValidationSeverity.Error => _olePastelRed,
            GanttValidationSeverity.Warning => _olePastelAmber,
            _ => (int)XlColorIndex.xlColorIndexNone,
        };

    /// <summary>Loads previously persisted highlight addresses from a workbook-scoped defined name.</summary>
    private void LoadHighlights(Workbook workbook)
    {
        if (_highlightsLoaded)
        {
            return;
        }

#pragma warning disable CA1031
        try
        {
            Name? name = null;
            try
            {
                name = workbook.Names.Item(_highlightsPropertyName);
            }
            catch
            {
                // Name does not exist yet — nothing to load.
            }

            if (name is null)
            {
                _highlightsLoaded = true;
                return;
            }

            string? refersTo = name.RefersTo;
            if (string.IsNullOrEmpty(refersTo))
            {
                _highlightsLoaded = true;
                return;
            }

            // RefersTo is of the form ="..." — strip the leading = and surrounding quotes.
            string value = refersTo.Trim();
            if (value.StartsWith('='))
            {
                value = value[1..];
            }
            if (value.Length >= 2 && value.StartsWith('"') && value.EndsWith('"'))
            {
                value = value[1..^1];
            }

            if (string.IsNullOrEmpty(value))
            {
                _highlightsLoaded = true;
                return;
            }

            _previousHighlights.Clear();
            foreach (string entry in value.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                string[] parts = entry.Split('|');
                if (parts.Length >= 2 && int.TryParse(parts[0], out int row))
                {
                    int? column = parts.Length >= 3 && int.TryParse(parts[2], out int col) ? col : null;
                    _ = _previousHighlights.Add((row, column));
                }
            }
        }
        catch
        {
            // Best-effort: persistence failure must not block highlighting.
        }
        finally
        {
            _highlightsLoaded = true;
        }
#pragma warning restore CA1031
    }

    /// <summary>Persists current highlight addresses to a workbook-scoped defined name.</summary>
    private void SaveHighlights(Workbook workbook)
    {
#pragma warning disable CA1031
        try
        {
            string value = string.Join(";", _previousHighlights.Select(h =>
                h.Column is { } col ? $"{h.Row}|{col}" : $"{h.Row}|"));
            string refersTo = "=\"" + value + "\"";

            Name? name = null;
            try
            {
                name = workbook.Names.Item(_highlightsPropertyName);
            }
            catch
            {
                // Name does not exist yet — add it below.
            }

            if (name is not null)
            {
                name.RefersTo = refersTo;
                return;
            }

            _ = workbook.Names.Add(_highlightsPropertyName, refersTo);
        }
        catch
        {
            // Best-effort: persistence failure must not block highlighting.
        }
#pragma warning restore CA1031
    }

    private HighlightSetup? ResolveSetup()
    {
        Application? application = _application;
        if (application is null)
        {
            return new HighlightSetup(null, null, GanttValidationReportRefusalReason.NoActiveWorkbook);
        }

        Workbook? workbook = application.ActiveWorkbook;
        if (workbook is null)
        {
            return new HighlightSetup(null, null, GanttValidationReportRefusalReason.NoActiveWorkbook);
        }

        Sheets sheets = workbook.Sheets;
        if (!TryFindTable(sheets, out Worksheet? worksheet, out ListObject? table) ||
            worksheet is null || table is null)
        {
            return new HighlightSetup(null, null, GanttValidationReportRefusalReason.TableMissing);
        }

        // Load previously persisted highlight addresses so corrections made
        // after a close/reopen are still reconciled against the tracking set.
        LoadHighlights(workbook);

        ProtectionGuardOutcome protection = _protectionGuard.QueryTarget(worksheet);
        return protection != ProtectionGuardOutcome.NotProtected
            ? new HighlightSetup(
                null, null,
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? GanttValidationReportRefusalReason.NoActiveWorkbook
                    : GanttValidationReportRefusalReason.TargetProtected)
            : new HighlightSetup(worksheet, table.DataBodyRange, null);
    }

    private static GanttValidationReportOutcome ProtectionRefusal(ProtectionGuardOutcome outcome) =>
        GanttValidationReportOutcome.Refused(
            outcome == ProtectionGuardOutcome.NoActiveWorkbook
                ? GanttValidationReportRefusalReason.NoActiveWorkbook
                : GanttValidationReportRefusalReason.TargetProtected);

    private static bool TryFindTable(
        Sheets sheets,
        out Worksheet? worksheet,
        out ListObject? table)
    {
        worksheet = null;
        table = null;

        try
        {
            foreach (object? sheet in sheets)
            {
                if (sheet is not Worksheet ws)
                {
                    continue;
                }

                foreach (ListObject listObject in ws.ListObjects)
                {
                    if (string.Equals(listObject.Name, "tblGanttData",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        worksheet = ws;
                        table = listObject;
                        return true;
                    }
                }
            }
        }
#pragma warning disable CA1031
        catch
        {
        }
#pragma warning restore CA1031

        return false;
    }

    private sealed class HighlightSetup(Worksheet? worksheet, Excel.Range? body, GanttValidationReportRefusalReason? refusalReason)
    {
        public Worksheet? Worksheet { get; } = worksheet;
        public Excel.Range? Body { get; } = body;
        public GanttValidationReportRefusalReason? RefusalReason { get; } = refusalReason;
    }

    // The three accessors below are the COM-parameterised-property seams. They are
    // kept as `internal virtual` (not protected) to match ExcelGanttValidationReporter
    // and so Moq can derive a testable subclass without naming Excel types.

    /// <summary>Returns the sheet at the one-based index. Test seam over the COM parameterised <c>Sheets.Item</c> property.</summary>
    internal virtual object GetSheetAt(Sheets sheets, int index) => sheets[index];

    /// <summary>Returns the list object at the one-based index. Test seam over the COM parameterised <c>ListObjects.Item</c> property.</summary>
    internal virtual ListObject GetTableAt(ListObjects listObjects, int index) => listObjects[index];

    /// <summary>Returns the list column at the one-based index. Test seam over the COM parameterised <c>ListColumns.Item</c> property.</summary>
    internal virtual ListColumn GetTableAt(ListColumns columns, int index) => columns[index];

    /// <summary>Returns the cell at the one-based body row and one-based table-column index. Test seam over the COM <c>Range.Cells[row, column]</c> indexer.</summary>
    internal virtual Excel.Range? GetCellRange(Excel.Range? body, int rowNumber, int columnIndex)
    {
        if (body is null)
        {
            return null;
        }

        try
        {
            return body.Cells[rowNumber, columnIndex];
        }
#pragma warning disable CA1031
        catch
        {
            return null;
        }
#pragma warning restore CA1031
    }
}
