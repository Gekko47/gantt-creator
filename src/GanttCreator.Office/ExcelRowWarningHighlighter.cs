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
/// <item><term><see cref="GanttValidationSeverity.Warning"/></term><description><see cref="GanttCatalogues.WarningFill"/> (#FFF2CC) — the code-owned pastel warning fill</description></item>
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

    /// <summary>
    /// The code-owned warning fill (<see cref="GanttCatalogues.WarningFill"/>), converted
    /// once to OLE_COLOR. The error fill stays the pastel red constant above.
    /// </summary>
    private static readonly int _oleWarningFill = ToOleColor(GanttCatalogues.WarningFill);

    /// <summary>Workbook-scoped defined name under which highlighted addresses are persisted so they survive workbook close/reopen.</summary>
    private const string _highlightsPropertyName = "GanttCreator.RowHighlights";

    /// <summary>
    /// Every address this instance has filled, as <c>(Row, Column)</c> — <c>Column</c> is
    /// <see langword="null"/> for a whole-row fill — with the severity it was painted at.
    /// Used by <see cref="ReconcileHighlights"/> to compute the minimal diff.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>The severity is part of the record (measured 2026-11-16).</b> This used to be an
    /// address-only set, so an address present in both the previous and the current target
    /// list read as "unchanged" and was skipped. A cell whose finding changed from
    /// <see cref="GanttValidationSeverity.Warning"/> to
    /// <see cref="GanttValidationSeverity.Error"/> therefore kept its amber fill for the
    /// life of the workbook: the one thing the reconcile maintains is the colour, and the
    /// colour was the one thing it never compared.
    /// </para>
    /// <para>
    /// <see langword="null"/> means "recorded by a build that did not store severities"
    /// (a workbook reopened after an upgrade). The address is known and its colour is not,
    /// so the next reconcile re-asserts that one cell once instead of trusting it.
    /// </para>
    /// <para>
    /// <b>The workbook-scoped defined name is the durable truth; this dictionary is one
    /// pass's cache.</b> The live composition builds a new highlighter for every refresh
    /// (<c>RefreshSheetCommand.CreateLiveOrchestrator</c>), so an instance cannot carry
    /// another workbook's addresses. <see cref="LoadHighlights"/> seeds the cache at the
    /// start of the pass and <see cref="ReconcileHighlights"/> writes it back.
    /// </para>
    /// </remarks>
    private readonly Dictionary<(int Row, int? Column), GanttValidationSeverity?> _previousHighlights = [];

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

#pragma warning disable IDE0045 // If/else can't be simplified: Remove and assignment are different operations
#pragma warning disable IDE0011 // Braces required by project style
        if (color == (int)XlColorIndex.xlColorIndexNone)
        {
            _ = _previousHighlights.Remove(addressKey);
        }
        else
        {
            _previousHighlights[addressKey] = severity;
        }
#pragma warning restore IDE0011
#pragma warning restore IDE0045

        // Batch write: set Interior.Color on the entire row range in one COM call
        // instead of one call per cell (18+ → 1).
        Excel.Range rowRange = GetRowRange(setup.Body, rowNumber);
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

        if (color == (int)XlColorIndex.xlColorIndexNone)
        {
            _ = _previousHighlights.Remove(addressKey);
        }
        else
        {
            _previousHighlights[addressKey] = severity;
        }

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
            foreach ((int Row, int? Column) address in _previousHighlights.Keys
                         .Where(entry => entry.Row == rowNumber)
                         .ToArray())
            {
                _ = _previousHighlights.Remove(address);
            }

            if (setup.Body is not null && rowNumber >= 1 && rowNumber <= setup.Body.Rows.Count)
            {
                // Batch clear: set Interior.Color on the entire row range in one COM call.
                Excel.Range rowRange = GetRowRange(setup.Body, rowNumber);
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

        // The desired state, one entry per address. Bounds are checked up front and a
        // severity with no fill colour is dropped rather than written as "clear".
        //
        // WORST SEVERITY WINS. Two issues can name the same cell - a Finish cell on a
        // Start-only type is both "not used by this Type" (Warning) and, if it holds
        // something unreadable, "unsupported value" (Error) - and the previous code
        // painted whichever target the list happened to yield first. Colour by argument
        // order is not a rule, so Error outranks Warning explicitly.
        Dictionary<(int Row, int? Column), GanttValidationSeverity> desired = [];
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
            if (ResolveOleColor(target.Severity) == (int)XlColorIndex.xlColorIndexNone)
            {
                continue;
            }

            (int Row, int? Column) key = (target.Row, target.Column);
            if (!desired.TryGetValue(key, out GanttValidationSeverity existing) ||
                SeverityRank(target.Severity) > SeverityRank(existing))
            {
                desired[key] = target.Severity;
            }
        }

        int written = 0;

        // 1. Clear addresses that were highlighted before but are no longer wanted. When a
        // row scope is active, only entries whose row is in scope are touched; entries for
        // unchanged rows stay on the sheet and in tracking.
        foreach (KeyValuePair<(int Row, int? Column), GanttValidationSeverity?> entry
                     in _previousHighlights.ToArray())
        {
            (int Row, int? Column) address = entry.Key;
            if (scope is not null && !scope.Contains(address.Row))
            {
                continue;
            }
            if (desired.ContainsKey(address))
            {
                continue;
            }

            if (address.Column is null)
            {
                // Whole-row clear: batch write on the row range.
                if (address.Row >= 1 && address.Row <= setup.Body.Rows.Count)
                {
                    Excel.Range rowRange = GetRowRange(setup.Body, address.Row);
                    rowRange.Interior.Color = (int)XlColorIndex.xlColorIndexNone;
                    _ = _previousHighlights.Remove(address);
                    written += setup.Body.Columns.Count;
                }
            }
            else if (address.Column >= 1 && address.Column <= setup.Body.Columns.Count)
            {
                // Single-cell clear.
                Excel.Range? cell = GetCellRange(setup.Body, address.Row, address.Column.Value);
                if (cell is not null)
                {
                    cell.Interior.Color = (int)XlColorIndex.xlColorIndexNone;
                    _ = _previousHighlights.Remove(address);
                    written++;
                }
            }
        }

        // 2. Whole-row targets: painted once, and repainted when the recorded severity
        // changed (or is unknown, which is what a reopened workbook reports).
        HashSet<int> rowsCoveredWholly = [];
        foreach (KeyValuePair<(int Row, int? Column), GanttValidationSeverity> entry in desired)
        {
            if (entry.Key.Column is not null)
            {
                continue;
            }
            if (scope is not null && !scope.Contains(entry.Key.Row))
            {
                continue;
            }
            if (_previousHighlights.TryGetValue(entry.Key, out GanttValidationSeverity? recorded) &&
                recorded == entry.Value)
            {
                continue;
            }

            Excel.Range rowRange = GetRowRange(setup.Body, entry.Key.Row);
            rowRange.Interior.Color = ResolveOleColor(entry.Value);
            _previousHighlights[entry.Key] = entry.Value;
            _ = rowsCoveredWholly.Add(entry.Key.Row);
            written += setup.Body.Columns.Count;
        }

        // 3. Cell targets. A row whose recorded fills already match this finding set costs
        // nothing; a row whose findings changed is ADOPTED: one row clear, then its cells.
        //
        // Why the row clear matters (measured 2026-11-16). A new table row inherits the
        // previous row's cell formats - including the pastel diagnostic fill - so a
        // milestone added under an invalid activity is born pink on Start and Finish. A
        // fill this add-in never wrote is absent from the tracker, so the old diff could
        // never clear it: the milestone's Finish stayed pink for the life of the workbook
        // even though a milestone does not use Finish. Clearing the row makes the SHEET,
        // not the tracker, the authority for a row whose findings changed.
        foreach (IGrouping<int, KeyValuePair<(int Row, int? Column), GanttValidationSeverity>> row in desired
                     .Where(entry => entry.Key.Column is not null)
                     .GroupBy(entry => entry.Key.Row))
        {
            int rowNumber = row.Key;
            if (scope is not null && !scope.Contains(rowNumber))
            {
                continue;
            }
            if (rowsCoveredWholly.Contains(rowNumber))
            {
                continue;
            }

            bool converged = true;
            foreach (KeyValuePair<(int Row, int? Column), GanttValidationSeverity> entry in row)
            {
                if (!_previousHighlights.TryGetValue(entry.Key, out GanttValidationSeverity? recorded) ||
                    recorded != entry.Value)
                {
                    converged = false;
                    break;
                }
            }
            if (converged)
            {
                continue;
            }

            Excel.Range rowRange = GetRowRange(setup.Body, rowNumber);
            rowRange.Interior.Color = (int)XlColorIndex.xlColorIndexNone;
            written += setup.Body.Columns.Count;

            foreach ((int Row, int? Column) tracked in _previousHighlights.Keys
                         .Where(entry => entry.Row == rowNumber)
                         .ToArray())
            {
                _ = _previousHighlights.Remove(tracked);
            }

            foreach (KeyValuePair<(int Row, int? Column), GanttValidationSeverity> entry in row)
            {
                Excel.Range? cell = GetCellRange(setup.Body, entry.Key.Row, entry.Key.Column!.Value);
                if (cell is null)
                {
                    continue;
                }

                cell.Interior.Color = ResolveOleColor(entry.Value);
                _previousHighlights[entry.Key] = entry.Value;
                written++;
            }
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
            GanttValidationSeverity.Warning => _oleWarningFill,
            _ => (int)XlColorIndex.xlColorIndexNone,
        };

    /// <summary>
    /// Converts a Core colour to the OLE_COLOR form Excel writes to
    /// <c>Interior.Color</c>: <c>0x00BBGGRR</c>.
    /// </summary>
    /// <param name="colour">The Core colour to convert.</param>
    /// <returns>The OLE_COLOR value.</returns>
    private static int ToOleColor(ColourHex colour)
    {
        uint argb = colour.ARGB;
        int red = (int)((argb >> 16) & 0xFF);
        int green = (int)((argb >> 8) & 0xFF);
        int blue = (int)(argb & 0xFF);
        return (blue << 16) | (green << 8) | red;
    }

    /// <summary>
    /// Orders severities so one address carrying two findings keeps the stronger colour.
    /// </summary>
    /// <param name="severity">The severity to rank.</param>
    /// <returns>2 for <see cref="GanttValidationSeverity.Error"/>, 1 for a warning, 0 otherwise.</returns>
    private static int SeverityRank(GanttValidationSeverity severity) => severity switch
    {
        GanttValidationSeverity.Error => 2,
        GanttValidationSeverity.Warning => 1,
        _ => 0,
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

            // The persisted form is row|column for a cell and row| for a whole
            // row, so the column is the second field. Reading a third field (as
            // this did) silently reloaded every cell highlight as a whole-row
            // one, and the first reconcile after a reopen then cleared the whole
            // row instead of the single cell.
            _previousHighlights.Clear();
            foreach (string entry in value.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                string[] parts = entry.Split('|');
                if (parts.Length >= 2 && int.TryParse(parts[0], out int row))
                {
                    int? column = int.TryParse(parts[1], out int col) ? col : null;
                    // Severity is not persisted (the name format is row|column and a third
                    // field is what once reloaded cell fills as whole-row ones). An unknown
                    // severity makes the next reconcile re-assert the colour once.
                    _previousHighlights[(row, column)] = null;
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
            string value = string.Join(";", _previousHighlights.Keys.Select(address =>
                address.Column is { } col ? $"{address.Row}|{col}" : $"{address.Row}|"));
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

    /// <summary>
    /// Finds the data table by walking the sheets and each sheet's tables by
    /// one-based index through <see cref="GetSheetAt"/> and <see cref="GetTableAt(ListObjects, int)"/>.
    /// </summary>
    /// <remarks>
    /// Index traversal rather than <c>foreach</c> is deliberate: the class documents
    /// these two seams as the only way it touches the Excel collections, and a COM
    /// collection enumerator cannot be stubbed through an expression tree. The index
    /// walk also avoids the <c>GetEnumerator</c> cast the old loop needed.
    /// </remarks>
    /// <param name="sheets">The workbook's sheets collection.</param>
    /// <param name="worksheet">The sheet holding the data table, when found.</param>
    /// <param name="table">The data table, when found.</param>
    /// <returns><see langword="true"/> when the data table was found.</returns>
    private bool TryFindTable(
        Sheets sheets,
        out Worksheet? worksheet,
        out ListObject? table)
    {
        worksheet = null;
        table = null;

        try
        {
            int sheetCount = sheets.Count;
            for (int sheetIndex = 1; sheetIndex <= sheetCount; sheetIndex++)
            {
                if (GetSheetAt(sheets, sheetIndex) is not Worksheet ws)
                {
                    continue;
                }

                ListObjects listObjects = ws.ListObjects;
                int tableCount = listObjects.Count;
                for (int tableIndex = 1; tableIndex <= tableCount; tableIndex++)
                {
                    ListObject listObject = GetTableAt(listObjects, tableIndex);
                    if (GanttChartContract.IsDataTableName(listObject.Name))
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

    /// <summary>Returns the one-based body row's range. Test seam over the COM parameterised <c>Range.Rows[row]</c> property.</summary>
    internal virtual Excel.Range GetRowRange(Excel.Range body, int rowNumber) => body.Rows[rowNumber];

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
