using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live, read-only <see cref="IPanelGridMeasurementPort"/> that measures the
/// data panel's exact column widths and body row height from the visible
/// <c>tblGanttData</c> table.
/// </summary>
/// <param name="application">
/// The Excel application object (for example <c>ExcelDnaUtil.Application</c>), or
/// <see langword="null"/> when the host supplied none (unit tests, non-Excel
/// host). A foreign object fails the interface cast and degrades to the
/// no-active-workbook refusal with no mutation.
/// </param>
/// <param name="protectionGuard">The shared read-only workbook-protection guard.</param>
/// <remarks>
/// <para>
/// COM ownership: the <c>Application</c>, <c>Workbook</c>, <c>Worksheet</c>,
/// <c>ListObject</c>, <c>ListColumn</c>, and <c>Range</c> objects reached here
/// are Excel-owned shared roots. This adapter takes no ownership of them, never
/// calls <c>FinalReleaseComObject</c>, and force-releases nothing. Every proxy
/// is held in a local and used without chained member expressions.
/// </para>
/// <para>
/// The <c>internal virtual</c> accessors isolate the Excel COM parameterised
/// properties (indexers) so contract tests can substitute them (CS0855); the
/// real indexer and <c>Width</c>/<c>RowHeight</c> behaviour is exercised by the
/// tagged live-Office integration test.
/// </para>
/// <para>
/// This adapter never mutates: it reads column widths and one row height and
/// constructs the <see cref="PanelCellGrid"/>. It never writes a cell, changes
/// a column width, or alters application state, so
/// <c>ProtectionGuardFirstTests</c> classifies it read-only.
/// </para>
/// </remarks>
public class ExcelPanelGridMeasurement(
    object? application,
    IWorksheetProtectionGuard? protectionGuard = null) : IPanelGridMeasurementPort
{
    private readonly Excel.Application? _application = application as Excel.Application;
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    /// <inheritdoc />
    public PanelGridOutcome Measure(IReadOnlyList<string> includedColumns)
    {
        ArgumentNullException.ThrowIfNull(includedColumns);

        if (includedColumns.Count == 0)
        {
            return PanelGridOutcome.Refused(PanelGridRefusalReason.InvalidMeasurement);
        }

        Excel.Application? application = _application;
        if (application is null)
        {
            return PanelGridOutcome.Refused(PanelGridRefusalReason.NoActiveWorkbook);
        }

        // The guard is the first read-only check for every adapter that reaches
        // the workbook (ADR-0008 D4). This one never mutates, but a protected
        // sheet is not measurable here either, and refusing early keeps one guard
        // for the whole mutating command path.
        ProtectionGuardOutcome protection = _protectionGuard.Query();
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return PanelGridOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? PanelGridRefusalReason.NoActiveWorkbook
                    : PanelGridRefusalReason.TargetProtected);
        }

        Excel.Workbook? workbook = application.ActiveWorkbook;
        if (workbook is null)
        {
            return PanelGridOutcome.Refused(PanelGridRefusalReason.NoActiveWorkbook);
        }

        Excel.Sheets? sheets = workbook.Sheets;
        if (sheets is null)
        {
            return PanelGridOutcome.Refused(PanelGridRefusalReason.TableMissing);
        }

        Excel.ListObject? table = FindGanttTable(sheets);
        if (table is null)
        {
            return PanelGridOutcome.Refused(PanelGridRefusalReason.TableMissing);
        }

        List<PanelColumn> columns = [];
        foreach (var name in includedColumns)
        {
            Excel.ListColumn? column = FindColumn(table, name);
            if (column is null)
            {
                return PanelGridOutcome.Refused(PanelGridRefusalReason.TableMissing);
            }

            var width = ReadColumnWidth(column);
            if (width is null)
            {
                return PanelGridOutcome.Refused(PanelGridRefusalReason.InvalidMeasurement);
            }

            columns.Add(new PanelColumn(name, width.Value));
        }

        var rowHeight = ReadRowHeight(table);
        if (rowHeight is null)
        {
            return PanelGridOutcome.Refused(PanelGridRefusalReason.InvalidMeasurement);
        }

        // Core validates the grid: a non-positive or non-finite width, a blank or
        // duplicated column name, and a missing required column are all refused
        // there with their own typed reasons rather than duplicated here.
        //
        // The grid now carries one height per body row plus its own header height.
        // This adapter still measures a single confirmed-uniform body height and
        // refuses a mixed body, so the row list is built by replicating that one
        // confirmed value once per body row. That preserves today's behaviour and
        // geometry exactly; measuring the rows individually - and the header row
        // separately from the body - is the adapter's own change, tracked as
        // remediation Commit C and deliberately not smuggled in here.
        var rowCount = GetBodyRowCountOf(table);
        if (rowCount <= 0)
        {
            return PanelGridOutcome.Refused(PanelGridRefusalReason.InvalidMeasurement);
        }

        PanelCellGridCreationOutcome created = PanelCellGrid.TryCreate(
            columns,
            [.. Enumerable.Repeat(rowHeight.Value, rowCount)],
            rowHeight.Value,
            includedColumns);
        return created.Succeeded && created.Grid is not null
            ? PanelGridOutcome.Ok(created.Grid)
            : PanelGridOutcome.Refused(PanelGridRefusalReason.InvalidMeasurement);
    }

    /// <summary>Counts the body rows, or returns zero when the body cannot be resolved.</summary>
    private int GetBodyRowCountOf(Excel.ListObject table) =>
        table.DataBodyRange is { } body ? GetBodyRowCount(body) : 0;
    // ---- Test seams (internal virtual, per the ExcelGanttTableReader pattern) ----

    /// <summary>
    /// Finds the <c>tblGanttData</c> table on the workbook. Test seam over the COM
    /// parameterised sheet and table indexers (CS0855).
    /// </summary>
    /// <param name="sheets">The workbook's sheet collection.</param>
    /// <returns>The Gantt table, or <see langword="null"/> when absent.</returns>
    internal virtual Excel.ListObject? FindGanttTable(Excel.Sheets sheets)
    {
        ArgumentNullException.ThrowIfNull(sheets);

        var sheetCount = GetSheetCount(sheets);
        for (var sheetIndex = 1; sheetIndex <= sheetCount; sheetIndex++)
        {
            Excel._Worksheet? sheet = GetSheetAt(sheets, sheetIndex);
            if (sheet is null)
            {
                continue;
            }

            Excel.ListObjects? objects = sheet.ListObjects;
            if (objects is null)
            {
                continue;
            }

            var tableCount = GetListObjectCount(objects);
            for (var tableIndex = 1; tableIndex <= tableCount; tableIndex++)
            {
                Excel.ListObject? table = GetListObjectAt(objects, tableIndex);
                if (table is not null &&
                    string.Equals(table.Name, GanttTableSchema.TableName, StringComparison.Ordinal))
                {
                    return table;
                }
            }
        }

        return null;
    }

    /// <summary>Reads the sheet count. Test seam over <c>Worksheets.Count</c>.</summary>
    /// <param name="sheets">The sheet collection.</param>
    /// <returns>The number of sheets.</returns>
    internal virtual int GetSheetCount(Excel.Sheets sheets) => sheets.Count;

    /// <summary>Reads one sheet by index. Test seam over <c>Sheets.Item</c>.</summary>
    /// <param name="sheets">The sheet collection.</param>
    /// <param name="index">The 1-based sheet index.</param>
    /// <returns>The sheet, or <see langword="null"/>.</returns>
    internal virtual Excel._Worksheet? GetSheetAt(Excel.Sheets sheets, int index) =>
        sheets[index] as Excel._Worksheet;

    /// <summary>Reads the table count. Test seam over <c>ListObjects.Count</c>.</summary>
    /// <param name="objects">The table collection.</param>
    /// <returns>The number of tables.</returns>
    internal virtual int GetListObjectCount(Excel.ListObjects objects) => objects.Count;

    /// <summary>Reads one table by index. Test seam over <c>ListObjects.Item</c>.</summary>
    /// <param name="objects">The table collection.</param>
    /// <param name="index">The 1-based table index.</param>
    /// <returns>The table, or <see langword="null"/>.</returns>
    internal virtual Excel.ListObject? GetListObjectAt(Excel.ListObjects objects, int index) =>
        objects[index];

    /// <summary>
    /// Finds a table column by its exact schema display name. Test seam over the
    /// <c>ListColumns.Item</c> indexed property and the <c>Name</c> read.
    /// </summary>
    /// <param name="table">The Gantt table.</param>
    /// <param name="name">The exact schema display name.</param>
    /// <returns>The column, or <see langword="null"/> when the table has no such column.</returns>
    internal virtual Excel.ListColumn? FindColumn(Excel.ListObject table, string name)
    {
        ArgumentNullException.ThrowIfNull(table);

        Excel.ListColumns? columns = table.ListColumns;
        if (columns is null)
        {
            return null;
        }

        var count = GetColumnCount(columns);
        for (var index = 1; index <= count; index++)
        {
            Excel.ListColumn? column = GetColumnAt(columns, index);
            if (column is not null && string.Equals(column.Name, name, StringComparison.Ordinal))
            {
                return column;
            }
        }

        return null;
    }

    /// <summary>Reads the column count. Test seam over <c>ListColumns.Count</c>.</summary>
    /// <param name="columns">The column collection.</param>
    /// <returns>The number of columns.</returns>
    internal virtual int GetColumnCount(Excel.ListColumns columns) => columns.Count;

    /// <summary>Reads one column by index. Test seam over <c>ListColumns.Item</c>.</summary>
    /// <param name="columns">The column collection.</param>
    /// <param name="index">The 1-based column index.</param>
    /// <returns>The column, or <see langword="null"/>.</returns>
    internal virtual Excel.ListColumn? GetColumnAt(Excel.ListColumns columns, int index) =>
        columns[index];

    /// <summary>
    /// Reads a column's width in points. Test seam over the COM
    /// <c>ListColumn.Range.Width</c> property, which is <c>System.Object</c>.
    /// </summary>
    /// <param name="column">The column to measure.</param>
    /// <returns>The width in points, or <see langword="null"/> when the host returned no numeric value.</returns>
    /// <remarks>
    /// A <c>ListColumn</c> does not expose <c>Width</c> directly; it inherits it
    /// through its own <c>Range</c> (probed 2026-09-27: <c>ListColumn</c> declares
    /// a <c>Range</c> member and <c>IRange.Width</c> is <c>Object</c>). Reading
    /// the column's range is therefore the real host path, and the
    /// <c>Convert.ToDouble</c> with the invariant culture is required because the
    /// value arrives boxed as a host number.
    /// </remarks>
    internal virtual double? ReadColumnWidth(Excel.ListColumn column)
    {
        ArgumentNullException.ThrowIfNull(column);

        return column.Range is { } range ? ToPoints(range.Width) : null;
    }

    /// <summary>
    /// Reads the table's body row height in points. Test seam over the COM
    /// <c>Range.RowHeight</c> property, which is <c>System.Object</c>.
    /// </summary>
    /// <param name="table">The Gantt table.</param>
    /// <returns>The row height in points, or <see langword="null"/> when the host returned no numeric value.</returns>
    /// <remarks>
    /// <para>
    /// A body with <em>uniform</em> row heights is the expected case: the aggregate
    /// then reports that height. A body with <em>mixed</em> row heights has no
    /// single height, and the host is documented to report it inconsistently - the
    /// <c>Range.RowHeight</c> reference states that a range of differing row heights
    /// "might return the height of the first row or might return Null", and
    /// <c>DBNull.Value</c> is what the PIA surfaces for the Null case.
    /// </para>
    /// <para>
    /// Both encodings are therefore an absent measurement and become the typed
    /// <see cref="PanelGridRefusalReason.InvalidMeasurement"/> refusal; converting
    /// one would throw <see cref="InvalidCastException"/> out of a read-only adapter
    /// and escape into the render command.
    /// </para>
    /// <para>
    /// Neither encoding covers the host returning the FIRST row's height, which is
    /// a number and would sail through both checks while describing a body whose
    /// rows are not all that tall. So the aggregate is confirmed against every body
    /// row; a body that cannot be enumerated is refused rather than measured, since
    /// a single height is exactly what could not be established.
    /// </para>
    /// </remarks>
    internal virtual double? ReadRowHeight(Excel.ListObject table)
    {
        ArgumentNullException.ThrowIfNull(table);

        return table.DataBodyRange is { } body ? ConfirmUniformRowHeight(body) : null;
    }

    /// <summary>
    /// Converts the body range's aggregate height, but only once every body row has
    /// confirmed it.
    /// </summary>
    /// <param name="body">The table's body range.</param>
    /// <returns>The height in points, or <see langword="null"/> when it is absent or not uniform.</returns>
    private double? ConfirmUniformRowHeight(Excel.Range body) =>
        ToPoints(body.RowHeight) is { } aggregate
            ? HasUniformRowHeight(body, aggregate) ? aggregate : null
            : null;

    /// <summary>
    /// Determines whether every body row reports the same height as the aggregate.
    /// </summary>
    /// <param name="body">The table's body range.</param>
    /// <param name="aggregate">The height the body range reported.</param>
    /// <returns><see langword="true"/> when every body row agrees with the aggregate.</returns>
    /// <remarks>
    /// The heights are compared for exact equality rather than within a tolerance.
    /// Rows that are genuinely the same height report the identical host value, and
    /// a tolerance here would quietly accept a body whose rows differ by a
    /// fraction of a point - the mixed-height case this method exists to catch. A
    /// body that reports <see cref="DBNull"/> for any single row also fails, since
    /// that row's height is not established either.
    /// </remarks>
    private bool HasUniformRowHeight(Excel.Range body, double aggregate)
    {
        var count = GetBodyRowCount(body);
        if (count <= 0)
        {
            return false;
        }

        for (var index = 1; index <= count; index++)
        {
            if (GetBodyRowAt(body, index) is not { } row || ToPoints(row.RowHeight) is not { } height)
            {
                return false;
            }

            if (!double.Equals(height, aggregate))
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>Reads the number of rows in a range. Test seam over <c>Range.Rows.Count</c>.</summary>
    /// <param name="body">The body range.</param>
    /// <returns>The row count, or zero when the host does not report one.</returns>
    internal virtual int GetBodyRowCount(Excel.Range body)
    {
        ArgumentNullException.ThrowIfNull(body);

        Excel.Range? rows = body.Rows;
        return rows?.Count ?? 0;
    }

    /// <summary>Reads one body row by its 1-based index. Test seam over <c>Range.Rows.Item</c>.</summary>
    /// <param name="body">The body range.</param>
    /// <param name="index">The 1-based row index.</param>
    /// <returns>The row range, or <see langword="null"/> when the host does not resolve it.</returns>
    internal virtual Excel.Range? GetBodyRowAt(Excel.Range body, int index)
    {
        ArgumentNullException.ThrowIfNull(body);

        return body.Rows is { } rows ? rows[index] : null;
    }

    /// <summary>
    /// Converts one host-reported measurement to points, treating an absent
    /// measurement as an absent measurement.
    /// </summary>
    /// <param name="raw">The boxed host value, or <see langword="null"/>.</param>
    /// <returns>The points, or <see langword="null"/> when the host reported none.</returns>
    /// <remarks>
    /// Both absence encodings are refused, not converted:
    /// <list type="bullet">
    /// <item><description><see langword="null"/>, which the PIA surfaces for a
    /// member the host has no value for.</description></item>
    /// <item><description><see cref="DBNull.Value"/>, which an
    /// <c>Object</c>-typed Excel member returns when a multi-cell or multi-row
    /// range has no single value — a body with mixed row heights reports its
    /// <c>RowHeight</c> this way. <c>Convert.ToDouble(DBNull.Value)</c> throws
    /// <see cref="InvalidCastException"/>, so treating only <see langword="null"/>
    /// as absent would turn a measurable-in-principle table into an unhandled
    /// exception rather than the typed refusal.</description></item>
    /// </list>
    /// The invariant culture is required either way because the value arrives
    /// boxed as a host number.
    /// </remarks>
    private static double? ToPoints(object? raw) =>
        raw is null or DBNull ? null : Convert.ToDouble(raw, CultureInfo.InvariantCulture);
}
