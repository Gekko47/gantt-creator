using System.Globalization;
using System.Runtime.InteropServices;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live, read-only <see cref="IPanelGridMeasurementPort"/> that measures the
/// data panel's exact column widths, per-row body heights, and header row height
/// from the visible <c>tblGanttData</c> table.
/// </summary>
/// <param name="application">
/// The Excel application object (for example <c>ExcelDnaUtil.Application</c>), or
/// <see langword="null"/> when the host supplied none (unit tests, non-Excel
/// host). A foreign object fails the interface cast and degrades to the
/// no-active-workbook refusal with no mutation.
/// </param>
/// <remarks>
/// <para>
/// <b>This adapter takes no protection guard.</b> It mutates nothing, and worksheet
/// protection blocks writes rather than reads, so a read-only measurement must not
/// refuse a readable target. Enforcement belongs at the Refresh/write boundary
/// (R4.9). The guard parameter was removed rather than left accepted-and-unused,
/// because a held-but-unconsulted field would read as "this adapter checks
/// protection" to the next reader. See ADR-0024.
/// </para>
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
/// This adapter never mutates: it reads column widths, each body row's height, and
/// the header row's height, and constructs the <see cref="PanelCellGrid"/>. It never
/// writes a cell, changes a column width, or alters application state, so
/// <c>ProtectionGuardFirstTests</c> classifies it read-only.
/// </para>
/// </remarks>
public class ExcelPanelGridMeasurement(object? application) : IPanelGridMeasurementPort
{
    private readonly Excel.Application? _application = application as Excel.Application;

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

        // This adapter writes nothing, and worksheet protection blocks writes rather
        // than reads: column widths and row heights are readable on a protected sheet.
        // It therefore does NOT consult the protection guard. The guard's rule
        // (ADR-0008 D4) is "first check in every *mutating* adapter", and this one is
        // not one; keeping the consultation made a read-only measurement - the input
        // to a read-only diagnostic or export - unavailable on any protected target,
        // which is a capability restriction with no product behind it. Enforcement
        // belongs at the Refresh/write boundary (R4.9), and that row is not built yet,
        // so removing it here opens no window in which a protected target is written.
        //
        // `NoActiveWorkbook` is unaffected and still typed: the check below reads
        // `ActiveWorkbook` directly rather than inferring it from the guard.

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

        // Every body row is measured individually and the body range's aggregate
        // RowHeight is never consulted. That is not a preference: the reference
        // documents that a range of differing row heights "might return the height of
        // the first row or might return Null", so the aggregate is unreliable in
        // *both* directions for a non-uniform body. `DBNull` was catchable; the
        // "first row's height" case is a plain number that passes every check while
        // describing a body most of which is taller. Reading the rows removes the
        // ambiguity instead of detecting it.
        List<double> rowHeights = [];
        if (table.DataBodyRange is { } body)
        {
            for (var rowIndex = 1; rowIndex <= GetBodyRowCount(body); rowIndex++)
            {
                if (GetBodyRowAt(body, rowIndex) is not { } row || ToPoints(row.RowHeight) is not { } height)
                {
                    // One absent row must not become one guessed height. The panel's row
                    // positions are cumulative, so a single wrong height would displace
                    // every row below it and the result would look plausible.
                    return PanelGridOutcome.Refused(PanelGridRefusalReason.InvalidMeasurement);
                }

                // A ZERO height is passed through deliberately (ADR-0034): Excel reports a
                // HIDDEN row's height as zero, and a collapsed outline group hides its
                // child rows, so zero is a real measurement of the displayed sheet rather
                // than a missing one. Core accepts it and the rows below share its top;
                // a negative height is still refused there.
                rowHeights.Add(height);
            }
        }

        if (rowHeights.Count == 0)
        {
            return PanelGridOutcome.Refused(PanelGridRefusalReason.InvalidMeasurement);
        }

        // Section 4: the header is its own Excel row with its own height, so it is
        // measured separately rather than reusing a body height.
        if (ReadHeaderRowHeight(table) is not { } headerHeight)
        {
            return PanelGridOutcome.Refused(PanelGridRefusalReason.InvalidMeasurement);
        }

        // ADR-0030 D3: the sheet's absolute origin is measured, not assumed. It is the
        // first BODY row's top edge, so the header row's height is already included
        // and this value IS the plot's top (D1). Refused rather than defaulted when
        // the host reports no figure: a zero origin is exactly the page-coordinate
        // assumption this change removes, and it would reintroduce the misalignment
        // silently.
        if (ReadOriginTop(table) is not { } originTop
            || ReadOriginLeft(table) is not { } originLeft
            || ReadTopPaddingHeight(table) is not { } topPaddingHeight
            || ReadBottomPaddingHeight(table) is not { } bottomPaddingHeight)
        {
            return PanelGridOutcome.Refused(PanelGridRefusalReason.InvalidMeasurement);
        }

        // Core validates the grid: a non-positive or non-finite width or height, a
        // blank or duplicated column name, a missing required column, and a negative
        // or non-finite origin are all refused there with their own typed reasons
        // rather than duplicated here.
        PanelCellGridCreationOutcome created = PanelCellGrid.TryCreate(
            columns,
            rowHeights,
            headerHeight,
            includedColumns,
            originTop,
            originLeft,
            topPaddingHeight,
            bottomPaddingHeight);
        return created.Succeeded && created.Grid is not null
            ? PanelGridOutcome.Ok(created.Grid)
            : PanelGridOutcome.Refused(PanelGridRefusalReason.InvalidMeasurement);
    }

    /// <summary>
    /// Reads the absolute worksheet Y of the first body row's top edge, in points
    /// (ADR-0030 D3).
    /// </summary>
    /// <param name="table">The Gantt table.</param>
    /// <returns>The origin, or <see langword="null"/> when the host reported none.</returns>
    /// <remarks>
    /// Read from <c>DataBodyRange</c>, not from the header range: the plot's top is
    /// the first BODY row's top, which is the header's bottom. Reading the header's
    /// top instead would put the plot a whole header-row too high — the same class of
    /// off-by-one-row error ADR-0030 corrects elsewhere.
    /// </remarks>
    internal virtual double? ReadOriginTop(Excel.ListObject table)
    {
        ArgumentNullException.ThrowIfNull(table);

        Excel.Range? body = table.DataBodyRange;
        return body is null ? null : ToPoints(body.Top);
    }

    /// <summary>
    /// Reads the height of the top padding row, in points (ADR-0031 D2).
    /// </summary>
    /// <param name="table">The Gantt table.</param>
    /// <returns>The height, or <see langword="null"/> when the host reported none.</returns>
    /// <remarks>
    /// The row index comes from <see cref="GanttSheetLayout.TopPaddingRowIndex"/>
    /// rather than a literal, so this adapter cannot disagree with the layout
    /// authority about which row the chart's top margin is.
    /// </remarks>
    internal virtual double? ReadTopPaddingHeight(Excel.ListObject table) =>
        ReadRowHeightAt(table, GanttSheetLayout.TopPaddingRowIndex);

    /// <summary>
    /// Reads the height of the bottom padding row, in points (ADR-0031 D2).
    /// </summary>
    /// <param name="table">The Gantt table.</param>
    /// <returns>The height, or <see langword="null"/> when the host reported none.</returns>
    /// <remarks>
    /// This row sits below a body whose length the user controls, so its index is
    /// derived from the measured body row count through
    /// <see cref="GanttSheetLayout.BottomPaddingRowIndex"/>. A literal here would be
    /// correct for exactly one table length and wrong for every other.
    /// </remarks>
    internal virtual double? ReadBottomPaddingHeight(Excel.ListObject table) =>
        table.DataBodyRange is not { } body
            ? null
            : ReadRowHeightAt(table, GanttSheetLayout.BottomPaddingRowIndex(GetBodyRowCount(body)));

    /// <summary>
    /// Reads one absolute worksheet row's height by its 1-based index.
    /// </summary>
    /// <param name="table">The Gantt table, used only to reach its worksheet.</param>
    /// <param name="rowIndex">The 1-based worksheet row index.</param>
    /// <returns>The height, or <see langword="null"/> when it cannot be read.</returns>
    /// <remarks>
    /// <b>Whole-worksheet rows, not table rows.</b> Both padding rows lie outside
    /// <c>tblGanttData</c>, so neither is reachable through
    /// <c>DataBodyRange</c> or <c>HeaderRowRange</c>; they are addressed through the
    /// worksheet's own <c>Rows</c> collection by absolute index.
    /// </remarks>
    private double? ReadRowHeightAt(Excel.ListObject table, int rowIndex)
    {
        ArgumentNullException.ThrowIfNull(table);

        // `ListObject` exposes no `Worksheet` member in this PIA, so the parent is
        // unwrapped explicitly. `Parent` is typed `object`, so a non-worksheet parent
        // arrives as something that will not cast; that degrades to the absent case
        // below rather than throwing.
        if (table.Parent is not Excel.Worksheet worksheet || worksheet.Rows is not { } rows)
        {
            return null;
        }

        Excel.Range? row = null;
        try
        {
            row = rows[rowIndex];
            return row is null ? null : ToPoints(row.RowHeight);
        }
        finally
        {
            // This adapter OWNS the Range it created and no caller can release it: the
            // proxy is never returned. Without this it would survive to the finaliser
            // and land in the Office leak ratchet, which is exactly the signal the
            // ratchet exists to keep meaningful.
            if (row is not null && Marshal.IsComObject(row))
            {
                _ = Marshal.FinalReleaseComObject(row);
            }
        }
    }

    /// <summary>
    /// Reads the absolute worksheet X of the panel's left edge, in points
    /// (ADR-0030 D3).
    /// </summary>
    /// <param name="table">The Gantt table.</param>
    /// <returns>The origin, or <see langword="null"/> when the host reported none.</returns>
    /// <remarks>
    /// The table does not begin at column A: the engine columns before
    /// <c>Type</c> are hidden but still occupy worksheet positions, so this is a
    /// measurement. Reading the header row's left would give the same X — the panel's
    /// left edge is the table's left edge — and the header range is the member this
    /// adapter already measures through.
    /// </remarks>
    internal virtual double? ReadOriginLeft(Excel.ListObject table)
    {
        ArgumentNullException.ThrowIfNull(table);

        Excel.Range? body = table.DataBodyRange;
        return body is null ? null : ToPoints(body.Left);
    }

    /// <summary>
    /// Reads the table's header row height in points. Test seam over
    /// <c>ListObject.HeaderRowRange</c> and its <c>RowHeight</c>.
    /// </summary>
    /// <param name="table">The Gantt table.</param>
    /// <returns>The header height, or <see langword="null"/> when the host reported none.</returns>
    internal virtual double? ReadHeaderRowHeight(Excel.ListObject table)
    {
        ArgumentNullException.ThrowIfNull(table);

        return table.HeaderRowRange is { } header ? ToPoints(header.RowHeight) : null;
    }

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
