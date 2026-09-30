using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using GanttCreator.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live adapter that writes the engine-owned <c>Duration</c> column in ONE bulk
/// pass over the column range (R4.7F D5).
/// </summary>
/// <param name="application">
/// The Excel application object, or <see langword="null"/> when the host supplied
/// none. A foreign object fails the interface cast and degrades to the
/// no-active-workbook refusal with no mutation.
/// </param>
/// <param name="protectionGuard">
/// The protection guard, consulted first per ADR-0008 D4 because this adapter
/// mutates.
/// </param>
/// <remarks>
/// <para>
/// <b>Why one ranged write rather than N cell writes.</b> D5 requires it and the cost
/// is observable: N COM round-trips over a 1,000-event table dominate Refresh time
/// (REV5 §22 instruction 12). The adapter therefore reads the whole <c>Duration</c>
/// column once, merges the plan into it, and assigns it back <b>once</b>. There is no
/// per-cell write path, so a regression that reintroduced one would have to add it.
/// </para>
/// <para>
/// <b>Why an empty plan writes nothing.</b> A correctly written sheet produces an
/// empty write set, and touching the range at all would mark the workbook dirty on
/// every Refresh.
/// </para>
/// <para>
/// <b>Why the adapter merges rather than overwrites.</b> The plan only names rows
/// whose value <em>changes</em>, so the range is read, patched at those rows, and
/// written back. Overwriting wholesale would be fewer reads but would stamp the
/// whole column on every Refresh.
/// </para>
/// <para>
/// COM ownership follows the existing adapters: the <c>Application</c>,
/// <c>Workbook</c>, <c>Worksheet</c>, <c>ListObject</c> and <c>Range</c> objects are
/// Excel-owned shared roots. This adapter takes no ownership and holds every proxy in
/// a local used without chained member expressions.
/// </para>
/// </remarks>
public class ExcelDurationWriter(
    object? application,
    IWorksheetProtectionGuard? protectionGuard = null) : IDurationWritePort
{
    private readonly Excel.Application? _application = application as Excel.Application;
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    /// <inheritdoc />
    public DurationWriteOutcome Write(DurationWritePlan plan)
    {
        ArgumentNullException.ThrowIfNull(plan);

        if (plan.WriteCount == 0)
        {
            // No-op: do not touch the range, so a correctly written sheet is not
            // marked dirty by a Refresh.
            return DurationWriteOutcome.Ok(0, 0);
        }

        Excel.Application? application = _application;
        Excel.Workbook? workbook = application?.ActiveWorkbook;
        if (workbook is null)
        {
            return DurationWriteOutcome.Refused(DurationWriteRefusalReason.NoActiveWorkbook);
        }

        if (!TryFindTable(workbook.Sheets, out Excel.Worksheet? worksheet, out Excel.ListObject? table)
            || worksheet is null
            || table is null)
        {
            return DurationWriteOutcome.Refused(DurationWriteRefusalReason.TableMissing);
        }

        // First check in a mutating adapter (ADR-0008 D4), before any COM write.
        ProtectionGuardOutcome protection = _protectionGuard.QueryTarget(worksheet);
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return DurationWriteOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? DurationWriteRefusalReason.NoActiveWorkbook
                    : DurationWriteRefusalReason.TargetProtected);
        }

        if (!TryGetDurationColumnRange(table, out Excel.Range? column) || column is null)
        {
            return DurationWriteOutcome.Refused(DurationWriteRefusalReason.TableMissing);
        }

        List<object?[]> current = ExcelValue2Matrix.ReadRows(GetRangeValues(column));
        if (current.Count == 0)
        {
            return DurationWriteOutcome.Refused(DurationWriteRefusalReason.WriteFailed);
        }

        // Merge the plan into the existing column, patching only the rows the plan
        // names. Row numbers are the one-based body index the plan uses, and the
        // matrix rows are in the same order.
        var merged = new object?[current.Count];
        for (var index = 0; index < current.Count; index++)
        {
            merged[index] = current[index].Length > 0 ? current[index][0] : null;
        }

        foreach (DurationWrite write in plan.Writes)
        {
            var target = write.RowNumber - 1;
            if (target < 0 || target >= merged.Length)
            {
                continue;
            }

            // An empty text is written as null, which is how Excel stores a blank
            // cell; writing "" would leave a zero-length string the reader would
            // report as present-but-empty.
            merged[target] = write.Text.Length == 0 ? null : write.Text;
        }

        try
        {
            SetRangeValues(column, ToTwoDimensional(merged));
        }
        catch (COMException)
        {
            return DurationWriteOutcome.Refused(DurationWriteRefusalReason.WriteFailed);
        }

        return DurationWriteOutcome.Ok(plan.WriteCount, plan.WriteCount);
    }

    /// <summary>
    /// Converts the merged column into the rectangular payload Excel's
    /// <c>Value2</c> assignment requires.
    /// </summary>
    /// <param name="values">The merged column values, one per body row.</param>
    /// <returns>A one-based <c>object[,]</c> for assignment.</returns>
    [SuppressMessage(
        "Performance",
        "CA1814:MultiDimArrays",
        Justification = "Excel's Value2 is a SAFEARRAY and is written through a multidimensional array. A jagged array cannot express it, so the analyzer's preferred shape is not available here.")]
    [SuppressMessage(
        "Performance",
        "CA1859:UseConcreteTypesWhenPossible",
        Justification = "The payload is assigned to a COM property typed as object, so the declared return type must stay object.")]
    private static object ToTwoDimensional(object?[] values)
    {
        var payload = new object[values.Length, 1];
        for (var row = 0; row < values.Length; row++)
        {
            payload[row, 0] = values[row]!;
        }

        return payload;
    }

    /// <summary>
    /// Locates the <c>Duration</c> column's body range. Test seam over the COM
    /// parameterised properties.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <param name="column">The body range of the <c>Duration</c> column.</param>
    /// <returns>Whether the column was found and has a body.</returns>
    internal virtual bool TryGetDurationColumnRange(
        Excel.ListObject table,
        out Excel.Range? column)
    {
        ArgumentNullException.ThrowIfNull(table);

        column = null;
        Excel.ListColumns? columns = table.ListColumns;
        if (columns is null)
        {
            return false;
        }

        var columnCount = columns.Count;
        for (var index = 1; index <= columnCount; index++)
        {
            Excel.ListColumn candidate = GetColumnAt(columns, index);
            if (string.Equals(candidate.Name, "Duration", StringComparison.Ordinal))
            {
                column = candidate.DataBodyRange;
                return column is not null;
            }
        }

        return false;
    }

    /// <summary>Reads the range's bulk <c>Value2</c> payload. Test seam.</summary>
    /// <param name="range">The range to read.</param>
    /// <returns>The raw payload.</returns>
    internal virtual object? GetRangeValues(Excel.Range range)
    {
        ArgumentNullException.ThrowIfNull(range);
        return range.Value2;
    }

    /// <summary>Assigns the range's bulk <c>Value2</c> payload. Test seam.</summary>
    /// <param name="range">The range to write.</param>
    /// <param name="payload">The two-dimensional payload.</param>
    internal virtual void SetRangeValues(Excel.Range range, object payload)
    {
        ArgumentNullException.ThrowIfNull(range);
        range.Value2 = payload;
    }

    /// <summary>Finds the Gantt worksheet and its table. Test seam.</summary>
    /// <param name="sheets">The workbook's sheet collection.</param>
    /// <param name="worksheet">The resolved worksheet.</param>
    /// <param name="table">The resolved table.</param>
    /// <returns>Whether both were found.</returns>
    internal virtual bool TryFindTable(
        Excel.Sheets sheets,
        out Excel.Worksheet? worksheet,
        out Excel.ListObject? table)
    {
        worksheet = null;
        table = null;
        foreach (Excel.Worksheet candidate in sheets)
        {
            foreach (Excel.ListObject candidateTable in candidate.ListObjects)
            {
                if (string.Equals(candidateTable.Name, "tblGanttData", StringComparison.Ordinal))
                {
                    worksheet = candidate;
                    table = candidateTable;
                    return true;
                }
            }
        }

        return false;
    }

    /// <summary>Returns the list column at the one-based index. Test seam.</summary>
    /// <param name="columns">The table's list columns.</param>
    /// <param name="index">The one-based column index.</param>
    /// <returns>The list column.</returns>
    internal virtual Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index) => columns[index];
}
