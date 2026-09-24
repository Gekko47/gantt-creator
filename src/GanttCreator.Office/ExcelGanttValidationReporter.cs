using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live <see cref="IGanttValidationReporter"/> over the Excel application object
/// supplied by the host at add-in load.
/// </summary>
/// <param name="application">
/// The Excel application object (for example <c>ExcelDnaUtil.Application</c>), or
/// <see lang="null"/> when the host supplied none (unit tests, non-Excel
/// host). A foreign object fails the interface cast and degrades to the
/// no-active-workbook refusal with no mutation.
/// </param>
/// <remarks>
/// <para>
/// COM ownership: the <c>Application</c>, <c>Workbook</c>, <c>Worksheet</c>,
/// <c>Range</c>, <c>ListObject</c>, <c>Comment</c>, and <c>SpecialCells</c>
/// objects reached here are Excel-owned shared roots. This adapter takes no
/// ownership of them, never calls <c>FinalReleaseComObject</c>, and
/// force-releases nothing (the ownership policy of
/// <see cref="ExcelApplicationAdapter"/>). Every proxy is held in a local and
/// used without chained member expressions; sheets are reached by index through
/// <see cref="GetSheetAt"/> and tables through <c>GetTableAt</c>.
/// </para>
/// <para>
/// Mutation order (work item R2.6 decision D-A): first locate the table and
/// build the column map, then remove the add-in's own report section from every
/// noted cell in the table body, then write one section per grouped cell. Only
/// add-in sections are touched; user-authored note text is preserved verbatim as
/// a prefix. Excel allows one classic note per cell and <c>Range.AddComment</c>
/// fails when a note exists (MS Learn, <c>Range.AddComment</c>), so a cell that
/// already carries a note is updated through <c>Comment.Text</c> with the
/// <c>Start</c> argument omitted — documented to delete the existing text and
/// accept the replacement — which also keeps repeated runs from duplicating
/// sections.
/// </para>
/// <para>
/// The <c>internal virtual</c> accessors (<see cref="GetSheetAt"/>,
/// <c>GetTableAt</c> (both <c>ListObjects</c> and <c>ListColumns</c> overloads),
/// <see cref="GetCellRange"/>, and <see cref="EnumerateCells"/>) isolate the Excel
/// COM parameterised properties (indexers). Expression trees cannot contain
/// indexed properties (CS0855), so contract tests substitute these seams and
/// every other member through Moq; the real indexer behaviour is exercised by
/// the tagged live-Office integration test.
/// </para>
/// </remarks>
/// <param name="protectionGuard">The shared read-only workbook-protection guard; defaults to a live guard over <paramref name="application"/>.</param>
public class ExcelGanttValidationReporter(
    object? application,
    IWorksheetProtectionGuard? protectionGuard = null) : IGanttValidationReporter
{
    private readonly Application? _application = application as Application;
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    /// <summary>
    /// Writes cell notes for each issue onto the offending cells of
    /// <c>tblGanttData</c>, clearing prior owned notes first. Idempotent.
    /// </summary>
    /// <param name="issues">
    /// Findings in <c>Severity → Row → Field → Code</c> order.
    /// </param>
    /// <returns>
    /// The count of notes written, or a typed refusal. On a refusal nothing
    /// was mutated.
    /// </returns>
    public GanttValidationReportOutcome Report(IReadOnlyList<Core.GanttValidationIssue> issues)
    {
        Application? application = _application;
        if (application is null)
        {
            return GanttValidationReportOutcome.Refused(
                GanttValidationReportRefusalReason.NoActiveWorkbook);
        }

        // Read-only setup first: resolve the table and column map. No mutation
        // happens before these succeed.
        Workbook? workbook = application.ActiveWorkbook;
        if (workbook is null)
        {
            return GanttValidationReportOutcome.Refused(
                GanttValidationReportRefusalReason.NoActiveWorkbook);
        }

        ProtectionGuardOutcome protection = _protectionGuard.Query();
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return GanttValidationReportOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? GanttValidationReportRefusalReason.NoActiveWorkbook
                    : GanttValidationReportRefusalReason.TargetProtected);
        }

        Sheets sheets = workbook.Sheets;
        if (!TryFindTable(sheets, out ListObject? table) || table is null)
        {
            return GanttValidationReportOutcome.Refused(
                GanttValidationReportRefusalReason.TableMissing);
        }

        if (!TryBuildColumnMap(table, out var columnMap) || columnMap is null)
        {
            return GanttValidationReportOutcome.Refused(
                GanttValidationReportRefusalReason.TableMissing);
        }

        Excel.Range? body = table.DataBodyRange;

        // Pure composition — no COM here.
        IReadOnlyList<NotePayload> notes = ValidationReportComposer.GroupIntoNotes(issues);

        // Stale add-in sections are removed on every run, not only when there is
        // something new to write: a row that has become valid must stop showing
        // its old note. User text is never removed.
        if (body is not null)
        {
            RemoveOwnedSections(body);
        }

        if (notes.Count == 0 || body is null)
        {
            return GanttValidationReportOutcome.Ok(0);
        }

        var written = 0;
        foreach (NotePayload note in notes)
        {
            var columnIndex = ResolveColumnIndex(columnMap, note.FieldName);
            if (columnIndex < 1)
            {
                // Field not present in the table — anchor on the row's Id cell
                // (always required) and name the field in the text instead.
                columnIndex = ResolveColumnIndex(columnMap, Core.GanttTableSchema.Default.Columns[0].Name);
            }

            Excel.Range cell = GetCellRange(body, note.RowNumber, columnIndex);
            WriteNote(cell, note.Text);
            written++;
        }

        return GanttValidationReportOutcome.Ok(written);
    }

    /// <summary>
    /// Removes the add-in's own report section from every noted cell in a range,
    /// leaving user-authored text untouched. A cell whose note is purely the
    /// add-in's loses the note entirely.
    /// </summary>
    /// <param name="range">The body range to scan.</param>
    /// <remarks>
    /// Ownership cannot be read from <c>Comment.Author</c> (read-only), so the
    /// section is recognised by the
    /// <see cref="ValidationReportComposer.NoteSentinelPrefix"/> text marker. The
    /// <c>SpecialCells</c> call throws when the range holds no notes at all,
    /// which is the expected "nothing to clear" case.
    /// </remarks>
    private void RemoveOwnedSections(Excel.Range range)
    {
        Excel.Range? existing;
        try
        {
            existing = range.SpecialCells(XlCellType.xlCellTypeComments);
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // No cells contain notes at all: nothing to clear. Expected and
            // benign — treated as "empty", not an unexpected COM failure.
            return;
        }

        if (existing is null)
        {
            return;
        }

        // SpecialCells may return several disjoint ranges; every cell is visited
        // through the enumeration seam so the decision logic is contract-testable.
        var toDelete = new List<Comment>();
        var toRewrite = new List<(Comment Comment, string Text)>();
        foreach (Excel.Range cell in EnumerateCells(existing))
        {
            Comment? comment = cell.Comment;
            if (comment is null)
            {
                continue;
            }

            var current = ReadCommentText(comment);
            if (current.IndexOf(ValidationReportComposer.NoteSentinelPrefix, StringComparison.Ordinal) < 0)
            {
                continue; // A user note we did not write: never touched.
            }

            var userText = ValidationReportComposer.StripOwnedSection(current);
            if (userText.Length == 0)
            {
                toDelete.Add(comment);
            }
            else
            {
                toRewrite.Add((comment, userText));
            }
        }

        foreach (Comment comment in toDelete)
        {
            comment.Delete();
        }

        foreach ((Comment comment, var userText) in toRewrite)
        {
            // Start omitted replaces the whole note text (MS Learn, Comment.Text).
            _ = comment.Text(userText);
        }
    }

    /// <summary>
    /// Writes the add-in report section into a cell's note, creating the note
    /// when the cell has none and otherwise replacing only the add-in section so
    /// the user's own text survives verbatim as a prefix.
    /// </summary>
    /// <param name="cell">The cell to annotate.</param>
    /// <param name="text">The report text (including the ownership marker).</param>
    private static void WriteNote(Excel.Range cell, string text)
    {
        Comment? existing = cell.Comment;
        if (existing is null)
        {
            // Only valid when the cell has no note: Range.AddComment fails on a
            // cell that already carries one.
            _ = cell.AddComment(text);
            return;
        }

        var userText = ValidationReportComposer.StripOwnedSection(ReadCommentText(existing));
        var combined = userText.Length == 0
            ? text
            : userText + Environment.NewLine + text;

        // Start omitted replaces the whole note text, so a repeat run cannot
        // accumulate duplicate sections and no trailing text can survive.
        _ = existing.Text(combined);
    }

    /// <summary>
    /// Reads the current text of a comment via <c>Comment.Text</c>. Isolated so
    /// contract tests can stub it; ownership detection depends on it.
    /// </summary>
    /// <param name="comment">The comment.</param>
    /// <returns>The current comment text.</returns>
    private static string ReadCommentText(Comment comment)
    {
        object? result = comment.Text();
        return result is null ? string.Empty : result.ToString() ?? string.Empty;
    }

    /// <summary>
    /// Resolves the table-column index for a schema field name from the column
    /// map built from live headers (Ordinal). Returns the one-based index, or
    /// <c>0</c> when the column is absent.
    /// </summary>
    private static int ResolveColumnIndex(int[] columnMap, string fieldName)
    {
        IReadOnlyList<Core.GanttTableColumn> schema = Core.GanttTableSchema.Default.Columns;
        for (var i = 0; i < schema.Count; i++)
        {
            if (string.Equals(schema[i].Name, fieldName, StringComparison.Ordinal))
            {
                return columnMap[i];
            }
        }

        return 0;
    }

    /// <summary>
    /// Scans every worksheet for a table named <c>tblGanttData</c>.
    /// </summary>
    /// <param name="sheets">The workbook's sheets.</param>
    /// <param name="table">The located table, or <see lang="null"/>.</param>
    /// <returns><see lang="true"/> when the table was found.</returns>
    private bool TryFindTable(Sheets sheets, out ListObject? table)
    {
        table = null;
        var count = sheets.Count;
        for (var index = 1; index <= count; index++)
        {
            var sheet = GetSheetAt(sheets, index);
            if (sheet is Worksheet worksheet)
            {
                ListObjects listObjects = worksheet.ListObjects;
                var tableCount = listObjects.Count;
                for (var tableIndex = 1; tableIndex <= tableCount; tableIndex++)
                {
                    ListObject candidate = GetTableAt(listObjects, tableIndex);
                    if (string.Equals(
                        candidate.Name,
                        Core.GanttTableSchema.TableName,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        table = candidate;
                        return true;
                    }
                }
            }
        }

        return false;
    }

    /// <summary>
    /// Builds the schema-column to table-column index map from live header
    /// names (Ordinal), so a reordered table still maps correctly.
    /// </summary>
    /// <param name="table">The Gantt data table.</param>
    /// <param name="columnMap">Per-schema-column one-based table indexes, or <see lang="null"/>.</param>
    /// <returns><see lang="true"/> when every required header resolved.</returns>
    private bool TryBuildColumnMap(ListObject table, out int[]? columnMap)
    {
        columnMap = null;
        IReadOnlyList<Core.GanttTableColumn> schema = Core.GanttTableSchema.Default.Columns;
        var map = new int[schema.Count];
        var byName = new Dictionary<string, int>(StringComparer.Ordinal);
        ListColumns columns = table.ListColumns;
        var columnCount = columns.Count;
        for (var index = 1; index <= columnCount; index++)
        {
            ListColumn column = GetTableAt(columns, index);
            byName[column.Name] = index;
        }

        for (var schemaIndex = 0; schemaIndex < schema.Count; schemaIndex++)
        {
            Core.GanttTableColumn expected = schema[schemaIndex];
            if (!byName.TryGetValue(expected.Name, out var tableIndex))
            {
                if (expected.IsRequired)
                {
                    return false;
                }

                map[schemaIndex] = -1;
            }
            else
            {
                map[schemaIndex] = tableIndex; // 1-based list-object column index == 1-based body cell column
            }
        }

        columnMap = map;
        return true;
    }

    // The three accessors below are the COM-parameterised-property seams. They are
    // kept as `internal virtual` (not protected) to match ExcelGanttTableReader
    // and so Moq can derive a TestableReporter without naming Excel types.

    /// <summary>
    /// Returns the sheet at the one-based index. Test seam over the COM
    /// parameterised <c>Sheets.Item</c> property (see the type remarks).
    /// </summary>
    internal virtual object GetSheetAt(Sheets sheets, int index)
        => sheets[index];

    /// <summary>
    /// Returns the list object at the one-based index. Test seam over the COM
    /// parameterised <c>ListObjects.Item</c> property (see the type remarks).
    /// </summary>
    internal virtual ListObject GetTableAt(ListObjects listObjects, int index)
        => listObjects[index];

    /// <summary>
    /// Returns the list column at the one-based index. Test seam over the COM
    /// parameterised <c>ListColumns.Item</c> property (see the type remarks).
    /// </summary>
    internal virtual ListColumn GetTableAt(ListColumns columns, int index)
        => columns[index];

    /// <summary>
    /// Returns the cell at the one-based body row and one-based table-column
    /// index. Test seam over the COM <c>Range.Cells[row, column]</c> indexer, so
    /// contract tests can intercept per-cell writes without a real worksheet.
    /// </summary>
    internal virtual Excel.Range GetCellRange(Excel.Range? body, int rowNumber, int columnIndex)
        => body!.Cells[rowNumber, columnIndex];

    /// <summary>
    /// Enumerates the individual cells of a range that holds notes. Test seam
    /// over the COM range enumerator (see the type remarks), so the
    /// stale-section decision logic is contract-testable without live Excel.
    /// </summary>
    /// <param name="range">The range returned by <c>SpecialCells(xlCellTypeComments)</c>.</param>
    /// <returns>The individual cells in the range, one at a time.</returns>
    internal virtual IEnumerable<Excel.Range> EnumerateCells(Excel.Range range)
    {
        foreach (var item in range.Cells)
        {
            if (item is Excel.Range cell)
            {
                yield return cell;
            }
        }
    }
}
