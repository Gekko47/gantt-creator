using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;
using GanttCreator.Core;
using GanttCreator.Office;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// Live-Excel gate for row-level error reporting (work item R2.6). Tagged
/// <c>[Trait("Category","OfficeIntegration")]</c> so <c>verify-quick.ps1</c> and
/// <c>verify.ps1</c> exclude them; run via <c>pwsh ./scripts/verify-office.ps1</c>
/// on the self-hosted runner (ADR-0001).
/// </summary>
/// <remarks>
/// <para>
/// The roadmap marks the R2.6 Office gate <b>Required</b>: "invalid rows show
/// actionable errors while original cells remain unchanged". These tests supply
/// the automated half by driving the real
/// <see cref="IGanttValidationReporter"/> against a live workbook; the manual F5
/// walk-through supplies the human half.
/// </para>
/// <para>
/// The assertions mirror
/// <see cref="ContractTests.ValidationReporterTests"/> (Moq PIA, no live Office)
/// so a drift between the faked COM surface and the real one is caught. The
/// ownership rule is the reason the live test exists: Excel permits one classic
/// note per cell and <c>Range.AddComment</c> fails when a note already exists
/// (MS Learn, <c>Range.AddComment</c>), so the reporter writes through
/// <c>Comment.Text</c> with <c>Start</c> omitted — documented to delete the
/// existing text and accept the replacement — to keep the user's own text.
/// </para>
/// </remarks>
public class ValidationReporterIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    private static string XllPath => OfficeFixtureTests.ResolvePackedXllPath();

    /// <summary>
    /// The core R2.6 invariant, live: every cell value in the body is identical
    /// before and after a report, every finding is annotated on its own cell with
    /// the add-in marker, and a stale add-in section on a cell that no longer has
    /// a finding is removed.
    /// </summary>
    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Report_annotates_the_offending_cells_without_changing_any_cell_value()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(
                fixture.RegisterXll(XllPath),
                $"Application.RegisterXLL returned false for '{XllPath}'.");

            Excel.Workbook workbook = fixture.CreateWorkbook();
            var initialiser = new ExcelWorkbookInitialiser(fixture.Excel);
            WorkbookInitialiseOutcome initialised = initialiser.Initialise();
            Assert.True(initialised.Succeeded, $"Initialise refused: {initialised.Refusal}");

            Excel.Worksheet sheet =
                (Excel.Worksheet)workbook.Sheets[GanttWorkbookContract.GanttSheetLabel];
            Excel.ListObject table = sheet.ListObjects[GanttTableSchema.TableName];
            ClearTableBody(table);

            Excel.ListRow validRow = table.ListRows.Add();
            Excel.ListRow invalidRow = table.ListRows.Add();
            Excel.Range body = table.DataBodyRange;
            Assert.NotNull(body);

            int validOffset = validRow.Range.Row - body.Row + 1;
            int invalidOffset = invalidRow.Range.Row - body.Row + 1;
            _output.WriteLine(
                $"Excel build {fixture.Excel.Version} (PID {fixture.ProcessId}); " +
                $"body={body.Address}; validOffset={validOffset}; invalidOffset={invalidOffset}");

            const double startSerial = 46266.0; // 2026-09-01
            const double finishSerial = 46270.0; // 2026-09-05

            SetBodyCell(body, validOffset, "Id", NewId('1'));
            SetBodyCell(body, validOffset, "LaneId", NewId('a'));
            SetBodyCell(body, validOffset, "StackIndex", 0.0);
            SetBodyCell(body, validOffset, "Type", "As-Planned Activity");
            SetBodyCell(body, validOffset, "Description", "valid row");
            SetBodyCell(body, validOffset, "Start", startSerial);
            SetBodyCell(body, validOffset, "Finish", finishSerial);
            SetBodyCell(body, validOffset, "Visible", true);

            // The invalid row: a well-formed Id so the test can locate it, but a
            // Type that is not in the catalogue, which the validator rejects.
            SetBodyCell(body, invalidOffset, "Id", NewId('2'));
            SetBodyCell(body, invalidOffset, "LaneId", NewId('a'));
            SetBodyCell(body, invalidOffset, "StackIndex", 0.0);
            SetBodyCell(body, invalidOffset, "Type", "Not a catalogued type");
            SetBodyCell(body, invalidOffset, "Description", "invalid row");
            SetBodyCell(body, invalidOffset, "Start", startSerial);
            SetBodyCell(body, invalidOffset, "Finish", finishSerial);
            SetBodyCell(body, invalidOffset, "Visible", true);

            // A stale add-in section on a cell the invalid row has no finding for.
            Excel.Range staleCell = body.Cells[invalidOffset, GetColumnIndex("Description")];
            _ = staleCell.AddComment($"{Sentinel()} stale section from an earlier run");
            Assert.NotNull(staleCell.Comment);

            List<string> before = SnapshotFingerprints(body);

            GanttTableReadOutcome read = new ExcelGanttTableReader(fixture.Excel).Read();
            Assert.True(read.Succeeded, $"Read refused: {read.Refusal}");
            GanttValidationOutcome validation = GanttRowValidator.Validate(read.Rows);
            Assert.Contains(validation.Issues, issue => issue.Severity == GanttValidationSeverity.Error);

            GanttValidationReportOutcome report =
                new ExcelGanttValidationReporter(fixture.Excel).Report(validation.Issues);
            Assert.True(report.Succeeded, $"Report refused: {report.Refusal}");

            GanttRowDto valid = read.Rows.Single(row => row.Id == GetBodyCell(body, validOffset, "Id")!.ToString());
            GanttRowDto invalid = read.Rows.Single(row => row.Id == GetBodyCell(body, invalidOffset, "Id")!.ToString());
            Assert.True(valid.RowNumber > 0);
            Assert.True(invalid.RowNumber > 0);

            // 1) No cell value changed anywhere in the body.
            List<string> after = SnapshotFingerprints(body);
            AssertValuesIdentical(before, after);

            // 2) The offending cell carries the add-in section naming the field.
            Excel.Range invalidTypeCell = body.Cells[invalid.RowNumber, GetColumnIndex("Type")];
            Excel.Comment? typeComment = invalidTypeCell.Comment;
            Assert.NotNull(typeComment);
            string typeNote = typeComment!.Text();
            _output.WriteLine($"Type note: {typeNote}");
            Assert.StartsWith(Sentinel(), typeNote, StringComparison.Ordinal);
            Assert.Contains("'Type'", typeNote, StringComparison.Ordinal);
            Assert.Contains("Error", typeNote, StringComparison.Ordinal);

            // 3) The stale add-in section on a now-unannotated cell is gone.
            Excel.Comment? staleAfter = staleCell.Comment;
            Assert.True(
                staleAfter is null,
                $"Expected the stale add-in note to be removed, found: {staleAfter?.Text()}");

            // 4) The valid row is not annotated at all.
            Excel.Range validTypeCell = body.Cells[valid.RowNumber, GetColumnIndex("Type")];
            Assert.Null(validTypeCell.Comment);

            _output.WriteLine(
                $"issues={validation.Issues.Count}; notesWritten={report.NotesWritten}; " +
                $"invalidRow={invalid.RowNumber}; validRow={valid.RowNumber}");
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    /// <summary>
    /// The ownership rule, live: a note the user typed is never deleted and never
    /// replaced — the add-in section is appended to it, and a second run replaces
    /// only that section instead of accumulating duplicates.
    /// </summary>
    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Report_preserves_a_user_note_and_replaces_only_the_add_in_section()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(
                fixture.RegisterXll(XllPath),
                $"Application.RegisterXLL returned false for '{XllPath}'.");

            Excel.Workbook workbook = fixture.CreateWorkbook();
            Assert.True(new ExcelWorkbookInitialiser(fixture.Excel).Initialise().Succeeded);

            Excel.Worksheet sheet =
                (Excel.Worksheet)workbook.Sheets[GanttWorkbookContract.GanttSheetLabel];
            Excel.ListObject table = sheet.ListObjects[GanttTableSchema.TableName];
            ClearTableBody(table);

            Excel.ListRow row = table.ListRows.Add();
            Excel.Range body = table.DataBodyRange;
            Assert.NotNull(body);
            int offset = row.Range.Row - body.Row + 1;

            SetBodyCell(body, offset, "Id", NewId('3'));
            SetBodyCell(body, offset, "LaneId", NewId('a'));
            SetBodyCell(body, offset, "StackIndex", 0.0);
            SetBodyCell(body, offset, "Type", "Not a catalogued type");
            SetBodyCell(body, offset, "Start", 46266.0);
            SetBodyCell(body, offset, "Finish", 46270.0);
            SetBodyCell(body, offset, "Visible", true);

            const string userText = "Reminder I typed myself.";
            Excel.Range typeCell = body.Cells[offset, GetColumnIndex("Type")];
            _ = typeCell.AddComment(userText);

            GanttTableReadOutcome read = new ExcelGanttTableReader(fixture.Excel).Read();
            Assert.True(read.Succeeded, $"Read refused: {read.Refusal}");
            GanttValidationOutcome validation = GanttRowValidator.Validate(read.Rows);
            Assert.Contains(validation.Issues, issue => issue.Severity == GanttValidationSeverity.Error);

            var reporter = new ExcelGanttValidationReporter(fixture.Excel);
            GanttValidationReportOutcome first = reporter.Report(validation.Issues);
            Assert.True(first.Succeeded, $"Report refused: {first.Refusal}");

            Excel.Comment? afterFirst = typeCell.Comment;
            Assert.NotNull(afterFirst);
            string firstText = afterFirst!.Text();
            _output.WriteLine($"note after first run: {firstText}");
            Assert.StartsWith(userText, firstText, StringComparison.Ordinal);
            Assert.Contains(Sentinel(), firstText, StringComparison.Ordinal);

            // A second run must not accumulate a second section.
            GanttValidationReportOutcome second = reporter.Report(validation.Issues);
            Assert.True(second.Succeeded, $"Second report refused: {second.Refusal}");

            Excel.Comment? afterSecond = typeCell.Comment;
            Assert.NotNull(afterSecond);
            string secondText = afterSecond!.Text();
            _output.WriteLine($"note after second run: {secondText}");
            Assert.StartsWith(userText, secondText, StringComparison.Ordinal);
            Assert.Equal(
                1,
                CountOccurrences(secondText, Sentinel()));
            Assert.Equal(firstText, secondText);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    private static string Sentinel() => ValidationReportComposer.NoteSentinelPrefix;

    /// <summary>A well-formed <see cref="GanttRowId"/>: <c>G-</c> plus 32 hex characters.</summary>
    private static string NewId(char last) => "G-" + new string('0', 31) + last;

    private static int CountOccurrences(string text, string value)
    {
        var count = 0;
        var index = text.IndexOf(value, StringComparison.Ordinal);
        while (index >= 0)
        {
            count++;
            index = text.IndexOf(value, index + value.Length, StringComparison.Ordinal);
        }

        return count;
    }

    /// <summary>
    /// Snapshots every body cell as a <c>type:value</c> fingerprint so any change
    /// in a value, or in its type, is detected after the report runs.
    /// </summary>
    private static List<string> SnapshotFingerprints(Excel.Range body)
    {
        object? raw = body.Value2;
        if (raw is not Array matrix)
        {
            throw new InvalidOperationException(
                $"Body Value2 was {raw?.GetType().Name ?? "null"}, expected an array.");
        }

        var fingerprints = new List<string>();
        for (int row = 1; row <= matrix.GetLength(0); row++)
        {
            for (int column = 1; column <= matrix.GetLength(1); column++)
            {
                object? value = matrix.GetValue(row, column);
                fingerprints.Add(string.Create(
                    CultureInfo.InvariantCulture,
                    $"{row},{column}:{value?.GetType().Name ?? "null"}:{value}"));
            }
        }

        return fingerprints;
    }

    private static void AssertValuesIdentical(
        List<string> before, List<string> after)
    {
        Assert.Equal(before, after);
    }

    /// <summary>
    /// Removes every body row from the table. Uses the atomic
    /// <c>DataBodyRange.Delete</c> path when it exists and a per-row fallback when
    /// <c>DataBodyRange</c> is null (a COM quirk seen on this build where
    /// <c>ListRows.Count</c> was 1 but <c>DataBodyRange</c> was null right after
    /// <c>Initialise</c>). The fallback re-reads <c>.Count</c> each pass.
    /// </summary>
    private static void ClearTableBody(Excel.ListObject table)
    {
        if (table.DataBodyRange is not null)
        {
            table.DataBodyRange.Delete();
        }

        while (true)
        {
            int count = table.ListRows.Count;
            if (count == 0)
            {
                break;
            }

            Excel.ListRow? last = table.ListRows[count];
            if (last is null)
            {
                break;
            }

            last.Delete();
        }
    }

    private static void SetBodyCell(
        Excel.Range body, int bodyRowNumber, string headerName, object value)
    {
        Excel.Range cell = body.Cells[bodyRowNumber, GetColumnIndex(headerName)];
        cell.Value2 = value;
    }

    private static object? GetBodyCell(Excel.Range body, int bodyRowNumber, string headerName)
    {
        Excel.Range cell = body.Cells[bodyRowNumber, GetColumnIndex(headerName)];
        return cell.Value2;
    }

    private static int GetColumnIndex(string headerName)
    {
        IReadOnlyList<GanttTableColumn> schema = GanttTableSchema.Default.Columns;
        for (int index = 0; index < schema.Count; index++)
        {
            if (string.Equals(schema[index].Name, headerName, StringComparison.Ordinal))
            {
                return index + 1; // 1-based
            }
        }

        throw new ArgumentException($"Unknown column '{headerName}'.", nameof(headerName));
    }
}
