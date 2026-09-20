using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;
using GanttCreator.Core;
using GanttCreator.Office;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// Live-Excel gate for the table reader (work item R2.4). Tagged
/// <c>[Trait("Category","OfficeIntegration")]</c> so <c>verify-quick.ps1</c>
/// and <c>verify.ps1</c> exclude them; run via
/// <c>pwsh ./scripts/verify-office.ps1</c> on the self-hosted runner (ADR-0001).
/// </summary>
/// <remarks>
/// Drives through the <see cref="IGanttTableReader"/> port, not the AddIn
/// command layer: this project references Core and Office but not AddIn. The
/// contract assertions mirror <see cref="Office.ContractTests.GanttTableReaderTests"/>
/// (Moq PIA, no live Office) so a drift between the port and the live path is
/// caught. The core assertion pins the R2.4 invariant: identical <c>Value2</c>
/// serials must produce identical <see cref="GanttRowDto"/>s under both supported
/// locale display formats (GB <c>DD/MM/YYYY</c> and US <c>MM/DD/YYYY</c>), which
/// proves the reader parses cell values rather than display text (Checklist F).
/// </remarks>
public class GanttTableReaderIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    private static string XllPath => OfficeFixtureTests.ResolvePackedXllPath();

    /// <summary>
    /// The adopt path produces one visible sheet with <c>tblGanttData</c>.
    /// Write real serial dates into the body, then read the table twice —
    /// once with GB display formatting and once with US — and assert the
    /// DTOs (including <see cref="DateOnly"/>) are identical. This is the
    /// Required live-Excel proof that the converter never parses <c>.Text</c>.
    /// </summary>
    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Read_returns_locale_invariant_dtos_under_gb_and_us_display_formats()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath),
                $"Application.RegisterXLL returned false for '{XllPath}'.");

            Excel.Workbook workbook = fixture.CreateWorkbook();

            // Adopt the blank sheet into the supported contract state.
            var initialiser = new ExcelWorkbookInitialiser(fixture.Excel);
            WorkbookInitialiseOutcome outcome = initialiser.Initialise();
            Assert.True(outcome.Succeeded, $"Initialise refused: {outcome.Refusal}");

            // Locate the table on the now-labelled sheet and clear any
            // pre-existing body rows so the geometry read below is the
            // geometry this test creates. Excel's ListObject always keeps a row
            // in its range and this build can leave the body one row wider than
            // ListRows.Count after DataBodyRange.Delete, so each write is
            // targeted through the ListRow object ListRows.Add returned and
            // every assertion is computed from the body offset that row
            // actually occupies.
            Excel.Worksheet ganttSheet = (Excel.Worksheet)workbook.Sheets[GanttWorkbookContract.GanttSheetLabel];
            Excel.ListObject table = ganttSheet.ListObjects[GanttTableSchema.TableName];
            ClearTableBody(table);
            Assert.Equal(0, table.ListRows.Count);

            // Add the two body rows: a span + a milestone (Start-only geometry
            // is R2.5; here we only prove value reading). Dates as OLE
            // Automation serials are locale-invariant regardless of NumberFormat.
            Excel.ListRow spanRow = table.ListRows.Add();
            Excel.ListRow milestoneRow = table.ListRows.Add();

            Excel.Range body = table.DataBodyRange;
            Assert.NotNull(body);

            Excel.Range bodyRows = body.Rows;
            Excel.Range spanRange = spanRow.Range;
            Excel.Range milestoneRange = milestoneRow.Range;
            int bodyRowCount = bodyRows.Count;
            int spanOffset = spanRange.Row - body.Row + 1;
            int milestoneOffset = milestoneRange.Row - body.Row + 1;
            _output.WriteLine(
                $"Excel build {fixture.Excel.Version} (PID {fixture.ProcessId}); " +
                $"ListRows.Count={table.ListRows.Count}; body={body.Address}; " +
                $"body.Rows.Count={bodyRowCount}; spanOffset={spanOffset}; " +
                $"milestoneOffset={milestoneOffset}");

            // Observed live on Excel 16.0 (2026-09-19): after
            // DataBodyRange.Delete cleared a one-row body, two ListRows.Add
            // calls produced ListRows.Count=3 with body A2:N4 — the added rows
            // landed at body offsets 1 and 3 and a blank body row sat between
            // them. The reader must return one DTO per body row, so the offsets
            // that the returned ListRow objects actually occupy locate the
            // written rows; adjacency is not assumed.
            Assert.True(
                milestoneOffset > spanOffset,
                $"Expected the milestone row below the span row; " +
                $"spanOffset={spanOffset}, milestoneOffset={milestoneOffset}.");

            const double startSerial = 44927.0; // 2023-01-01
            const double finishSerial = 44931.0; // 2023-01-05
            const double milestoneSerial = 44932.0; // 2023-01-06

            SetBodyCell(body, spanOffset, "Id", "G-span");
            SetBodyCell(body, spanOffset, "LaneId", "L-1");
            SetBodyCell(body, spanOffset, "StackIndex", 1.0);
            SetBodyCell(body, spanOffset, "Type", "As-Planned Activity");
            SetBodyCell(body, spanOffset, "Start", startSerial);
            SetBodyCell(body, spanOffset, "Finish", finishSerial);
            SetBodyCell(body, spanOffset, "Visible", true);

            SetBodyCell(body, milestoneOffset, "Id", "G-milestone");
            SetBodyCell(body, milestoneOffset, "LaneId", "L-1");
            SetBodyCell(body, milestoneOffset, "StackIndex", 2.0);
            SetBodyCell(body, milestoneOffset, "Type", "Milestone");
            SetBodyCell(body, milestoneOffset, "Start", milestoneSerial);
            SetBodyCell(body, milestoneOffset, "Finish", "   ");
            SetBodyCell(body, milestoneOffset, "Visible", "TRUE");

            // Format A: GB DD/MM/YYYY.
            body.NumberFormat = "dd/MM/yyyy";
            string gbFinishText = GetBodyCellText(body, spanOffset, "Finish");
            GanttTableReadOutcome gb = new ExcelGanttTableReader(fixture.Excel).Read();
            AssertReadCorrect(gb, spanOffset, milestoneOffset, bodyRowCount);

            // Format B: US MM/DD/YYYY. The same serials now display
            // differently, so identical DTOs prove the reader read the Value2
            // serials and never parsed display text.
            body.NumberFormat = "MM/dd/yyyy";
            string usFinishText = GetBodyCellText(body, spanOffset, "Finish");
            GanttTableReadOutcome us = new ExcelGanttTableReader(fixture.Excel).Read();
            AssertReadCorrect(us, spanOffset, milestoneOffset, bodyRowCount);

            _output.WriteLine($"2023-01-05 display: GB '{gbFinishText}' vs US '{usFinishText}'");
            Assert.NotEqual(gbFinishText, usFinishText);

            // The central invariant: the display format changed, the DTOs did not.
            Assert.Equal(gb.Rows.Count, us.Rows.Count);
            for (int index = 0; index < gb.Rows.Count; index++)
            {
                Assert.Equal(gb.Rows[index], us.Rows[index]);
            }
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    /// <summary>
    /// Removes every body row from the table. Uses the atomic
    /// DataBodyRange.Delete path when it exists and a per-row fallback when
    /// DataBodyRange is null (a COM quirk seen on this build where
    /// ListRows.Count was 1 but DataBodyRange was null right after
    /// Initialise). The fallback re-reads .Count each pass.
    /// </summary>
    private static void ClearTableBody(Excel.ListObject table)
    {
        if (table.DataBodyRange is not null)
        {
            table.DataBodyRange.Delete();
            return;
        }

        // DataBodyRange null but ListRows may still hold rows on this build.
        // Delete from the end; re-read .Count each pass.
        while (true)
        {
            int count = table.ListRows.Count;
            if (count == 0) { break; }
            Excel.ListRow? last = table.ListRows[count];
            if (last is null) { break; }
            last.Delete();
        }
    }

    /// <summary>
    /// A 1904 workbook is explicitly refused (ADR-0006). The refusal is typed
    /// and mutates nothing: the existing 1900 sheet remains read.
    /// </summary>
    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Read_refuses_a_1904_workbook()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath),
                $"Application.RegisterXLL returned false for '{XllPath}'.");

            Excel.Workbook workbook = fixture.CreateWorkbook();
            workbook.Date1904 = true;
            // Activate so the reader can reach ActiveWorkbook.
            workbook.Activate();

            GanttTableReadOutcome outcome = new ExcelGanttTableReader(fixture.Excel).Read();

            Assert.False(outcome.Succeeded);
            Assert.Equal(GanttTableReadRefusalReason.DateSystemUnsupported, outcome.Refusal);
            Assert.Empty(outcome.Rows);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    /// <summary>
    /// Asserts the DTOs handed back for one read: one per body row, the two
    /// rows this test populated at their body offsets, and every other row
    /// wholly blank (nothing invented, nothing duplicated).
    /// </summary>
    private static void AssertReadCorrect(
        GanttTableReadOutcome outcome, int spanOffset, int milestoneOffset, int expectedRowCount)
    {
        Assert.True(outcome.Succeeded, $"Read refused: {outcome.Refusal}");
        Assert.Null(outcome.Refusal);
        Assert.Equal(expectedRowCount, outcome.Rows.Count);

        GanttRowDto span = outcome.Rows[spanOffset - 1];
        Assert.Equal(spanOffset, span.RowNumber);
        Assert.Equal("G-span", span.Id);
        Assert.Equal("L-1", span.LaneId);
        Assert.Equal(1, span.StackIndex);
        Assert.Equal("As-Planned Activity", span.TypeText);
        Assert.Equal(new DateOnly(2023, 1, 1), span.Start);
        Assert.Equal(new DateOnly(2023, 1, 5), span.Finish);
        Assert.True(span.Visible);

        GanttRowDto milestone = outcome.Rows[milestoneOffset - 1];
        Assert.Equal(milestoneOffset, milestone.RowNumber);
        Assert.Equal("G-milestone", milestone.Id);
        Assert.Equal("L-1", milestone.LaneId);
        Assert.Equal(2, milestone.StackIndex);
        Assert.Equal("Milestone", milestone.TypeText);
        Assert.Equal(new DateOnly(2023, 1, 6), milestone.Start);
        Assert.Null(milestone.Finish); // written as whitespace -> blank
        Assert.True(milestone.Visible); // "TRUE"

        for (int index = 0; index < outcome.Rows.Count; index++)
        {
            if (index == spanOffset - 1 || index == milestoneOffset - 1)
            {
                continue;
            }

            AssertBlank(outcome.Rows[index]);
        }
    }

    /// <summary>
    /// Asserts a DTO carries no value at all, so an untouched body row can
    /// never be mistaken for data the reader invented.
    /// </summary>
    private static void AssertBlank(GanttRowDto row)
    {
        Assert.Null(row.Id);
        Assert.Null(row.LaneId);
        Assert.Null(row.StackIndex);
        Assert.Null(row.TypeText);
        Assert.Null(row.Description);
        Assert.Null(row.Start);
        Assert.Null(row.Finish);
        Assert.Null(row.ParentId);
        Assert.Null(row.StyleKey);
        Assert.Null(row.LabelPositionText);
        Assert.Null(row.FillColourText);
        Assert.Null(row.StrokeColourText);
        Assert.Null(row.Visible);
        Assert.Null(row.SortOrder);
    }

    /// <summary>
    /// Writes a value into the body cell of the given 1-based body row whose
    /// header is <paramref name="headerName"/>. Column position is resolved
    /// from the Default schema column order (Ordinal).
    /// </summary>
    private static void SetBodyCell(Excel.Range body, int bodyRowNumber, string headerName, object value)
    {
        Excel.Range cell = body.Cells[bodyRowNumber, GetColumnIndex(headerName)];
        cell.Value2 = value;
    }

    /// <summary>
    /// Reads the display text of the body cell in the given 1-based body row
    /// whose header is <paramref name="headerName"/>. Used to prove the two
    /// locale display formats really render the same serial differently.
    /// </summary>
    private static string GetBodyCellText(Excel.Range body, int bodyRowNumber, string headerName)
    {
        Excel.Range cell = body.Cells[bodyRowNumber, GetColumnIndex(headerName)];
        return cell.Text;
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