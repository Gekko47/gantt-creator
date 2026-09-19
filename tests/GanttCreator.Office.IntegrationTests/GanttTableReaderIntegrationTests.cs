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

            // Locate the table on the now-labelled sheet.
            Excel.Worksheet ganttSheet = (Excel.Worksheet)workbook.Sheets[GanttWorkbookContract.GanttSheetLabel];
            Excel.ListObject table = ganttSheet.ListObjects[GanttTableSchema.TableName];
            Excel.Range body = table.DataBodyRange;

            // Two rows: a span + a milestone (Start-only geometry is R2.5; here
            // we only prove value reading). Dates as OLE Automation serials so
            // they are locale-invariant regardless of NumberFormat.
            const double startSerial = 44927.0; // 2023-01-01
            const double finishSerial = 44931.0; // 2023-01-05
            const double milestoneSerial = 44932.0; // 2023-01-06

            // Body rows start at row 1 (row 0 = header).
            SetBodyCell(body, 1, "Id", "G-span");
            SetBodyCell(body, 1, "LaneId", "L-1");
            SetBodyCell(body, 1, "StackIndex", 1.0);
            SetBodyCell(body, 1, "Type", "As-Planned Activity");
            SetBodyCell(body, 1, "Start", startSerial);
            SetBodyCell(body, 1, "Finish", finishSerial);
            SetBodyCell(body, 1, "Visible", true);

            SetBodyCell(body, 2, "Id", "G-milestone");
            SetBodyCell(body, 2, "LaneId", "L-1");
            SetBodyCell(body, 2, "StackIndex", 2.0);
            SetBodyCell(body, 2, "Type", "Milestone");
            SetBodyCell(body, 2, "Start", milestoneSerial);
            SetBodyCell(body, 2, "Finish", "   ");
            SetBodyCell(body, 2, "Visible", "TRUE");

            _output.WriteLine($"Excel build: {fixture.Excel.Version} (PID {fixture.ProcessId})");

            // Format A: GB DD/MM/YYYY.
            body.NumberFormat = "dd/MM/yyyy";
            GanttTableReadOutcome gb = new ExcelGanttTableReader(fixture.Excel).Read();
            AssertReadCorrect(gb);

            // Format B: US MM/DD/YYYY. Same serials -> same DTOs.
            body.NumberFormat = "MM/dd/yyyy";
            GanttTableReadOutcome us = new ExcelGanttTableReader(fixture.Excel).Read();
            AssertReadCorrect(us);

            // The central invariant: display format did not change the DTOs.
            Assert.Equal(gb.Rows.Count, us.Rows.Count);
            for (int index = 0; index < gb.Rows.Count; index++)
            {
                GanttRowDto? left = gb.Rows[index];
                GanttRowDto? right = us.Rows[index];
                Assert.NotNull(left);
                Assert.NotNull(right);
                Assert.Equal(left.RowNumber, right.RowNumber);
                Assert.Equal(left.Id, right.Id);
                Assert.Equal(left.TypeText, right.TypeText);
                Assert.Equal(left.Start, right.Start);
                Assert.Equal(left.Finish, right.Finish);
                Assert.Equal(left.StackIndex, right.StackIndex);
                Assert.Equal(left.Visible, right.Visible);
            }
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
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

    private static void AssertReadCorrect(GanttTableReadOutcome outcome)
    {
        Assert.True(outcome.Succeeded, $"Read refused: {outcome.Refusal}");
        Assert.Null(outcome.Refusal);
        Assert.Equal(2, outcome.Rows.Count);

        Assert.Equal(1, outcome.Rows[0].RowNumber);
        Assert.Equal("G-span", outcome.Rows[0].Id);
        Assert.Equal("L-1", outcome.Rows[0].LaneId);
        Assert.Equal(1, outcome.Rows[0].StackIndex);
        Assert.Equal("As-Planned Activity", outcome.Rows[0].TypeText);
        Assert.Equal(new DateOnly(2023, 1, 1), outcome.Rows[0].Start);
        Assert.Equal(new DateOnly(2023, 1, 5), outcome.Rows[0].Finish);
        Assert.True(outcome.Rows[0].Visible);

        Assert.Equal(2, outcome.Rows[1].RowNumber);
        Assert.Equal("G-milestone", outcome.Rows[1].Id);
        Assert.Equal("L-1", outcome.Rows[1].LaneId);
        Assert.Equal(2, outcome.Rows[1].StackIndex);
        Assert.Equal("Milestone", outcome.Rows[1].TypeText);
        Assert.Equal(new DateOnly(2023, 1, 6), outcome.Rows[1].Start);
        Assert.Null(outcome.Rows[1].Finish); // blank
        Assert.True(outcome.Rows[1].Visible); // "TRUE"
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