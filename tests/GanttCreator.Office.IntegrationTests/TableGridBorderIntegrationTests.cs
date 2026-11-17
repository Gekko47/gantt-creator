using System.Globalization;
using GanttCreator.Core;
using Xunit.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// End-to-end gate for the table's internal cell grid: every visible column boundary must carry
/// a rule in the header row and in every body row, after Initialise and after the rows the
/// Add-row command creates.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why this exists.</b> Owner report 2026-11-15: "there are some table rows where the vertical
/// line between Start and Finish does not appear", later narrowed to "still missing on a clean
/// build and a newly initialised workbook, and it is only that one boundary". Three offline
/// theories were measured and refuted on a live host - insert format inheritance under both
/// CopyOrigins, the 2-D matrix value write, and the pastel warning fill. What remained was a
/// choice between two classes: the sheet's own state for that boundary (a direct border override,
/// merged cells, or the two cells belonging to different ranges) and a rendering artefact where
/// the border exists but is not painted.
/// </para>
/// <para>
/// <b>How it discriminates.</b> Each boundary is read four independent ways and the readings are
/// required to agree:
/// <c>Borders[xlEdgeLeft/xlEdgeRight].LineStyle</c> (direct formatting),
/// <c>Range(adjacentPair).Borders[xlInsideVertical]</c> (range-level direct formatting),
/// <c>DisplayFormat</c> (what the host would paint), and <c>MergeCells</c> (the "two cells are one
/// range" case). A boundary that fails only the <c>DisplayFormat</c> read is a painting artefact;
/// one that fails the direct reads, or whose cells report merged, is sheet state. The assertion
/// message names which reading disagreed for which row and boundary, so a failure is a diagnosis
/// rather than a bare "expected true".
/// </para>
/// </remarks>
public class TableGridBorderIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    private static string XllPath => OfficeFixtureTests.ResolvePackedXllPath();

    private static GanttRowId FixedId(char suffix) =>
        GanttRowId.Parse("G-0123456789abcdef0123456789abcde" + suffix);

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Every_visible_boundary_carries_a_rule_in_the_header_and_every_body_row()
    {
        OfficeFixture fixture = new() { SuppressLeakSignal = true };
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath), "RegisterXLL returned false.");

            Excel.Workbook workbook = fixture.CreateWorkbook();
            WorkbookInitialiseOutcome initialised = new ExcelWorkbookInitialiser(fixture.Excel).Initialise();
            Assert.True(initialised.Succeeded, $"Initialise refused: {initialised.Refusal}");
            OfficeFixture.DisableChartProtection(workbook);

            Excel.Worksheet sheet = (Excel.Worksheet)workbook.Sheets[GanttWorkbookContract.GanttSheetLabel];
            Excel.ListObject table = sheet.ListObjects[GanttTableSchema.TableName];

            // The owner's state: one row per Add-row command, after a fresh Initialise. The first
            // command reuses the initial blank body row, the rest are appended, so the matrix
            // covers both provenance paths in one sheet.
            ExcelGanttRowInserter inserter = new(fixture.Excel);
            for (int index = 0; index < 4; index++)
            {
                GanttRowInsertOutcome insert = inserter.Insert(
                    GanttEntityType.AsPlannedActivity,
                    () => FixedId((char)('1' + index)));
                Assert.True(insert.Succeeded, $"Insert {index} refused: {insert.Refusal}");
            }

            AssertEveryBoundaryHasRule(sheet, table, "after Initialise and four Add-row commands");

            // The owner's rows are PINK: they are invalid (no dates), so the refresh's validation
            // report fills the offending cells. The reporter fills Start and Finish per cell, and
            // also has a whole-row path. Apply both shapes here, because that filled state is the
            // one the report is about and it is the one state the earlier replays never had.
            ExcelRowWarningHighlighter highlighter = new(fixture.Excel);
            int bodyRows = table.ListRows.Count;
            int startOffset = ColumnOffset(table, "Start");
            int finishOffset = ColumnOffset(table, "Finish");
            for (int row = 1; row <= bodyRows; row++)
            {
                _ = highlighter.HighlightCell(row, startOffset, GanttValidationSeverity.Error);
                _ = highlighter.HighlightCell(row, finishOffset, GanttValidationSeverity.Error);
            }

            AssertEveryBoundaryHasRule(sheet, table, "after the validation report fills Start and Finish per cell");

            // And the whole-row path, which is the other shape the reporter uses.
            for (int row = 1; row <= bodyRows; row++)
            {
                _ = highlighter.HighlightRow(row, GanttValidationSeverity.Error);
            }

            AssertEveryBoundaryHasRule(sheet, table, "after the validation report fills whole rows");
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    /// <summary>
    /// Production order: the validation report fills the initial body row, and THEN an Add
    /// extends the table.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The sibling fact inserts four rows and only then fills, which is the one order the host
    /// tolerates. Production is the other way round — the coalesced refresh fills after every
    /// Add, so the next Add always extends a table whose <c>Start</c> and <c>Finish</c> cells
    /// are already pink.
    /// </para>
    /// <para>
    /// Measured on Excel 16.0 (<c>scripts/probe-insert-border-inheritance.ps1</c>, phases B
    /// and G): in that order the host drops the shared <c>Start|Finish</c> rule on the filled
    /// row <em>and</em> on the row the Add creates — through <c>Borders</c> and through
    /// <c>DisplayFormat</c> — while every other boundary on those rows keeps its rule. Those
    /// two cells are the only adjacent pair the report fills, which is why the owner reported
    /// exactly that boundary and no other. This fact fails without the insert-time repair in
    /// <c>ExcelGanttRowInserter.EnsureTableBorderColour</c> and passes with it.
    /// </para>
    /// </remarks>
    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Every_boundary_survives_the_reports_fill_followed_by_an_add()
    {
        OfficeFixture fixture = new() { SuppressLeakSignal = true };
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath), "RegisterXLL returned false.");

            Excel.Workbook workbook = fixture.CreateWorkbook();
            WorkbookInitialiseOutcome initialised = new ExcelWorkbookInitialiser(fixture.Excel).Initialise();
            Assert.True(initialised.Succeeded, $"Initialise refused: {initialised.Refusal}");
            OfficeFixture.DisableChartProtection(workbook);

            Excel.Worksheet sheet = (Excel.Worksheet)workbook.Sheets[GanttWorkbookContract.GanttSheetLabel];
            Excel.ListObject table = sheet.ListObjects[GanttTableSchema.TableName];

            // Add #1 puts a real body row on the sheet. The body is EMPTY straight after
            // Initialise, and HighlightCell reports Ok(0) for a row beyond the body's
            // current extent -- so filling before this insert would fill nothing and make
            // the rest of this fact vacuous.
            ExcelGanttRowInserter inserter = new(fixture.Excel);
            GanttRowInsertOutcome first = inserter.Insert(
                GanttEntityType.AsPlannedActivity,
                () => FixedId('1'));
            Assert.True(first.Succeeded, $"First insert refused: {first.Refusal}");
            Assert.Equal(1, table.ListRows.Count);

            // The state a Refresh leaves after that Add: the row has no dates, so the report
            // fills both of its date cells.
            ExcelRowWarningHighlighter highlighter = new(fixture.Excel);
            int startOffset = ColumnOffset(table, "Start");
            int finishOffset = ColumnOffset(table, "Finish");
            GanttValidationReportOutcome startFill =
                highlighter.HighlightCell(1, startOffset, GanttValidationSeverity.Error);
            GanttValidationReportOutcome finishFill =
                highlighter.HighlightCell(1, finishOffset, GanttValidationSeverity.Error);

            // The premise of this fact is the pink state, so it is asserted rather than
            // assumed, and asserted by the WRITE COUNT: a refused or out-of-extent fill is
            // "successful" with zero writes, which is the shape that let an earlier gate stay
            // green over a visible defect. 0x00C8C8FF is the report's pastel red
            // (RGB 255,200,200) as the host reports it.
            Assert.True(startFill.Succeeded, $"The Start fill was refused: {startFill.Refusal}");
            Assert.True(finishFill.Succeeded, $"The Finish fill was refused: {finishFill.Refusal}");
            Assert.Equal(1, startFill.NotesWritten);
            Assert.Equal(1, finishFill.NotesWritten);
            Excel.Range filledRow = sheet.Cells[table.Range.Row + 1, table.Range.Column];
            Excel.Range filledStart = (Excel.Range)filledRow.Offset[0, startOffset - 1];
            Excel.Range filledFinish = (Excel.Range)filledRow.Offset[0, finishOffset - 1];
            Assert.Equal(
                0x00C8C8FF,
                Convert.ToInt32(filledStart.Interior.Color, CultureInfo.InvariantCulture));
            Assert.Equal(
                0x00C8C8FF,
                Convert.ToInt32(filledFinish.Interior.Color, CultureInfo.InvariantCulture));

            // Then the user's next Add, which is what extends the table.
            GanttRowInsertOutcome second = inserter.Insert(
                GanttEntityType.AsPlannedActivity,
                () => FixedId('2'));
            Assert.True(second.Succeeded, $"Second insert refused: {second.Refusal}");

            AssertEveryBoundaryHasRule(sheet, table, "after the report's fill, then one more Add");
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    private static int ColumnOffset(Excel.ListObject table, string headerName)
    {
        for (int index = 1; index <= table.ListColumns.Count; index++)
        {
            if (string.Equals(table.ListColumns[index].Name, headerName, StringComparison.Ordinal))
            {
                return index;
            }
        }

        throw new InvalidOperationException($"No column '{headerName}'.");
    }

    private void AssertEveryBoundaryHasRule(Excel.Worksheet sheet, Excel.ListObject table, string phase)
    {
        _output.WriteLine("=== " + phase + " ===");
        Excel.Range tableRange = table.Range;
        int firstRow = tableRange.Row;
        int lastRow = firstRow + tableRange.Rows.Count - 1;
        int firstColumn = tableRange.Column;
        int lastColumn = firstColumn + tableRange.Columns.Count - 1;

        // The visible columns, left to right, with their worksheet column index. A hidden column
        // is skipped: its boundary is not on screen, so it cannot be reported as missing.
        List<(string Header, int Column)> visible = [];
        for (int offset = 0; offset < table.ListColumns.Count; offset++)
        {
            int column = firstColumn + offset;
            if (Convert.ToBoolean(sheet.Columns[column].Hidden, CultureInfo.InvariantCulture))
            {
                continue;
            }

            visible.Add((table.ListColumns[offset + 1].Name, column));
        }

        // Derived from the schema rather than restated: R5.21 added the sixth visible
        // column (Activity ID), and a literal here fails a correct table for the old count.
        int expectedVisible = GanttTableSchema.Default.Columns.Count(column => !column.IsHidden);
        Assert.Equal(expectedVisible, visible.Count);
        _output.WriteLine("visible columns: " + string.Join(", ", visible.Select(entry => entry.Header)));

        List<string> failures = [];
        for (int boundary = 0; boundary + 1 < visible.Count; boundary++)
        {
            int left = visible[boundary].Column;
            int right = visible[boundary + 1].Column;
            string name = visible[boundary].Header + "|" + visible[boundary + 1].Header;

            for (int row = firstRow; row <= lastRow; row++)
            {
                string reading = ReadBoundary(sheet, row, left, right, out bool merged);
                _output.WriteLine(string.Create(
                    CultureInfo.InvariantCulture,
                    $"row {row,3} {name,-22} {reading}"));

                if (merged)
                {
                    failures.Add(string.Create(
                        CultureInfo.InvariantCulture,
                        $"row {row} {name}: MERGED - the two cells are one range"));
                    continue;
                }

                if (!reading.Contains("present", StringComparison.Ordinal))
                {
                    failures.Add(string.Create(CultureInfo.InvariantCulture, $"row {row} {name}: {reading}"));
                }
            }
        }

        Assert.True(
            failures.Count == 0,
            "Missing internal grid rules:" + Environment.NewLine + string.Join(Environment.NewLine, failures));
    }

    /// <summary>
    /// Reads one column boundary four ways and reports them together, so a failure says which
    /// reading disagreed.
    /// </summary>
    private static string ReadBoundary(
        Excel.Worksheet sheet,
        int row,
        int leftColumn,
        int rightColumn,
        out bool merged)
    {
        Excel.Range leftCell = sheet.Cells[row, leftColumn];
        Excel.Range rightCell = sheet.Cells[row, rightColumn];

        merged = Convert.ToBoolean(leftCell.MergeCells, CultureInfo.InvariantCulture)
            || Convert.ToBoolean(rightCell.MergeCells, CultureInfo.InvariantCulture);

        Excel.Borders leftBorders = leftCell.Borders;
        Excel.Borders rightBorders = rightCell.Borders;
        int directRight = Convert.ToInt32(
            leftBorders[Excel.XlBordersIndex.xlEdgeRight].LineStyle, CultureInfo.InvariantCulture);
        int directLeft = Convert.ToInt32(
            rightBorders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle, CultureInfo.InvariantCulture);

        Excel.Range pair = sheet.Range[leftCell, rightCell];
        Excel.Borders pairBorders = pair.Borders;
        int pairInside = Convert.ToInt32(
            pairBorders[Excel.XlBordersIndex.xlInsideVertical].LineStyle, CultureInfo.InvariantCulture);

        Excel.DisplayFormat pairDisplay = pair.DisplayFormat;
        int displayed = Convert.ToInt32(
            pairDisplay.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle, CultureInfo.InvariantCulture);

        // "present" means the host would paint a rule: the displayed reading is the one the eye
        // sees, and it is allowed to be supplied by the style rather than by direct formatting.
        // The direct readings are reported when they are absent so the failure names the class.
        bool present = displayed == (int)Excel.XlLineStyle.xlContinuous;
        string direct = directRight == (int)Excel.XlLineStyle.xlContinuous
            && directLeft == (int)Excel.XlLineStyle.xlContinuous
            ? "direct=yes"
            : "direct=no";

        return string.Create(
            CultureInfo.InvariantCulture,
            $"pairDirect={pairInside,6} display={displayed,6} {direct} -> {(present ? "present" : "MISSING")}");
    }
}
