using GanttCreator.Core;
using Xunit.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>Live-Excel acceptance gate for the R2.8 add-row commands.</summary>
public class AddRowIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    private static string XllPath => OfficeFixtureTests.ResolvePackedXllPath();

    private static GanttRowId FixedId(char suffix) =>
        GanttRowId.Parse("G-0123456789abcdef0123456789abcde" + suffix);

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Insert_appends_all_three_command_rows_without_rendering()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath),
                $"Application.RegisterXLL returned false for '{XllPath}'.");

            Excel.Workbook workbook = fixture.CreateWorkbook();
            ExcelWorkbookInitialiser initialiser = new(fixture.Excel);
            WorkbookInitialiseOutcome initialised = initialiser.Initialise();
            Assert.True(initialised.Succeeded, $"Initialise refused: {initialised.Refusal}");

            var sheet = (Excel.Worksheet)workbook.Sheets[GanttWorkbookContract.GanttSheetLabel];
            Excel.ListObject table = sheet.ListObjects[GanttTableSchema.TableName];
            var shapesBefore = sheet.Shapes.Count;
            var inserter = new ExcelGanttRowInserter(fixture.Excel);

            GanttRowInsertOutcome activity = inserter.Insert(
                GanttEntityType.AsPlannedActivity,
                () => FixedId('1'));
            var activityIndex = RequiredBodyIndex(activity);
            table.ListRows[activityIndex].Range.Select();
            GanttRowInsertOutcome milestone = inserter.Insert(
                GanttEntityType.AsPlannedMilestone,
                () => FixedId('2'));
            var milestoneIndex = RequiredBodyIndex(milestone);
            table.ListRows[milestoneIndex].Range.Select();
            GanttRowInsertOutcome delineator = inserter.Insert(
                GanttEntityType.Delineator,
                () => FixedId('3'));

            Assert.True(activity.Succeeded);
            Assert.True(milestone.Succeeded);
            Assert.True(delineator.Succeeded);
            Assert.True(activity.BodyIndex > 0);
            Assert.True(milestone.BodyIndex > activity.BodyIndex);
            Assert.True(delineator.BodyIndex > milestone.BodyIndex);
            Assert.True(table.ListRows.Count >= delineator.BodyIndex);
            Assert.Equal(shapesBefore, sheet.Shapes.Count);
            Assert.Contains(
                workbook.Names.Cast<Excel.Name>(),
                name => string.Equals(name.Name, GanttWorkbookContract.TypeOptionsDefinedName, StringComparison.OrdinalIgnoreCase));
            Excel.Name typeOptions = workbook.Names.Item(GanttWorkbookContract.TypeOptionsDefinedName);
            Assert.Contains(GanttWorkbookContract.ConfigSheetName, typeOptions.RefersTo, StringComparison.Ordinal);
            Assert.Contains("$B$2:$B$17", typeOptions.RefersTo, StringComparison.Ordinal);
            Excel.ListColumn typeColumn = table.ListColumns["Type"];
            Assert.NotNull(typeColumn.DataBodyRange);
            Excel.Validation validation = typeColumn.DataBodyRange.Validation;
            Assert.Equal((int)Excel.XlDVType.xlValidateList, validation.Type);
            Assert.Equal($"={GanttWorkbookContract.TypeOptionsDefinedName}", validation.Formula1);
            Assert.True(validation.InCellDropdown);
            _output.WriteLine("R2.9 TypeOptions name and Type-column validation are present after add-row fallback.");

            AssertRow(table.ListRows[activity.BodyIndex], "As-Planned Activity", "AsPlannedActivity", FixedId('1').Value);
            AssertRow(table.ListRows[milestone.BodyIndex], "As-Planned Milestone", "AsPlannedMilestone", FixedId('2').Value);
            Assert.Equal(3, table.DataBodyRange.Rows.Count);
            Assert.Equal(4, table.Range.Rows.Count);
            AssertRow(table.ListRows[delineator.BodyIndex], "Delineator", "DefaultDelineator", FixedId('3').Value);

            _output.WriteLine("R2.8 appended activity, milestone, and delineator rows; no trailing blank row or shapes changed.");
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Insert_first_activity_reuses_the_initial_blank_row()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath),
                $"Application.RegisterXLL returned false for '{XllPath}'.");

            Excel.Workbook workbook = fixture.CreateWorkbook();
            Assert.True(new ExcelWorkbookInitialiser(fixture.Excel).Initialise().Succeeded);
            var sheet = (Excel.Worksheet)workbook.Sheets[GanttWorkbookContract.GanttSheetLabel];
            Excel.ListObject table = sheet.ListObjects[GanttTableSchema.TableName];

            Assert.Equal(0, table.ListRows.Count);
            Assert.Equal(2, table.Range.Rows.Count);
            Assert.Null(table.DataBodyRange);

            GanttRowInsertOutcome outcome = new ExcelGanttRowInserter(fixture.Excel).Insert(
                GanttEntityType.AsPlannedActivity,
                () => FixedId('0'));

            Assert.True(outcome.Succeeded, $"Insert refused: {outcome.Refusal}");
            Assert.Equal(1, outcome.BodyIndex);
            Assert.Equal(1, table.ListRows.Count);
            Assert.Equal(2, table.Range.Rows.Count);
            Assert.NotNull(table.DataBodyRange);
            Assert.Equal(1, table.DataBodyRange.Rows.Count);
            // Derived from the schema, not hard-coded. R4.7A added SiblingOrder
            // and the schema grew to 15 columns; a literal here silently went
            // stale and only the live Office gate caught it.
            Assert.Equal(GanttTableSchema.Default.Columns.Count, table.DataBodyRange.Columns.Count);
            AssertRow(table.ListRows[1], "As-Planned Activity", "AsPlannedActivity", FixedId('0').Value);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    /// <summary>
    /// ADR-0035 D3: an active cell INSIDE the table inserts immediately BELOW it, and
    /// the reserved bottom padding row still moves down.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This inverts <c>Every_insert_appends_and_shifts_no_following_row</c>, which
    /// asserted the opposite and was correct for the append-only rule ADR-0035 first
    /// landed. That rule was reverted: the owner recognised the mid-table shift as the
    /// chart tracking the sheet, not as a defect.
    /// </para>
    /// <para>
    /// The padding-row half is asserted in the SAME test body on purpose. It is the
    /// same event, and the COM-proxy leak ratchet counts forced kills per test body
    /// (a workbook left open keeps the host alive), so splitting one narrative across
    /// two bodies would add a kill to a ceiling the repository records as not
    /// raisable, for no extra coverage.
    /// </para>
    /// </remarks>
    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task An_in_table_selection_inserts_below_it_and_the_padding_row_still_moves_down()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath),
                $"Application.RegisterXLL returned false for '{XllPath}'.");

            Excel.Workbook workbook = fixture.CreateWorkbook();
            Assert.True(new ExcelWorkbookInitialiser(fixture.Excel).Initialise().Succeeded);

            // The workbook is NOT tracked: CreateWorkbook hands that proxy to the
            // fixture, which closes and releases it during teardown, and tracking it
            // here released the same RCW twice. Every proxy THIS test creates is
            // tracked, so this test stops adding to the COM-proxy leak signal.
            using var scope = new OfficeFixture.ComScope();
            var sheet = (Excel.Worksheet)scope.Track(workbook.Sheets[GanttWorkbookContract.GanttSheetLabel]);
            Excel.ListObject table = scope.Track(sheet.ListObjects[GanttTableSchema.TableName]);
            var inserter = new ExcelGanttRowInserter(fixture.Excel);

            Assert.True(inserter.Insert(
                GanttEntityType.AsPlannedActivity,
                () => FixedId('1')).Succeeded);
            Assert.True(inserter.Insert(
                GanttEntityType.AsPlannedMilestone,
                () => FixedId('2')).Succeeded);
            Assert.True(inserter.Insert(
                GanttEntityType.Delineator,
                () => FixedId('3')).Succeeded);

            Excel.Range body = scope.Track(table.DataBodyRange);
            int paddingRowBefore = body.Row + body.Rows.Count;

            // Select the FIRST body row: the new activity belongs immediately below it,
            // NOT at the top and NOT appended.
            scope.Track(table.ListRows[1].Range).Select();
            GanttRowInsertOutcome inserted = inserter.Insert(
                GanttEntityType.AsPlannedActivity,
                () => FixedId('4'));

            Assert.True(inserted.Succeeded, $"Insert refused: {inserted.Refusal}");
            Assert.Equal(2, inserted.BodyIndex);
            Assert.Equal(4, table.DataBodyRange.Rows.Count);

            // The load-bearing ordering: the new row sits BELOW the selection, and the
            // rows it displaced kept their identities and moved down.
            AssertId(table.ListRows[1], FixedId('1').Value);
            AssertId(table.ListRows[2], FixedId('4').Value);
            AssertId(table.ListRows[3], FixedId('2').Value);
            AssertId(table.ListRows[4], FixedId('3').Value);

            // ---- The reserved bottom padding row (ADR-0035 D2) ----
            //
            // ListRows.Add(position) inserts a REAL worksheet row, so everything
            // below shifts: the padding row must have moved down by one row. This is
            // the behaviour the append-only rule lost, because ListRows.Add() with no
            // position CLAIMS the padding row instead of pushing it.
            Excel.Range bodyAfter = scope.Track(table.DataBodyRange);
            int lastBodyRow = bodyAfter.Row + bodyAfter.Rows.Count - 1;
            int tableLastRow = table.Range.Row + table.Range.Rows.Count - 1;
            int paddingRowIndex = lastBodyRow + 1;

            // The margin row is genuinely OUTSIDE the table. This is the load-bearing
            // check, and it is stated against the TABLE'S OWN RANGE rather than
            // against lastBodyRow -- an earlier version compared the derived index
            // with itself and was a tautology that could never fail.
            Assert.True(
                paddingRowIndex > tableLastRow,
                $"The reserved padding row ({paddingRowIndex}) must sit below the table's last row ({tableLastRow}).");

            // And it MOVED DOWN rather than being absorbed. Before ADR-0035 the
            // append claimed this row and turned it into a body row; a positional
            // insert shifts it, and the new row below the table restores the same for
            // the append branch.
            Assert.True(
                paddingRowIndex > paddingRowBefore,
                $"The padding row must move down, not be absorbed (was {paddingRowBefore}, margin is now {paddingRowIndex}).");

            Excel.Range paddingRow = scope.Track(sheet.Rows[paddingRowIndex]);
            object? paddingValue = paddingRow.Value2;

            // A whole-row Range.Value2 is ALWAYS a 2-D SAFEARRAY, never a scalar, so
            // every cell has to be inspected; a scalar check misreads a populated row.
            bool paddingIsEmpty = paddingValue is object[,] cells
                ? cells.Cast<object?>().All(static value =>
                    value is null or DBNull || (value is string text && text.Length == 0))
                : paddingValue is null
                    or DBNull
                    || (paddingValue is string single && single.Length == 0);

            Assert.True(
                paddingIsEmpty,
                $"The reserved padding row must be empty, but reported '{paddingValue}'.");
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    /// <summary>
    /// A protected worksheet refuses the insert and the table is left untouched.
    /// </summary>
    /// <remarks>
    /// Unrelated to ADR-0035. The live coverage for the reserved bottom padding row
    /// lives in <c>Every_insert_appends_and_shifts_no_following_row</c>, which is
    /// where the append that used to consume the margin row is performed.
    /// </remarks>
    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Insert_refuses_a_protected_sheet_without_adding_a_row()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath),
                $"Application.RegisterXLL returned false for '{XllPath}'.");

            Excel.Workbook workbook = fixture.CreateWorkbook();
            WorkbookInitialiseOutcome initialised = new ExcelWorkbookInitialiser(fixture.Excel).Initialise();
            Assert.True(initialised.Succeeded, $"Initialise refused: {initialised.Refusal}");

            var sheet = (Excel.Worksheet)workbook.Sheets[GanttWorkbookContract.GanttSheetLabel];
            Excel.ListObject table = sheet.ListObjects[GanttTableSchema.TableName];
            var rowsBefore = table.ListRows.Count;
            sheet.Protect(Type.Missing, true, Type.Missing, true, true, false, false);

            GanttRowInsertOutcome outcome = new ExcelGanttRowInserter(fixture.Excel).Insert(
                GanttEntityType.AsPlannedActivity,
                () => FixedId('4'));

            Assert.False(outcome.Succeeded);
            Assert.Equal(GanttRowInsertRefusalReason.TargetProtected, outcome.Refusal);
            Assert.Equal(rowsBefore, table.ListRows.Count);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    private static int RequiredBodyIndex(GanttRowInsertOutcome outcome)
    {
        Assert.True(outcome.Succeeded, $"Insert refused: {outcome.Refusal}");
        _ = Assert.NotNull(outcome.BodyIndex);
        return outcome.BodyIndex.Value;
    }

    private static void AssertId(Excel.ListRow row, string id)
    {
        Assert.Equal(id, Convert.ToString(
            row.Range.Cells[1, 1].Value2,
            System.Globalization.CultureInfo.InvariantCulture));
    }

    private static void AssertRow(Excel.ListRow row, string type, string styleKey, string id)
    {
        Excel.Range range = row.Range;

        // Resolved by COLUMN NAME, never by a literal index. R4.7A inserted
        // SiblingOrder after ParentId, which moved StyleKey from index 9 to 10
        // and left this helper reading a blank cell: the assertion compared the
        // expected style key against null and passed for the wrong reason on
        // every earlier run. R4.7C inserted Duration after Finish, moving it to 11,
        // and the cross-check below caught that rather than letting it pass.
        // Excel's 1-based Cells indices are offset by the header row, hence the +1.
        Assert.Equal(id, ColumnValue(range, "Id", 1));
        Assert.Equal(type, ColumnValue(range, "Type", 4));
        Assert.Equal(styleKey, ColumnValue(range, "StyleKey", 11));
    }

    /// <summary>
    /// Reads one cell by the column's position in the workbook schema, so a future
    /// column insertion cannot silently shift what this helper reads.
    /// </summary>
    /// <param name="rowRange">The row's range.</param>
    /// <param name="columnName">The schema column name.</param>
    /// <param name="expectedIndex">The 1-based Excel index the schema implies, asserted as a cross-check.</param>
    /// <returns>The cell's text.</returns>
    private static string ColumnValue(Excel.Range rowRange, string columnName, int expectedIndex)
    {
        var index = GanttTableSchema.Default.Columns
            .Select((column, position) => (column.Name, Position: position))
            .Single(entry => string.Equals(entry.Name, columnName, StringComparison.Ordinal))
            .Position;

        // The cross-check is the point: a mismatch here means the workbook and the
        // schema have diverged, which is a defect rather than a stale literal.
        Assert.Equal(expectedIndex, index + 1);

        return Convert.ToString(
            rowRange.Cells[1, index + 1].Value2,
            System.Globalization.CultureInfo.InvariantCulture);
    }
}
