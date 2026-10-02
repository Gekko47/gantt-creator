using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Office;
using Xunit.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// Live-Excel gate for the outline-group writer (R4.7D D5). Tagged
/// <c>[Trait("Category","OfficeIntegration")]</c> so <c>verify-quick.ps1</c> and
/// <c>verify.ps1</c> exclude it; run via <c>pwsh ./scripts/verify-office.ps1</c>.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why this suite exists.</b> The Refresh command failed for every user with
/// <c>COMException 0x800A03EC</c> — "Unable to set the OutlineLevel property of
/// the Range class" — raised from <c>WriteOutlineLevel</c>. The cause was that
/// the span was resolved through <c>Worksheet.Range[...]</c>, a range whose
/// <c>OutlineLevel</c> the host refuses to write; <c>Worksheet.Rows[...]</c>
/// returns the identical span shape and accepts the write. The contract tests
/// pin which accessor is used, but only a live host can show that the write
/// actually lands, so the regression is asserted here against real Excel.
/// </para>
/// <para>
/// The reported failing shape is reproduced deliberately: the sheet in the bug
/// report had a flat table with no <c>ParentId</c> anywhere, which means the plan
/// contains no groups and the reset pass writes level 1 across the entire body.
/// That is the exact write that threw, so the first test asserts it succeeds now.
/// </para>
/// </remarks>
public class OutlineGroupIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    private static string XllPath => OfficeFixtureTests.ResolvePackedXllPath();

    private static GanttRowId NewId() => GanttRowId.New();

    private static GanttEvent Event(int rowNumber, GanttRowId id, GanttRowId? parentId = null) =>
        new(
            rowNumber,
            id,
            NewId(),
            0,
            GanttEntityType.AsPlannedActivity,
            "Build frame",
            new DateOnly(2026, 1, 1),
            new DateOnly(2026, 12, 31),
            parentId,
            "AsPlannedActivity",
            null,
            null,
            null,
            true,
            null
        );

    /// <summary>
    /// Creates an initialised Gantt sheet with <paramref name="bodyRows"/> data rows.
    /// </summary>
    private static (Excel.Worksheet Sheet, Excel.ListObject Table) InitialisedGantt(
        OfficeFixture fixture,
        OfficeFixture.ComScope scope,
        int bodyRows
    )
    {
        fixture.CreateWorkbook();
        var initialiser = new ExcelWorkbookInitialiser(fixture.Excel);
        WorkbookInitialiseOutcome initialised = initialiser.Initialise();
        Assert.True(initialised.Succeeded, $"Initialise refused: {initialised.Refusal}");

        Excel.Sheets sheets = scope.Track(fixture.Excel.ActiveWorkbook!.Sheets);
        Excel.Worksheet sheet = scope.Track((Excel.Worksheet)sheets[GanttWorkbookContract.GanttSheetLabel]);
        Excel.ListObjects objects = scope.Track(sheet.ListObjects);
        Excel.ListObject table = scope.Track(objects[GanttTableSchema.TableName]);

        // The initialiser creates the table from a header row alone, so a body has to
        // be added before the writer has any rows to outline.
        Excel.ListRows rows = scope.Track(table.ListRows);
        for (var i = 0; i < bodyRows; i++)
        {
            scope.Track(rows.Add());
        }

        return (sheet, table);
    }

    /// <summary>Reads one body row's outline level from the host.</summary>
    private static int ReadLevel(OfficeFixture.ComScope scope, Excel.ListObject table, int bodyRow)
    {
        Excel.Range body = scope.Track(table.DataBodyRange!);
        Excel.Range row = scope.Track(body.Rows[bodyRow]);
        return Convert.ToInt32(row.OutlineLevel, CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// The reported failure: a flat table with no <c>ParentId</c> anywhere, so the
    /// plan has no groups and the reset pass writes level 1 across the whole body.
    /// This write raised <c>0x800A03EC</c> for every user; it must now succeed.
    /// </summary>
    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task A_flat_table_writes_level_one_across_the_body_without_a_com_error()
    {
        var fixture = new OfficeFixture();
        var scope = new OfficeFixture.ComScope();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            (_, Excel.ListObject table) = InitialisedGantt(fixture, scope, bodyRows: 5);

            var writer = new ExcelOutlineGroupWriter(fixture.Excel);
            List<GanttEvent> flat = [.. Enumerable.Range(1, 5).Select(i => Event(i, NewId()))];

            OutlineGroupOutcome outcome = writer.Apply(flat);

            _output.WriteLine(
                $"Excel build {fixture.Excel.Version} (PID {fixture.ProcessId}); "
                    + $"flat table: applied={outcome.GroupsApplied} reset={outcome.RowsUngrouped} refusal={outcome.Refusal}"
            );

            Assert.True(outcome.Succeeded, $"Flat-table refresh refused: {outcome.Refusal}");
            Assert.Equal(0, outcome.GroupsApplied);

            // The host's own answer, read back rather than assumed: every body row
            // must actually carry level 1 on the sheet.
            Excel.Range body = scope.Track(table.DataBodyRange!);
            for (var i = 1; i <= body.Rows.Count; i++)
            {
                Assert.Equal(ExcelOutlineGroupWriter.TopLevelOutlineLevel, ReadLevel(scope, table, i));
            }

            _output.WriteLine($"All {body.Rows.Count} body rows report OutlineLevel 1.");
        }
        finally
        {
            scope.Dispose();
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    /// <summary>
    /// A parent with contiguous children writes a real, collapsible level-2 group,
    /// and the parent row itself stays at level 1 so its collapse control survives.
    /// </summary>
    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task A_parent_with_children_produces_a_collapsible_group_and_leaves_the_parent_at_level_one()
    {
        var fixture = new OfficeFixture();
        var scope = new OfficeFixture.ComScope();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            (Excel.Worksheet sheet, Excel.ListObject table) = InitialisedGantt(fixture, scope, bodyRows: 4);

            GanttRowId parent = NewId();
            List<GanttEvent> hierarchy = [Event(1, parent), Event(2, NewId(), parent), Event(3, NewId(), parent), Event(4, NewId())];

            var writer = new ExcelOutlineGroupWriter(fixture.Excel);
            OutlineGroupOutcome outcome = writer.Apply(hierarchy);

            _output.WriteLine(
                $"Excel build {fixture.Excel.Version} (PID {fixture.ProcessId}); "
                    + $"grouped: applied={outcome.GroupsApplied} reset={outcome.RowsUngrouped} refusal={outcome.Refusal}"
            );

            Assert.True(outcome.Succeeded, $"Grouped refresh refused: {outcome.Refusal}");
            Assert.Equal(1, outcome.GroupsApplied);

            // Body row 1 is the parent, rows 2-3 its children, row 4 unrelated.
            Assert.Equal(ExcelOutlineGroupWriter.TopLevelOutlineLevel, ReadLevel(scope, table, 1));
            Assert.Equal(ExcelOutlineGroupWriter.ChildOutlineLevel, ReadLevel(scope, table, 2));
            Assert.Equal(ExcelOutlineGroupWriter.ChildOutlineLevel, ReadLevel(scope, table, 3));
            Assert.Equal(ExcelOutlineGroupWriter.TopLevelOutlineLevel, ReadLevel(scope, table, 4));

            // The group must be genuinely collapsible, not merely a written number:
            // collapsing hides the child rows.
            int firstBodyRow = scope.Track(table.DataBodyRange!).Row;
            Excel.Outline outline = scope.Track(sheet.Outline);
            outline.ShowLevels(1, 1);
            try
            {
                Excel.Range child = scope.Track(sheet.Rows[firstBodyRow + 1]);
                Assert.True(child.Hidden, "Child row did not hide when the group was collapsed.");
            }
            finally
            {
                outline.ShowLevels(2, 2);
            }

            _output.WriteLine("Group collapsed and expanded through the host's own outline controls.");
        }
        finally
        {
            scope.Dispose();
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    /// <summary>
    /// Refresh is idempotent for the outline: re-applying the same hierarchy leaves
    /// the same levels, and deleting the parent resets the formerly grouped rows
    /// instead of leaving a stale group behind.
    /// </summary>
    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Reapplying_the_same_hierarchy_is_stable_and_deleting_the_parent_resets_the_children()
    {
        var fixture = new OfficeFixture();
        var scope = new OfficeFixture.ComScope();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            (_, Excel.ListObject table) = InitialisedGantt(fixture, scope, bodyRows: 3);

            GanttRowId parent = NewId();
            List<GanttEvent> hierarchy = [Event(1, parent), Event(2, NewId(), parent), Event(3, NewId())];

            var writer = new ExcelOutlineGroupWriter(fixture.Excel);

            OutlineGroupOutcome first = writer.Apply(hierarchy);
            Assert.True(first.Succeeded, $"First refresh refused: {first.Refusal}");
            Assert.Equal(ExcelOutlineGroupWriter.ChildOutlineLevel, ReadLevel(scope, table, 2));

            OutlineGroupOutcome second = writer.Apply(hierarchy);
            Assert.True(second.Succeeded, $"Second refresh refused: {second.Refusal}");
            Assert.Equal(1, second.GroupsApplied);
            Assert.Equal(ExcelOutlineGroupWriter.ChildOutlineLevel, ReadLevel(scope, table, 2));

            // The parent is gone, so the child is promoted and no group remains.
            OutlineGroupOutcome afterDelete = writer.Apply([Event(1, NewId()), Event(2, NewId())]);
            _output.WriteLine(
                $"Excel build {fixture.Excel.Version} (PID {fixture.ProcessId}); "
                    + $"after delete: applied={afterDelete.GroupsApplied} reset={afterDelete.RowsUngrouped}"
            );

            Assert.True(afterDelete.Succeeded, $"Post-delete refresh refused: {afterDelete.Refusal}");
            Assert.Equal(0, afterDelete.GroupsApplied);
            Assert.Equal(ExcelOutlineGroupWriter.TopLevelOutlineLevel, ReadLevel(scope, table, 2));
        }
        finally
        {
            scope.Dispose();
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }
}
