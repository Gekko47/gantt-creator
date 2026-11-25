using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// R6.5 Office gate: the staging lifecycle leaves zero temporary shapes and restores the
/// application state on the success path and after forced failures, and the group-first
/// cleanup order is measured on the host rather than assumed.
/// </summary>
/// <remarks>
/// <para>
/// The harness drives the production chain exactly as R6.3/R6.4's gates do: initialise a
/// real table, write real rows, read and validate them, measure the real grid, read the
/// workbook's settings and styles, build the EXPORT request, resolve the bounds, then run
/// <see cref="StagingLifecycle.Run"/> with the real <see cref="ExcelShapeWriter"/> (staging
/// scheme), the real <see cref="ExcelShapeGrouper"/> and the real
/// <see cref="ExcelApplicationStateScope"/>.
/// </para>
/// <para>
/// <b>D2's deletion-order probe.</b> The success body wraps the writer in a recording
/// decorator and records what each cleanup delete answered. The work item defers the
/// group-first deletion question to the Tier-A upgrade; this body records the measured
/// answer (whether the host removed the children with the group) rather than asserting an
/// assumption. Both observed outcomes are accepted as long as no delete reports any other
/// refusal and zero residue follows.
/// </para>
/// <para>
/// <b>Forced failures.</b> The second body throws from the Nth create through a
/// faithfully-forwarding decorator (so the composer's birth-list cleanup is what runs),
/// then throws from the grouping port (so the lifecycle's registry cleanup is what runs).
/// Both assert zero staged residue, an intact live chart and restored state.
/// </para>
/// </remarks>
[Trait("Category", "OfficeSmoke")]
public sealed class StagingLifecycleIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task The_success_path_leaves_zero_residue_and_restores_state()
    {
        OfficeFixture fixture = new();

        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            using OfficeFixture.ComScope scope = new();
            (Excel.Worksheet sheet, Excel.Sheets sheets, ExportBoundsOutcome resolution, Excel.Shape sentinel, Excel.Range witness, object? witnessBefore, int sheetCountBefore) = Prepare(fixture, scope);

            ExcelShapeWriter stagingWriter = new(
                fixture.Excel,
                ownershipScheme: ShapeOwnershipScheme.Staging);
            ExcelShapeWriter liveWriter = new(fixture.Excel);
            RecordingShapePort recording = new(stagingWriter);
            Assert.True(liveWriter.Create(LiveProbe("live:probe-1")).Succeeded);
            Assert.True(liveWriter.Create(LiveProbe("live:probe-2")).Succeeded);

            (int scrollRow, int scrollColumn, double zoom, string selection) = CaptureView(fixture, scope);
            bool copyInvoked = false;

            // The lifecycle disposes the scope inside its finally; `using` is the
            // analyzer-satisfying double dispose, which ExcelApplicationStateScope
            // guards with its own `_disposed` flag.
            using ExcelApplicationStateScope state = new(fixture.Excel);
            StagingLifecycleOutcome outcome = StagingLifecycle.Run(new(
                state,
                recording,
                new ExcelShapeGrouper(fixture.Excel),
                resolution,
                Copy: _ =>
                {
                    copyInvoked = true;
                    return StagingCopyResult.Success;
                }));

            Assert.True(outcome.Succeeded, $"Lifecycle refused: {outcome.Refusal}");
            Assert.True(copyInvoked, "The copy seam was not invoked.");
            Assert.NotNull(outcome.Group);
            Assert.True(outcome.Group.Succeeded, $"Grouping refused: {outcome.Group.Refusal}");
            Assert.Empty(outcome.Residue);

            // Zero temporary shapes, and the live chart and sentinels survive.
            Assert.Empty(stagingWriter.ListOwned());
            Assert.Equal(0, CountStaged(scope, sheet));
            Assert.Equal(["live:probe-1", "live:probe-2"], liveWriter.ListOwned());
            AssertViewRestored(
                fixture,
                scope,
                sheets,
                sentinel,
                witness,
                witnessBefore,
                sheetCountBefore,
                scrollRow,
                scrollColumn,
                zoom,
                selection);

            // D2 probe: the group delete is first and succeeded. The host removed the
            // children with the group (measured 2026-11-19), so the lifecycle skips the
            // child loop after one ownership read; a host that kept any child would show
            // its attempt here and it must be Succeeded or NotFound, never another refusal.
            Assert.NotEmpty(recording.Deletes);
            Assert.Equal(ShapeStagingTag.GroupPrimitiveId, recording.Deletes[0].Id);
            Assert.True(
                recording.Deletes[0].Outcome.Succeeded,
                $"Group delete refused: {recording.Deletes[0].Outcome.Refusal}");

            bool childrenRemovedWithGroup = recording.Deletes.Count == 1;
            foreach ((string id, ShapeWriteOutcome delete) in recording.Deletes.Skip(1))
            {
                Assert.True(
                    delete.Succeeded || delete.Refusal == ShapeWriteRefusal.NotFound,
                    $"Child delete '{id}' refused: {delete.Refusal}");
                childrenRemovedWithGroup &= delete.Refusal == ShapeWriteRefusal.NotFound;
            }

            Assert.True(
                recording.Deletes.Count == 1
                    || recording.Deletes.Count == outcome.Group.ChildOrder.Count + 1,
                "Cleanup must either skip the children the group removed or attempt every child; "
                    + $"observed {recording.Deletes.Count} deletes for {outcome.Group.ChildOrder.Count} children.");

            _output.WriteLine(string.Create(
                CultureInfo.InvariantCulture,
                $"Excel build {fixture.Excel.Version}; staged={outcome.Group.ChildOrder.Count}"
                + $" groupDelete=ok childrenRemovedWithGroup={childrenRemovedWithGroup}"
                + $" deletes={recording.Deletes.Count} residue={outcome.Residue.Count}"));

            CleanupLive(liveWriter, sentinel);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Forced_mid_composition_and_grouping_failures_leave_zero_residue_and_restore_state()
    {
        OfficeFixture fixture = new();

        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            using OfficeFixture.ComScope scope = new();
            (Excel.Worksheet sheet, Excel.Sheets sheets, ExportBoundsOutcome resolution, Excel.Shape sentinel, Excel.Range witness, object? witnessBefore, int sheetCountBefore) = Prepare(fixture, scope);

            ExcelShapeWriter stagingWriter = new(
                fixture.Excel,
                ownershipScheme: ShapeOwnershipScheme.Staging);
            ExcelShapeWriter liveWriter = new(fixture.Excel);
            Assert.True(liveWriter.Create(LiveProbe("live:probe-1")).Succeeded);
            Assert.True(liveWriter.Create(LiveProbe("live:probe-2")).Succeeded);

            // Failure 1: the third create throws. The composer's birth list is the only
            // registry that sees the first two, so this proves its exception cleanup.
            (int scrollRow, int scrollColumn, double zoom, string selection) = CaptureView(fixture, scope);
            ThrowingCreatePort throwingCreate = new(stagingWriter, failAt: 3);
            using ExcelApplicationStateScope createState = new(fixture.Excel);

            InvalidOperationException createThrown = Assert.Throws<InvalidOperationException>(() =>
                StagingLifecycle.Run(new(
                    createState,
                    throwingCreate,
                    new ExcelShapeGrouper(fixture.Excel),
                    resolution)));
            Assert.Equal("create failed", createThrown.Message);
            Assert.Equal(0, CountStaged(scope, sheet));
            Assert.Empty(stagingWriter.ListOwned());
            Assert.Equal(["live:probe-1", "live:probe-2"], liveWriter.ListOwned());
            AssertViewRestored(
                fixture,
                scope,
                sheets,
                sentinel,
                witness,
                witnessBefore,
                sheetCountBefore,
                scrollRow,
                scrollColumn,
                zoom,
                selection);

            // Failure 2: the grouping port throws after the full staged set exists. The
            // lifecycle's registry is populated, so its finally is what cleans up.
            (scrollRow, scrollColumn, zoom, selection) = CaptureView(fixture, scope);
            ThrowingGroupingPort throwingGrouping = new(new ExcelShapeGrouper(fixture.Excel));
            using ExcelApplicationStateScope groupState = new(fixture.Excel);

            InvalidOperationException groupThrown = Assert.Throws<InvalidOperationException>(() =>
                StagingLifecycle.Run(new(
                    groupState,
                    stagingWriter,
                    throwingGrouping,
                    resolution)));
            Assert.Equal("grouping failed", groupThrown.Message);
            Assert.Equal(0, CountStaged(scope, sheet));
            Assert.Empty(stagingWriter.ListOwned());
            Assert.Equal(["live:probe-1", "live:probe-2"], liveWriter.ListOwned());
            AssertViewRestored(
                fixture,
                scope,
                sheets,
                sentinel,
                witness,
                witnessBefore,
                sheetCountBefore,
                scrollRow,
                scrollColumn,
                zoom,
                selection);

            _output.WriteLine(string.Create(
                CultureInfo.InvariantCulture,
                $"Excel build {fixture.Excel.Version}; forcedCreateFailure=clean forcedGroupingFailure=clean residue=0"));

            CleanupLive(liveWriter, sentinel);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    internal static (
        Excel.Worksheet Sheet,
        Excel.Sheets Sheets,
        ExportBoundsOutcome Resolution,
        Excel.Shape Sentinel,
        Excel.Range Witness,
        object? WitnessBefore,
        int SheetCountBefore) Prepare(OfficeFixture fixture, OfficeFixture.ComScope scope)
    {
        Excel.Workbook workbook = fixture.CreateWorkbook();
        WorkbookInitialiseOutcome initialised =
            new ExcelWorkbookInitialiser(fixture.Excel).Initialise();
        Assert.True(initialised.Succeeded, $"Initialise refused: {initialised.Refusal}");
        OfficeFixture.DisableChartProtection(workbook);

        Excel.Sheets sheets = scope.Track(workbook.Worksheets);
        int sheetCountBefore = sheets.Count;
        Excel.Worksheet sheet = scope.Track(
            (Excel.Worksheet)sheets[GanttWorkbookContract.GanttSheetLabel]);
        Excel.ListObject table = ResolveTable(scope, sheet);
        Excel.ListRows listRows = scope.Track(table.ListRows);

        while (listRows.Count > 0)
        {
            scope.Track(listRows[1]).Delete();
        }

        Excel.ListRow spanRow = scope.Track(listRows.Add());
        Excel.ListRow milestoneRow = scope.Track(listRows.Add());
        SetRowCell(scope, spanRow, "Id", GanttRowId.New().Value);
        SetRowCell(scope, spanRow, "Type", "As-Planned Activity");
        SetRowCell(scope, spanRow, "Description", "Build phase");
        SetRowCell(scope, spanRow, "Start", 45658.0);
        SetRowCell(scope, spanRow, "Finish", 45670.0);
        SetRowCell(scope, milestoneRow, "Id", GanttRowId.New().Value);
        SetRowCell(scope, milestoneRow, "Type", "As-Planned Milestone");
        SetRowCell(scope, milestoneRow, "Start", 45675.0);
        TrimBlankBodyRows(scope, fixture.Excel, listRows);

        GanttTableReadOutcome read = new ExcelGanttTableReader(fixture.Excel).Read();
        Assert.True(read.Succeeded, $"Read refused: {read.Refusal}");
        GanttValidationOutcome validation = GanttRowValidator.Validate(read.Rows);
        Assert.Equal(2, validation.Events.Count);

        string[] measuredColumns =
        [
            .. GanttTableSchema.Default.Columns
                .Where(column => !column.IsHidden)
                .Select(column => column.Name),
        ];
        PanelGridOutcome gridOutcome =
            new ExcelPanelGridMeasurement(fixture.Excel).Measure(measuredColumns);
        Assert.True(gridOutcome.Succeeded, $"Grid refused: {gridOutcome.Refusal}");

        ConfigReadOutcome config = new ExcelConfigCatalogueReader(fixture.Excel).Read();
        Assert.True(config.Succeeded, $"Settings read refused: {config.Refusal}");

        SceneBuildRequestOutcome request = new ExcelSceneBuildRequestFactory(new AptosTextMetrics()).CreateExport(
            validation.Events,
            config.Settings,
            config.Styles,
            gridOutcome.Grid!,
            config.Metrics);
        Assert.True(request.Succeeded, $"Export request refused: {request.Message}");

        ExportBoundsOutcome resolution = ExportBoundsResolver.TryResolve(
            request.Request,
            ExportPanelTheme.Resolve(),
            ExportSelection.TablePlusChart);
        Assert.True(resolution.Succeeded, $"Resolution refused: {resolution.Refusal}");
        Assert.NotNull(resolution.Scene);

        Excel.Shapes shapes = scope.Track(sheet.Shapes);
        Excel.Shape sentinel = scope.Track(shapes.AddTextbox(
            MsoTextOrientation.msoTextOrientationHorizontal,
            40,
            500,
            120,
            24));
        sentinel.Name = "User note";
        sentinel.TextFrame2.TextRange.Text = "Do not touch";
        Excel.Range spanRange = scope.Track(spanRow.Range);
        Excel.Range witness = scope.Track(spanRange.Cells[1, ColumnIndex("Description")]);

        return (
            sheet,
            sheets,
            resolution,
            sentinel,
            witness,
            witness.Value2,
            sheetCountBefore);
    }

    internal static (int ScrollRow, int ScrollColumn, double Zoom, string Selection) CaptureView(
        OfficeFixture fixture,
        OfficeFixture.ComScope scope)
    {
        Excel.Window window = scope.Track(fixture.Excel.ActiveWindow);
        window.ScrollRow = 2;
        window.ScrollColumn = 1;
        window.Zoom = 90;
        return (window.ScrollRow, window.ScrollColumn, window.Zoom, SelectionAddress(scope, fixture.Excel));
    }

    internal static void AssertViewRestored(
        OfficeFixture fixture,
        OfficeFixture.ComScope scope,
        Excel.Sheets sheets,
        Excel.Shape sentinel,
        Excel.Range witness,
        object? witnessBefore,
        int sheetCountBefore,
        int scrollRow,
        int scrollColumn,
        double zoom,
        string selection)
    {
        Assert.Equal(sheetCountBefore, sheets.Count);
        Assert.Equal(witnessBefore, witness.Value2);
        Assert.Equal("User note", sentinel.Name);
        Assert.Equal("Do not touch", sentinel.TextFrame2.TextRange.Text);

        Excel.Window window = scope.Track(fixture.Excel.ActiveWindow);
        Assert.Equal(scrollRow, window.ScrollRow);
        Assert.Equal(scrollColumn, window.ScrollColumn);
        Assert.Equal(zoom, window.Zoom, precision: 6);
        Assert.Equal(selection, SelectionAddress(scope, fixture.Excel));
    }

    internal static void CleanupLive(ExcelShapeWriter liveWriter, Excel.Shape sentinel)
    {
        foreach (string identifier in liveWriter.ListOwned())
        {
            _ = liveWriter.Delete(identifier);
        }

        sentinel.Delete();
    }

    internal static OfficeShapeRequest LiveProbe(string id) =>
        new(
            id,
            OfficeShapeKind.Rectangle,
            new OfficeShapeGeometry(Bounds: new RectD(700, 480, 40, 12)),
            ZLayer.ActivityBody);

    internal static int CountStaged(OfficeFixture.ComScope scope, Excel.Worksheet sheet)
    {
        Excel.Shapes shapes = scope.Track(sheet.Shapes);
        int count = 0;
        int total = shapes.Count;
        for (int index = 1; index <= total; index++)
        {
            Excel.Shape shape = scope.Track(shapes.Item(index));
            if (shape.Name.StartsWith(ShapeStagingTag.NamePrefix, StringComparison.Ordinal))
            {
                count++;
            }
        }

        return count;
    }

    private static string SelectionAddress(OfficeFixture.ComScope scope, Excel.Application application)
    {
        object selection = scope.Track(application.Selection);
        return selection is Excel.Range range
            ? Convert.ToString(range.Address, CultureInfo.InvariantCulture)
            : string.Empty;
    }

    private static Excel.ListObject ResolveTable(OfficeFixture.ComScope scope, Excel.Worksheet sheet)
    {
        Excel.ListObjects objects = scope.Track(sheet.ListObjects);
        return scope.Track(objects[GanttTableSchema.TableName]);
    }

    private static void SetRowCell(
        OfficeFixture.ComScope scope,
        Excel.ListRow row,
        string headerName,
        object value)
    {
        Excel.Range range = scope.Track(row.Range);
        Excel.Range cell = scope.Track(range.Cells[1, ColumnIndex(headerName)]);
        cell.Value2 = value;
    }

    private static void TrimBlankBodyRows(
        OfficeFixture.ComScope scope,
        Excel.Application application,
        Excel.ListRows listRows)
    {
        for (int index = listRows.Count; index >= 1; index--)
        {
            Excel.ListRow row = scope.Track(listRows[index]);
            Excel.Range range = scope.Track(row.Range);
            double nonEmpty = Convert.ToDouble(
                application.WorksheetFunction.CountA(range),
                CultureInfo.InvariantCulture);
            if (nonEmpty == 0)
            {
                row.Delete();
            }
        }
    }

    private static int ColumnIndex(string headerName)
    {
        for (int index = 0; index < GanttTableSchema.Default.Columns.Count; index++)
        {
            if (string.Equals(GanttTableSchema.Default.Columns[index].Name, headerName, StringComparison.Ordinal))
            {
                return index + 1;
            }
        }

        throw new InvalidOperationException($"'{headerName}' is not a schema column.");
    }

    /// <summary>
    /// A faithfully-forwarding decorator over the real staging writer that records what
    /// every cleanup delete answered (the D2 probe).
    /// </summary>
    /// <remarks>
    /// F9: <c>BeginBatch</c>, <c>ListOwned</c> and every read are forwarded unchanged, so
    /// the composer and the lifecycle behave exactly as they do against the real port.
    /// </remarks>
    private sealed class RecordingShapePort(IShapeWritePort inner) : IShapeWritePort
    {
        public List<(string Id, ShapeWriteOutcome Outcome)> Deletes { get; } = [];

        public ShapeWriteOutcome Create(OfficeShapeRequest request) => inner.Create(request);

        public ShapeWriteOutcome Update(OfficeShapeRequest request) => inner.Update(request);

        public ShapeWriteOutcome UpdateContent(OfficeShapeRequest request) => inner.UpdateContent(request);

        public ShapeGroupWriteOutcome ApplyGroupAppearance(IReadOnlyList<OfficeShapeRequest> requests) =>
            inner.ApplyGroupAppearance(requests);

        public ShapeWriteOutcome UpdateGeometry(OfficeShapeRequest request) => inner.UpdateGeometry(request);

        public ShapeWriteOutcome Delete(string primitiveId)
        {
            ShapeWriteOutcome outcome = inner.Delete(primitiveId);
            Deletes.Add((primitiveId, outcome));
            return outcome;
        }

        public IDisposable BeginBatch() => inner.BeginBatch();

        public IReadOnlyList<string> ListOwned() => inner.ListOwned();

        public ShapeWriteOutcome LockOwnedShapes() => inner.LockOwnedShapes();

        public ShapeWriteOutcome UnlockShape(string primitiveId) => inner.UnlockShape(primitiveId);

        public bool GeometryMatches(OfficeShapeRequest request) => inner.GeometryMatches(request);

        public ShapeWriteOutcome ApplyZOrder(IReadOnlyList<string> backToFront) =>
            inner.ApplyZOrder(backToFront);
    }

    /// <summary>A port that throws at the Nth create and forwards everything else.</summary>
    private sealed class ThrowingCreatePort(IShapeWritePort inner, int failAt) : IShapeWritePort
    {
        private int _creates;

        public ShapeWriteOutcome Create(OfficeShapeRequest request)
        {
            _creates++;
            return _creates == failAt
                ? throw new InvalidOperationException("create failed")
                : inner.Create(request);
        }

        public ShapeWriteOutcome Update(OfficeShapeRequest request) => inner.Update(request);

        public ShapeWriteOutcome UpdateContent(OfficeShapeRequest request) => inner.UpdateContent(request);

        public ShapeGroupWriteOutcome ApplyGroupAppearance(IReadOnlyList<OfficeShapeRequest> requests) =>
            inner.ApplyGroupAppearance(requests);

        public ShapeWriteOutcome UpdateGeometry(OfficeShapeRequest request) => inner.UpdateGeometry(request);

        public ShapeWriteOutcome Delete(string primitiveId) => inner.Delete(primitiveId);

        public IDisposable BeginBatch() => inner.BeginBatch();

        public IReadOnlyList<string> ListOwned() => inner.ListOwned();

        public ShapeWriteOutcome LockOwnedShapes() => inner.LockOwnedShapes();

        public ShapeWriteOutcome UnlockShape(string primitiveId) => inner.UnlockShape(primitiveId);

        public bool GeometryMatches(OfficeShapeRequest request) => inner.GeometryMatches(request);

        public ShapeWriteOutcome ApplyZOrder(IReadOnlyList<string> backToFront) =>
            inner.ApplyZOrder(backToFront);
    }

    /// <summary>A grouping port whose grouping call throws; the ungroup is forwarded.</summary>
    private sealed class ThrowingGroupingPort(IShapeGroupingPort inner) : IShapeGroupingPort
    {
        public ShapeGroupReadback GroupStaged(IReadOnlyList<string> shapeNames) =>
            throw new InvalidOperationException("grouping failed");

        public ShapeWriteOutcome UngroupStaged(string groupName) => inner.UngroupStaged(groupName);
    }
}
