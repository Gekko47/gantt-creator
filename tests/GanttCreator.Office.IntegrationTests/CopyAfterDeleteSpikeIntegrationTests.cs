using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;
using Mso = Microsoft.Office.Core;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// R6.7's copy-after-delete/delayed-rendering compatibility spike.
/// </summary>
/// <remarks>
/// <para>
/// <b>The question.</b> After <c>ShapeRange.Copy</c>, does deleting the staging shapes
/// invalidate a delayed clipboard format on the reference host? The answer decides
/// whether Phase 6's composition rows may delete staging shapes immediately
/// (architecture "Editable export composition" step 7) or need the short-lived
/// staging-workbook fallback. The spike changes no production code; it records
/// evidence and the ADR proposes the outcome from that evidence only.
/// </para>
/// <para>
/// <b>Control arm.</b> Every delete cycle is paired with a no-delete cycle in its own
/// test, so a failure in the delete arm is attributable to the deletion rather than to
/// ambient clipboard flakiness. A failing delete arm beside a passing control arm is
/// evidence; both arms failing is a harness defect.
/// </para>
/// <para>
/// <b>Delay windows.</b> Two fixed windows (2 s and 10 s) are measurement parameters,
/// not synchronization: delayed rendering is serviced when a consumer requests the
/// format, so the paste assertion is the primary measurement and the windows only
/// expose spontaneous invalidation (for example another application re-taking the
/// clipboard). <see cref="Task.Delay(TimeSpan)"/> blocks the body without
/// <c>Thread.Sleep</c> and without a continuation, so the STA COM proxies stay on one
/// thread.
/// </para>
/// <para>
/// <b>Deterministic staging.</b> Four named shapes (three rectangles and one text box)
/// at a fixed origin, under a prefix deliberately distinct from the production
/// <c>GanttCreator.Owned.v1</c> tag. The Excel paste target is the same test-owned
/// worksheet at a fixed cell; the PowerPoint target is the test-owned presentation and
/// blank slide created by <see cref="PowerPointFixture"/>. Raw evidence records are
/// appended to the path named by <see cref="EvidencePathVariable"/> when it is set
/// (the office-evidence folder is git-ignored); the summarised matrix is committed in
/// the work item and ADR.
/// </para>
/// </remarks>
public class CopyAfterDeleteSpikeIntegrationTests(ITestOutputHelper output)
{
    /// <summary>Environment variable naming the JSON-lines evidence file, when one is wanted.</summary>
    internal const string EvidencePathVariable = "GANTTCREATOR_R6_7_EVIDENCE_PATH";

    private const string StagingNamePrefix = "Spike.Stage.v1.";
    private const int StagingShapeCount = 4;

    /// <summary>Clip cycles per Excel arm; each cycle is a full delete-and-paste sequence.</summary>
    private const int ExcelCycleCount = 3;

    /// <summary>Repeat count for the PowerPoint materialisation arm.</summary>
    private const int PowerPointCycleCount = 3;

    /// <summary>The Excel <c>msoGroup</c> shape type identifier.</summary>
    private const int MssoGroupShapeType = 6;

    private static readonly TimeSpan ShortDelay = TimeSpan.FromSeconds(2);
    private static readonly TimeSpan LongDelay = TimeSpan.FromSeconds(10);

    private readonly ITestOutputHelper _output = output;

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Delete_after_copy_keeps_the_clipboard_materialisable_for_an_excel_paste()
    {
        OfficeFixture fixture = new();

        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            using OfficeFixture.ComScope scope = new();

            // The workbook is the FIXTURE's proxy: it closes and releases it in teardown.
            // Releasing it through ComScope would detach the RCW and make that close throw.
            Excel.Workbook workbook = fixture.CreateWorkbook();
            Excel.Sheets sheets = scope.Track(workbook.Worksheets);
            Excel.Worksheet sheet = scope.Track((Excel.Worksheet)sheets[1]);
            string excelBuild = DescribeExcelBuild(fixture.Excel);
            ClipboardFormatProbe probe = new();

            for (int cycle = 1; cycle <= ExcelCycleCount; cycle++)
            {
                RunExcelCycle(scope, sheet, probe, excelBuild, cycle, deleteStagingShapes: true);
            }
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Copy_without_delete_is_the_control_arm_and_materialises_an_excel_paste()
    {
        OfficeFixture fixture = new();

        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            using OfficeFixture.ComScope scope = new();

            // The workbook is the FIXTURE's proxy: it closes and releases it in teardown.
            // Releasing it through ComScope would detach the RCW and make that close throw.
            Excel.Workbook workbook = fixture.CreateWorkbook();
            Excel.Sheets sheets = scope.Track(workbook.Worksheets);
            Excel.Worksheet sheet = scope.Track((Excel.Worksheet)sheets[1]);
            string excelBuild = DescribeExcelBuild(fixture.Excel);
            ClipboardFormatProbe probe = new();

            // One control cycle is the baseline arm; the delete arm carries the
            // repetition because the control establishes attribution, not survival.
            RunExcelCycle(scope, sheet, probe, excelBuild, cycle: 1, deleteStagingShapes: false);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Delete_after_copy_keeps_the_clipboard_materialisable_for_a_powerpoint_paste()
    {
        OfficeFixture excelFixture = new();
        PowerPointFixture powerPointFixture = new();

        try
        {
            await excelFixture.InitializeAsync().ConfigureAwait(true);

            // PowerPoint starts BEFORE the first copy: launching a host must not be able
            // to disturb the clipboard state the spike is about to measure.
            powerPointFixture.Initialize();

            using OfficeFixture.ComScope scope = new();

            // The workbook is the FIXTURE's proxy: it closes and releases it in teardown.
            // Releasing it through ComScope would detach the RCW and make that close throw.
            Excel.Workbook workbook = excelFixture.CreateWorkbook();
            Excel.Sheets sheets = scope.Track(workbook.Worksheets);
            Excel.Worksheet sheet = scope.Track((Excel.Worksheet)sheets[1]);
            string excelBuild = DescribeExcelBuild(excelFixture.Excel);
            ClipboardFormatProbe probe = new();

            for (int cycle = 1; cycle <= PowerPointCycleCount; cycle++)
            {
                bool includeDelayWindows = cycle == 1;
                ClipboardObservations observations = CopyAndObserve(
                    scope,
                    sheet,
                    probe,
                    includeDelayWindows,
                    deleteStagingShapes: true);

                // The delete arm must leave no staging shape behind before the paste.
                Assert.Equal(0, CountStagingShapes(sheet));

                PowerPointPasteOutcome paste = powerPointFixture.PasteSpecialShape();
                PowerPointUngroupOutcome ungroup = powerPointFixture.UngroupPastedShapes();

                _output.WriteLine(string.Create(
                    CultureInfo.InvariantCulture,
                    $"powerpoint paste: shapes={paste.ShapeCount} groupItems={paste.GroupItemCount} width={paste.WidthPt} height={paste.HeightPt}"));
                foreach (string description in paste.ShapeDescriptions)
                {
                    _output.WriteLine("powerpoint shape: " + description);
                }
                _output.WriteLine(string.Create(
                    CultureInfo.InvariantCulture,
                    $"powerpoint ungroup: children={ungroup.ChildCount} wasGroup={ungroup.WasGroup}"));

                // The pasted content must survive as four editable children, whether the
                // host returned one group (the expected shape) or four loose shapes.
                Assert.Equal(StagingShapeCount, ungroup.ChildCount);

                powerPointFixture.DeletePastedShapes();

                Record(new SpikeEvidenceRecord(
                    Test: nameof(Delete_after_copy_keeps_the_clipboard_materialisable_for_a_powerpoint_paste),
                    Arm: "delete",
                    Cycle: cycle,
                    ExcelBuild: excelBuild,
                    PowerPointBuild: powerPointFixture.BuildDescription,
                    StagingShapeCount: StagingShapeCount,
                    StagingDeleted: observations.StagingDeleted,
                    ImmediateSequenceNumber: observations.Immediate.SequenceNumber,
                    ImmediateFormatNames: [.. observations.Immediate.Names],
                    ShortDelaySequenceNumber: observations.ShortDelay?.SequenceNumber,
                    ShortDelayFormatNames: observations.ShortDelay is { } shortDelay ? [.. shortDelay.Names] : null,
                    LongDelaySequenceNumber: observations.LongDelay?.SequenceNumber,
                    LongDelayFormatNames: observations.LongDelay is { } longDelay ? [.. longDelay.Names] : null,
                    PastedShapeCount: paste.ShapeCount,
                    PastedGroupItemCount: paste.GroupItemCount,
                    PastedWidthPt: paste.WidthPt,
                    PastedHeightPt: paste.HeightPt,
                    UngroupedChildCount: ungroup.ChildCount));
            }
        }
        finally
        {
            await powerPointFixture.DisposeAsync().ConfigureAwait(true);
            await excelFixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    private void RunExcelCycle(
        OfficeFixture.ComScope scope,
        Excel.Worksheet sheet,
        ClipboardFormatProbe probe,
        string excelBuild,
        int cycle,
        bool deleteStagingShapes)
    {
        ClipboardObservations observations = CopyAndObserve(
            scope,
            sheet,
            probe,
            includeDelayWindows: true,
            deleteStagingShapes);

        // The control arm must paste while its staging group is still alive, so the
        // cleanup happens only after the paste. The delete arm has nothing to clean.
        ExcelPasteOutcome paste = PasteIntoExcelAndVerify(scope, sheet);

        if (observations.ControlGroup is { } controlGroup)
        {
            controlGroup.Delete();
        }

        Assert.Equal(0, CountStagingShapes(sheet));

        Record(new SpikeEvidenceRecord(
            Test: deleteStagingShapes
                ? nameof(Delete_after_copy_keeps_the_clipboard_materialisable_for_an_excel_paste)
                : nameof(Copy_without_delete_is_the_control_arm_and_materialises_an_excel_paste),
            Arm: deleteStagingShapes ? "delete" : "control",
            Cycle: cycle,
            ExcelBuild: excelBuild,
            PowerPointBuild: "not-applicable",
            StagingShapeCount: StagingShapeCount,
            StagingDeleted: observations.StagingDeleted,
            ImmediateSequenceNumber: observations.Immediate.SequenceNumber,
            ImmediateFormatNames: [.. observations.Immediate.Names],
            ShortDelaySequenceNumber: observations.ShortDelay?.SequenceNumber,
            ShortDelayFormatNames: observations.ShortDelay is { } shortDelay ? [.. shortDelay.Names] : null,
            LongDelaySequenceNumber: observations.LongDelay?.SequenceNumber,
            LongDelayFormatNames: observations.LongDelay is { } longDelay ? [.. longDelay.Names] : null,
            PastedShapeCount: paste.ShapeCount,
            PastedGroupItemCount: paste.GroupItemCount,
            PastedWidthPt: paste.WidthPt,
            PastedHeightPt: paste.HeightPt,
            UngroupedChildCount: paste.UngroupedChildCount));
    }

    /// <summary>
    /// Creates, groups, copies, optionally deletes, then reads the clipboard immediately
    /// and (when asked) after each bounded measurement window.
    /// </summary>
    private ClipboardObservations CopyAndObserve(
        OfficeFixture.ComScope scope,
        Excel.Worksheet sheet,
        ClipboardFormatProbe probe,
        bool includeDelayWindows,
        bool deleteStagingShapes)
    {
        string[] names = CreateStagingShapes(scope, sheet);
        Excel.ShapeRange staged = scope.Track(sheet.Shapes.Range[names]);
        Excel.Shape group = scope.Track(staged.Group());
        group.Name = StagingNamePrefix + "group";

        group.Copy();

        if (deleteStagingShapes)
        {
            group.Delete();
        }

        ClipboardSnapshot immediate = AssertSnapshot(probe, "immediate");
        ClipboardSnapshot? shortDelay = null;
        ClipboardSnapshot? longDelay = null;

        if (includeDelayWindows)
        {
            AwaitWindow(ShortDelay);
            shortDelay = AssertSnapshot(probe, "short-delay");
            AwaitWindow(LongDelay);
            longDelay = AssertSnapshot(probe, "long-delay");
        }

        // The control arm keeps its staging group alive so the paste that follows measures
        // the no-delete baseline. In the delete arm the proxy is already invalid and must
        // not be touched again.
        Excel.Shape? controlGroup = deleteStagingShapes ? null : group;
        return new ClipboardObservations(deleteStagingShapes, controlGroup, immediate, shortDelay, longDelay);
    }

    private static ExcelPasteOutcome PasteIntoExcelAndVerify(OfficeFixture.ComScope scope, Excel.Worksheet sheet)
    {
        // Worksheet.Paste returns void and Excel's PIA does not expose Shapes.Paste, so
        // the pasted range is identified deterministically by the shape-count delta. In
        // the delete arm the staging area is already empty; in the control arm the staging
        // group is alive and included in the "before" count, so neither contaminates the
        // delta.
        Excel.Range destination = scope.Track(sheet.Range["A60"]);
        destination.Select();
        int before = sheet.Shapes.Count;

        sheet.Paste();

        int after = sheet.Shapes.Count;
        int shapeCount = after - before;
        Assert.True(
            shapeCount > 0,
            string.Create(CultureInfo.InvariantCulture, $"Worksheet.Paste produced no shapes (before {before}, after {after})."));

        object[] pastedIndices = [.. Enumerable.Range(before + 1, shapeCount).Cast<object>()];
        Excel.ShapeRange pasted = scope.Track(sheet.Shapes.Range[pastedIndices]);
        double widthPt = pasted.Width;
        double heightPt = pasted.Height;

        // The same usable-range invariant as the PowerPoint fixture: a non-throwing COM
        // call is never treated as a successful materialisation by itself.
        Assert.True(widthPt > 0, string.Create(CultureInfo.InvariantCulture, $"Pasted width {widthPt} is not positive."));
        Assert.True(heightPt > 0, string.Create(CultureInfo.InvariantCulture, $"Pasted height {heightPt} is not positive."));

        int groupItemCount = 0;
        if (shapeCount == 1)
        {
            Excel.Shape first = scope.Track(pasted.Item(1));
            if ((int)first.Type == MssoGroupShapeType)
            {
                groupItemCount = first.GroupItems.Count;
            }
        }

        int childCount;
        if (groupItemCount > 0)
        {
            Excel.ShapeRange children = scope.Track(pasted.Ungroup());
            childCount = children.Count;
            children.Delete();
        }
        else
        {
            childCount = shapeCount;
            pasted.Delete();
        }

        Assert.Equal(StagingShapeCount, childCount);
        return new ExcelPasteOutcome(shapeCount, groupItemCount, widthPt, heightPt, childCount);
    }

    private string[] CreateStagingShapes(OfficeFixture.ComScope scope, Excel.Worksheet sheet)
    {
        string[] names = new string[StagingShapeCount];

        for (int index = 0; index < StagingShapeCount - 1; index++)
        {
            const double left = 400;
            const double top = 300;
            const double width = 100;
            const double height = 40;
            const double rowStep = 50;

            // msoShapeRectangle = 1. The origin is fixed and far from A1 so the staging
            // area never overlaps the paste target at A60.
            Excel.Shape rectangle = scope.Track(sheet.Shapes.AddShape(
                Mso.MsoAutoShapeType.msoShapeRectangle,
                (float)(left + (index * 120)),
                (float)(top + (index * rowStep)),
                (float)width,
                (float)height));
            rectangle.Name = string.Create(CultureInfo.InvariantCulture, $"{StagingNamePrefix}rect{index}");
            names[index] = rectangle.Name;
        }

        Excel.Shape textBox = scope.Track(sheet.Shapes.AddTextbox(
            Mso.MsoTextOrientation.msoTextOrientationHorizontal,
            400f,
            460f,
            340f,
            30f));
        textBox.Name = StagingNamePrefix + "text";
        names[StagingShapeCount - 1] = textBox.Name;

        foreach (string name in names)
        {
            _output.WriteLine("staging shape: " + name);
        }

        return names;
    }

    private ClipboardSnapshot AssertSnapshot(ClipboardFormatProbe probe, string observation)
    {
        ClipboardReadResult result = probe.Read();
        Assert.True(result.Succeeded, $"{observation}: {result.Failure}");
        ClipboardSnapshot snapshot = Assert.IsType<ClipboardSnapshot>(result.Snapshot);

        _output.WriteLine(string.Create(
            CultureInfo.InvariantCulture,
            $"{observation}: sequence={snapshot.SequenceNumber} formats=[{string.Join(", ", snapshot.Names)}]"));

        // Presence of a format is asserted; the exact set is evidence for the ADR. A
        // clipboard holding no format at all cannot materialise anything, so this is the
        // deterministic floor the delete arm must not fall through.
        Assert.NotEmpty(snapshot.Formats);

        // The measured Office drawing format is the one a shape paste consumes. Asserting
        // it by name (resolved from the host, not guessed) is what makes "the format
        // survived" a deterministic claim rather than a count.
        Assert.True(
            snapshot.Contains(ClipboardFormatProbe.Excel2007InternalShapeFormatName),
            $"{observation}: '{ClipboardFormatProbe.Excel2007InternalShapeFormatName}' is absent; formats=[{string.Join(", ", snapshot.Names)}]");
        return snapshot;
    }

    private static int CountStagingShapes(Excel.Worksheet sheet)
    {
        int count = 0;
        Excel.Shapes shapes = sheet.Shapes;
        try
        {
            int total = shapes.Count;
            for (int index = 1; index <= total; index++)
            {
                Excel.Shape shape = shapes.Item(index);
                try
                {
                    if (shape.Name.StartsWith(StagingNamePrefix, StringComparison.Ordinal))
                    {
                        count++;
                    }
                }
                finally
                {
                    _ = Marshal.ReleaseComObject(shape);
                }
            }
        }
        finally
        {
            _ = Marshal.ReleaseComObject(shapes);
        }

        return count;
    }

    private static void AwaitWindow(TimeSpan window)
    {
        // A measurement window, not synchronization: blocking here keeps the STA COM
        // proxies on the same thread (no continuation), and the assertion reads whatever
        // the window produced. No retry loop exists.
        Task.Delay(window).GetAwaiter().GetResult();
    }

    private static string DescribeExcelBuild(Excel.Application application) =>
        string.Create(CultureInfo.InvariantCulture, $"Excel {application.Version} build {application.Build}");

    private void Record(SpikeEvidenceRecord record)
    {
        string json = JsonSerializer.Serialize(record);
        _output.WriteLine(json);

        string? path = Environment.GetEnvironmentVariable(EvidencePathVariable);
        if (string.IsNullOrEmpty(path))
        {
            return;
        }

        try
        {
            File.AppendAllText(path, json + Environment.NewLine);
        }
        catch (IOException)
        {
            // Evidence capture is diagnostics: a write failure must not fail the spike,
            // and the same record is already on the test output above.
        }
        catch (UnauthorizedAccessException)
        {
            // Same.
        }
    }

    private sealed record ClipboardObservations(
        bool StagingDeleted,
        Excel.Shape? ControlGroup,
        ClipboardSnapshot Immediate,
        ClipboardSnapshot? ShortDelay,
        ClipboardSnapshot? LongDelay);

    private sealed record ExcelPasteOutcome(
        int ShapeCount,
        int GroupItemCount,
        double WidthPt,
        double HeightPt,
        int UngroupedChildCount);

    /// <summary>
    /// One evidence row. The raw record is diagnostics; the committed evidence table in
    /// the work item and ADR is generated from these rows.
    /// </summary>
    private sealed record SpikeEvidenceRecord(
        string Test,
        string Arm,
        int Cycle,
        string ExcelBuild,
        string PowerPointBuild,
        int StagingShapeCount,
        bool StagingDeleted,
        uint ImmediateSequenceNumber,
        string[] ImmediateFormatNames,
        uint? ShortDelaySequenceNumber,
        string[]? ShortDelayFormatNames,
        uint? LongDelaySequenceNumber,
        string[]? LongDelayFormatNames,
        int PastedShapeCount,
        int PastedGroupItemCount,
        double PastedWidthPt,
        double PastedHeightPt,
        int UngroupedChildCount);
}
