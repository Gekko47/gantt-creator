using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Export;
using GanttCreator.Core.Scene;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// R7.6 Office gate: the verified group reaches the user's active slide through the
/// production R7.2–R7.5 chain; a deck with unsaved changes survives intact; and the D-R7.2
/// zero-deck create arm is arranged on a fixture-started host the fixture also tears down.
/// </summary>
/// <remarks>
/// <para>
/// <b>Non-destructive.</b> Body 1 mutates only a fixture-owned deck, leaves the pasted group
/// removed through the production delete, and asserts the deck is still open, still dirty and
/// its file untouched. Body 2 creates the deck the transfer itself creates, on a host the
/// fixture started, so the fixture's quit is the clean teardown owner (no orphan process, no
/// user session at risk).
/// </para>
/// <para>
/// <b>One staging pass per body.</b> The R7.3 staging → verified copy path loads the
/// clipboard; the transfer then drives the production target resolver, paste stage and
/// placement stage, exactly as the R7.6 command's hook composes them.
/// </para>
/// </remarks>
public sealed class SendToPowerPointIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task The_transfer_lands_and_never_saves_or_closes_a_deck_with_unsaved_changes()
    {
        OfficeFixture excelFixture = new();
        PowerPointFixture powerPointFixture = new();
        string? path = null;
        try
        {
            await excelFixture.InitializeAsync().ConfigureAwait(true);

            // The work item's Required case: a deck with a file on disk plus an in-memory
            // mutation that is never saved.
            powerPointFixture.Initialize(withWindow: true, slideCount: 1);
            path = powerPointFixture.SavePrimaryPresentationToTemp();
            DateTime savedAt = File.GetLastWriteTimeUtc(path);
            powerPointFixture.AddUnsavedMarker("R7.6 unsaved marker");
            Assert.True(powerPointFixture.PrimaryPresentationHasUnsavedChanges());
            Assert.True(powerPointFixture.PrimarySlideHasText("R7.6 unsaved marker"));

            (StagedGroupOutcome group, ExportBoundsOutcome resolution) = StageAndCopy(excelFixture);

            using PowerPointTransferAdapter adapter = new(new PowerPointProxyOwnership(), _output.WriteLine);
            PowerPointTargetDecision decision = new PowerPointTargetResolver(adapter).Resolve();
            Assert.Equal(PowerPointTargetAction.ProceedWithActiveSlide, decision.Action);

            int groupId = PasteAndPlace(adapter, group, resolution);

            // The transfer landed on the user's deck, which is still open, still dirty, and
            // whose file was never written.
            Assert.True(powerPointFixture.PrimarySlideHasText("R7.6 unsaved marker"));
            Assert.True(powerPointFixture.PrimaryPresentationHasUnsavedChanges());
            Assert.Equal(savedAt, File.GetLastWriteTimeUtc(path));

            // Leave the deck as found: delete only the pasted group through the production port.
            PowerPointGroupDeleteResult deleted = adapter.DeletePastedGroup(groupId);
            Assert.Equal(PowerPointGroupDeleteOutcome.Deleted, deleted.Outcome);

            _output.WriteLine(string.Create(
                CultureInfo.InvariantCulture,
                $"host={powerPointFixture.BuildDescription} decision={decision.Action} group={groupId}"));
        }
        finally
        {
            await powerPointFixture.DisposeAsync().ConfigureAwait(true);
            TryDelete(path);
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task The_zero_deck_create_arm_pastes_and_ungroups_on_a_fixture_owned_host()
    {
        OfficeFixture excelFixture = new();
        PowerPointFixture powerPointFixture = new();
        try
        {
            await excelFixture.InitializeAsync().ConfigureAwait(true);

            // No presentation: the D-R7.2 create arm's state, arranged on a host this fixture
            // started so its Dispose is the clean teardown owner.
            powerPointFixture.InitializeWithoutPresentation();

            (StagedGroupOutcome group, ExportBoundsOutcome resolution) = StageAndCopy(excelFixture);

            using PowerPointTransferAdapter adapter = new(new PowerPointProxyOwnership(), _output.WriteLine);
            PowerPointTargetDecision decision = new PowerPointTargetResolver(adapter).Resolve();
            Assert.Equal(PowerPointTargetAction.CreatePresentation, decision.Action);

            PowerPointPresentationCreation creation = adapter.CreatePresentationWithBlankSlide();
            Assert.True(
                creation.Created,
                $"Creation refused: hresult=0x{creation.HResult:X8}");
            Assert.NotNull(creation.SlideIndex);

            int groupId = PasteAndPlace(adapter, group, resolution);

            // Editability: the pasted group ungroups to the staged children.
            int editable = powerPointFixture.UngroupShapeById(groupId);
            Assert.Equal(group.ChildOrder.Count, editable);
            powerPointFixture.DeletePastedShapes();

            _output.WriteLine(string.Create(
                CultureInfo.InvariantCulture,
                $"host={powerPointFixture.BuildDescription} createdSlide={creation.SlideIndex} children={editable}"));
        }
        finally
        {
            await powerPointFixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    private static (StagedGroupOutcome Group, ExportBoundsOutcome Resolution) StageAndCopy(
        OfficeFixture excelFixture)
    {
        using OfficeFixture.ComScope scope = new();
        (Excel.Worksheet sheet, Excel.Sheets sheets, ExportBoundsOutcome resolution, Excel.Shape sentinel, Excel.Range witness, object? witnessBefore, int sheetCountBefore) =
            StagingLifecycleIntegrationTests.Prepare(excelFixture, scope);
        _ = sheet;
        _ = sheets;
        _ = sentinel;
        _ = witness;
        _ = witnessBefore;
        _ = sheetCountBefore;

        ExcelShapeWriter stagingWriter = new(
            excelFixture.Excel,
            ownershipScheme: ShapeOwnershipScheme.Staging);

        using ExcelApplicationStateScope state = new(excelFixture.Excel);
        StagingLifecycleOutcome lifecycle = StagingLifecycle.Run(new(
            state,
            stagingWriter,
            new ExcelShapeGrouper(excelFixture.Excel),
            resolution,
            Copy: group => StagedShapeCopier.Copy(
                group,
                new ExcelShapeCopier(excelFixture.Excel),
                new WindowsClipboardFormatReader())));

        Assert.True(lifecycle.Succeeded, $"Lifecycle refused: {lifecycle.Refusal}");
        Assert.NotNull(lifecycle.Copy);
        Assert.True(
            lifecycle.Copy.Succeeded,
            $"Copy refused: {lifecycle.Copy.Refusal} {lifecycle.Copy.PortRefusal} {lifecycle.Copy.Detail}");

        StagedGroupOutcome group = Assert.IsType<StagedGroupOutcome>(lifecycle.Group);
        return (group, resolution);
    }

    private static int PasteAndPlace(
        PowerPointTransferAdapter adapter,
        StagedGroupOutcome group,
        ExportBoundsOutcome resolution)
    {
        RectD bounds = Assert.IsType<RectD>(group.Bounds);
        double insetPt = ChartPaddingPt.Uniform(
            GanttCatalogues.MetricDefault("ChartOuterPaddingPt")).MaxPt;

        PowerPointPasteStageResult paste = new PowerPointPasteStage(adapter).Paste(
            new PowerPointPasteExpectation(group.ChildOrder.Count, bounds.Width, bounds.Height));
        Assert.True(paste.Succeeded, $"Paste refused: {paste.Refusal} {paste.Validation?.Refusal}");
        int groupId = paste.GroupId
            ?? throw new InvalidOperationException("The paste stage reported no group identity.");

        PowerPointPlacementResult placed = new PowerPointPlacementStage(adapter).Place(
            groupId,
            resolution.Bounds!.Bounds,
            insetPt);
        Assert.True(
            placed.Succeeded,
            $"Placement refused: {placed.Refusal} fit={placed.FitRefusal} hresult=0x{placed.HResult:X8}");
        Assert.True(
            Assert.IsType<PowerPointPlacementDeltas>(placed.Deltas).MaxPt
                <= HostOwnedGeometry.GeometryMatchTolerancePt,
            "The placement read-back drifted.");
        return groupId;
    }

    private static void TryDelete(string? path)
    {
        if (path is null)
        {
            return;
        }

        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (IOException)
        {
            // The file is in the OS temp folder; a failed delete must not fail the body.
        }
        catch (UnauthorizedAccessException)
        {
            // Same.
        }
    }
}
