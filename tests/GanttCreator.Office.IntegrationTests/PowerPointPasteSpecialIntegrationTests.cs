using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Export;
using GanttCreator.Core.Scene;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// R7.3 Office gate: a real <c>PasteSpecial(ppPasteShape)</c> transfer at the production
/// shape count — the reference fixture composition the R6.8 live gate stages — driven through
/// the production target resolver and paste stage, with the returned range walked so the
/// group is located by type and the <c>1 + N</c> form is validated.
/// </summary>
/// <remarks>
/// <para>
/// <b>Measured, not assumed (work item Step 0).</b> The body records the measured range
/// structure — staged child count, staged envelope, pasted count, group count, pasted bounds
/// and the validation outcome — BEFORE any form assertion, so a divergent production form
/// leaves evidence in the run log instead of only a red test. A form difference is the work
/// item's stop condition, never a reason to widen the validator.
/// </para>
/// <para>
/// <b>Ownership.</b> The Excel fixture owns its workbook and the PowerPoint fixture owns the
/// host it started and its deck; the production adapter releases only its own proxies and
/// never saves, closes or quits. The group the production stage pasted is deleted through the
/// stage result's group identity, because the fixture's own paste tracking cannot see it.
/// </para>
/// <para>
/// <b>The empty-clipboard arm measures the D4 filter.</b> It empties the clipboard, pastes,
/// and records the host's HRESULT while asserting the classification; a divergent HRESULT is
/// the R7.3 stop condition that corrects the classifier's documented candidate from evidence.
/// </para>
/// </remarks>
public sealed class PowerPointPasteSpecialIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task A_production_sized_paste_returns_the_group_and_its_children_and_validates()
    {
        OfficeFixture excelFixture = new();
        PowerPointFixture powerPointFixture = new();

        try
        {
            await excelFixture.InitializeAsync().ConfigureAwait(true);

            // PowerPoint starts BEFORE the copy so launching it cannot disturb the clipboard
            // the lifecycle is about to fill (the R6.7 spike's ordering).
            powerPointFixture.Initialize(withWindow: true, slideCount: 1);

            using OfficeFixture.ComScope scope = new();
            (Excel.Worksheet sheet, Excel.Sheets sheets, ExportBoundsOutcome resolution, Excel.Shape sentinel, Excel.Range witness, object? witnessBefore, int sheetCountBefore) =
                StagingLifecycleIntegrationTests.Prepare(excelFixture, scope);

            ExcelShapeWriter stagingWriter = new(
                excelFixture.Excel,
                ownershipScheme: ShapeOwnershipScheme.Staging);
            ExcelShapeWriter liveWriter = new(excelFixture.Excel);
            Assert.True(liveWriter.Create(StagingLifecycleIntegrationTests.LiveProbe("live:probe-1")).Succeeded);

            (int scrollRow, int scrollColumn, double zoom, string selection) =
                StagingLifecycleIntegrationTests.CaptureView(excelFixture, scope);

            // The lifecycle disposes the scope inside its finally; `using` is the
            // analyzer-satisfying double dispose, which ExcelApplicationStateScope guards.
            using ExcelApplicationStateScope state = new(excelFixture.Excel);
            StagingLifecycleOutcome outcome = StagingLifecycle.Run(new(
                state,
                stagingWriter,
                new ExcelShapeGrouper(excelFixture.Excel),
                resolution,
                Copy: group => StagedShapeCopier.Copy(
                    group,
                    new ExcelShapeCopier(excelFixture.Excel),
                    new WindowsClipboardFormatReader())));

            Assert.True(outcome.Succeeded, $"Lifecycle refused: {outcome.Refusal}");
            Assert.NotNull(outcome.Copy);
            Assert.True(
                outcome.Copy.Succeeded,
                $"Copy refused: {outcome.Copy.Refusal} {outcome.Copy.PortRefusal} {outcome.Copy.Detail}");

            StagedGroupOutcome group = Assert.IsType<StagedGroupOutcome>(outcome.Group);
            Assert.True(group.Succeeded);
            RectD bounds = Assert.IsType<RectD>(group.Bounds);
            int stagedChildren = group.ChildOrder.Count;

            PowerPointTransferAdapter adapter = new(new PowerPointProxyOwnership());
            try
            {
                PowerPointTargetDecision decision = new PowerPointTargetResolver(adapter).Resolve();
                Assert.Equal(PowerPointTargetAction.ProceedWithActiveSlide, decision.Action);

                PowerPointPasteExpectation expectation = new(stagedChildren, bounds.Width, bounds.Height);
                PowerPointPasteStageResult result = new PowerPointPasteStage(adapter).Paste(expectation);

                // Record before asserting: a divergent production form leaves evidence.
                PowerPointPasteFacts? facts = result.Validation?.ValidatedFacts;
                _output.WriteLine(string.Create(
                    CultureInfo.InvariantCulture,
                    $"excel={excelFixture.Excel.Version} build={excelFixture.Excel.Build} host={powerPointFixture.BuildDescription}"
                        + $" staged={stagedChildren} stagedBounds={bounds.Width}x{bounds.Height}"
                        + $" pasted={(facts is null ? -1 : facts.Count)} groups={(facts is null ? -1 : facts.GroupCount)}"
                        + $" pastedBounds={(facts is null ? 0 : facts.WidthPt)}x{(facts is null ? 0 : facts.HeightPt)}"
                        + $" succeeded={result.Succeeded} refusal={result.Refusal}"
                        + $" hresult=0x{result.HResult:X8} validation={result.Validation?.Refusal}"
                        + $" groupId={result.GroupId}"));

                Assert.True(result.Succeeded, $"Paste refused: {result.Refusal} {result.Validation?.Refusal}");
                Assert.NotNull(facts);
                Assert.Equal(stagedChildren + 1, facts.Count);
                Assert.Equal(1, facts.GroupCount);
                Assert.Equal(PowerPointPasteValidator.MsoGroupShapeType, facts.GroupShapeType);

                // Cleanup: delete the group the production stage pasted.
                if (result.GroupId is int groupId)
                {
                    _ = powerPointFixture.DeleteShapeById(groupId);
                }
            }
            finally
            {
                adapter.Dispose();
            }

            // Zero staging residue and the live chart intact.
            Assert.Empty(stagingWriter.ListOwned());
            Assert.Equal(0, StagingLifecycleIntegrationTests.CountStaged(scope, sheet));
            Assert.Equal(["live:probe-1"], liveWriter.ListOwned());

            StagingLifecycleIntegrationTests.AssertViewRestored(
                excelFixture,
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

            StagingLifecycleIntegrationTests.CleanupLive(liveWriter, sentinel);
        }
        finally
        {
            await powerPointFixture.DisposeAsync().ConfigureAwait(true);
            await excelFixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task An_empty_clipboard_paste_is_classified_PastePayloadMissing()
    {
        ClipboardFormatProbe probe = new();
        Assert.True(probe.TryEmpty(), "The harness could not empty the clipboard.");

        PowerPointFixture powerPointFixture = new();
        try
        {
            powerPointFixture.Initialize(withWindow: true, slideCount: 1);

            PowerPointTransferAdapter adapter = new(new PowerPointProxyOwnership());
            try
            {
                PowerPointTargetDecision decision = new PowerPointTargetResolver(adapter).Resolve();
                Assert.Equal(PowerPointTargetAction.ProceedWithActiveSlide, decision.Action);

                PowerPointPasteStageResult result =
                    new PowerPointPasteStage(adapter).Paste(new(1, 100, 100));

                _output.WriteLine(string.Create(
                    CultureInfo.InvariantCulture,
                    $"host={powerPointFixture.BuildDescription} refusal={result.Refusal}"
                        + $" hresult=0x{result.HResult:X8}"));

                Assert.Equal(PowerPointPasteStageRefusal.PastePayloadMissing, result.Refusal);
            }
            finally
            {
                adapter.Dispose();
            }
        }
        finally
        {
            await powerPointFixture.DisposeAsync().ConfigureAwait(true);
        }
    }
}
