using System.Globalization;
using GanttCreator.Core.Scene;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// R6.6 Office gate: the production copy stage copies the staged group, verifies the measured
/// Office drawing formats on the real clipboard, and the clipboard still materialises the group
/// after the staging shapes are gone (ADR-0066 D1/D2).
/// </summary>
/// <remarks>
/// <para>
/// The harness drives the production chain exactly as R6.3-R6.5's gates do: initialise a real
/// table, write real rows, read and validate them, measure the real grid, build the EXPORT
/// request, resolve the bounds, then run <see cref="StagingLifecycle.Run"/> with the real
/// <see cref="ExcelShapeWriter"/> (staging scheme), the real <see cref="ExcelShapeGrouper"/>,
/// the real <see cref="ExcelShapeCopier"/> and the real
/// <see cref="WindowsClipboardFormatReader"/> filling the copy seam.
/// </para>
/// <para>
/// <b>The observed format set is recorded, and only the required pair is asserted.</b> The
/// R6.7 spike asserted presence and kept the exact set as ADR evidence; R6.6's contract is the
/// measured drawing pair (ADR-0066 D2), so a non-required format's variance is recorded for
/// R10.1/R10.5 rather than failing this gate, while a required format's absence fails it - the
/// work item's stop condition firing, never a widening of the check.
/// </para>
/// <para>
/// <b>Materialisation is the primary measurement.</b> After the lifecycle's cleanup, the body
/// pastes into a test-owned cell, identifies the pasted range by the shape-count delta, asserts
/// positive bounds and the expected child count, ungroups to prove editability, and deletes the
/// pasted shapes. A non-throwing paste is never treated as materialisation by itself.
/// </para>
/// </remarks>
public sealed class ClipboardCopyIntegrationTests(ITestOutputHelper output)
{
    /// <summary>The full 15-format set the R6.7 spike recorded on the reference host.</summary>
    private static readonly string[] _recordedSpikeFormatSet =
    [
        "DataObject",
        "Preferred DropEffect",
        "InShellDragLoop",
        "Excel 2007 Internal Shape",
        "Art::GVML ClipFormat",
        "image/svg+xml",
        "PNG",
        "JFIF",
        "GIF",
        "CF_BITMAP",
        "CF_ENHMETAFILE",
        "CF_METAFILEPICT",
        "Ole Private Data",
        "CF_DIB",
        "CF_DIBV5",
    ];

    /// <summary>The Excel <c>msoGroup</c> shape type identifier.</summary>
    private const int _msoGroupShapeType = 6;

    private readonly ITestOutputHelper _output = output;

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task The_verified_copy_keeps_the_formats_materialisable_after_cleanup()
    {
        OfficeFixture fixture = new();

        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            using OfficeFixture.ComScope scope = new();
            (
                Excel.Worksheet sheet,
                Excel.Sheets sheets,
                ExportBoundsOutcome resolution,
                Excel.Shape sentinel,
                Excel.Range witness,
                object? witnessBefore,
                int sheetCountBefore) = StagingLifecycleIntegrationTests.Prepare(fixture, scope);

            ExcelShapeWriter stagingWriter = new(
                fixture.Excel,
                ownershipScheme: ShapeOwnershipScheme.Staging);
            ExcelShapeWriter liveWriter = new(fixture.Excel);
            Assert.True(liveWriter.Create(StagingLifecycleIntegrationTests.LiveProbe("live:probe-1")).Succeeded);

            (int scrollRow, int scrollColumn, double zoom, string selection) =
                StagingLifecycleIntegrationTests.CaptureView(fixture, scope);

            // The lifecycle disposes the scope inside its finally; `using` is the
            // analyzer-satisfying double dispose, which ExcelApplicationStateScope guards.
            using ExcelApplicationStateScope state = new(fixture.Excel);
            StagingLifecycleOutcome outcome = StagingLifecycle.Run(new(
                state,
                stagingWriter,
                new ExcelShapeGrouper(fixture.Excel),
                resolution,
                Copy: group => StagedShapeCopier.Copy(
                    group,
                    new ExcelShapeCopier(fixture.Excel),
                    new WindowsClipboardFormatReader())));

            Assert.True(outcome.Succeeded, $"Lifecycle refused: {outcome.Refusal}");
            Assert.NotNull(outcome.Copy);
            Assert.True(
                outcome.Copy.Succeeded,
                $"Copy refused: {outcome.Copy.Refusal} {outcome.Copy.PortRefusal} {outcome.Copy.Detail}");
            Assert.Null(outcome.Copy.Refusal);

            // Post-cleanup inspection with the spike's independent probe: the measured pair
            // survives the staging deletion (ADR-0066 D1), observed separately from the
            // production reader that verified it inside the seam.
            ClipboardFormatProbe probe = new();
            ClipboardReadResult read = probe.Read();
            Assert.True(read.Succeeded, read.Failure);
            ClipboardSnapshot snapshot = Assert.IsType<ClipboardSnapshot>(read.Snapshot);

            foreach (string required in StagedShapeCopier.RequiredDrawingFormatNames)
            {
                Assert.True(
                    snapshot.Contains(required),
                    $"'{required}' is absent after cleanup; formats=[{string.Join(", ", snapshot.Names)}]");
            }

            // Recorded comparison against the spike's measured set (plan finding A1): a
            // non-required variance is evidence for R10.1/R10.5, not a failure of this gate.
            List<string> divergences = [.. _recordedSpikeFormatSet
                .Where(name => !snapshot.Contains(name))
                .Select(name => "missing:" + name)
                .Concat(snapshot.Names
                    .Where(name => !_recordedSpikeFormatSet.Contains(name, StringComparer.OrdinalIgnoreCase))
                    .Select(name => "extra:" + name))];

            _output.WriteLine(string.Create(
                CultureInfo.InvariantCulture,
                $"Excel {fixture.Excel.Version} build {fixture.Excel.Build};"
                    + $" sequence={snapshot.SequenceNumber}"
                    + $" formats=[{string.Join(", ", snapshot.Names)}]"
                    + $" divergences={divergences.Count} [{string.Join(", ", divergences)}]"));

            // The lifecycle's finally restored the view and selection; assert before the
            // paste moves the selection to its destination.
            StagingLifecycleIntegrationTests.AssertViewRestored(
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

            // Materialisation (ADR-0066 D2): a real consumer paste must produce the group.
            int stagedCount = outcome.Group!.ChildOrder.Count;
            ExcelPasteOutcome paste = PasteIntoExcelAndVerify(scope, sheet, stagedCount);

            // Zero residue and the live chart intact.
            Assert.Empty(stagingWriter.ListOwned());
            Assert.Equal(0, StagingLifecycleIntegrationTests.CountStaged(scope, sheet));
            Assert.Equal(["live:probe-1"], liveWriter.ListOwned());

            _output.WriteLine(string.Create(
                CultureInfo.InvariantCulture,
                $"paste: shapes={paste.ShapeCount} groupItems={paste.GroupItemCount}"
                    + $" width={paste.WidthPt} height={paste.HeightPt} children={paste.ChildCount}"
                    + $" staged={stagedCount} residue={outcome.Residue.Count}"));

            StagingLifecycleIntegrationTests.CleanupLive(liveWriter, sentinel);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    /// <summary>
    /// Pastes the clipboard into a fixed test-owned cell, identifies the pasted range by the
    /// shape-count delta, asserts positive bounds, and proves editability by ungrouping.
    /// </summary>
    /// <param name="scope">The COM scope tracking every proxy this body acquires.</param>
    /// <param name="sheet">The test-owned worksheet.</param>
    /// <param name="expectedChildren">The staged set's child count.</param>
    /// <returns>What the materialisation produced.</returns>
    internal static ExcelPasteOutcome PasteIntoExcelAndVerify(
        OfficeFixture.ComScope scope,
        Excel.Worksheet sheet,
        int expectedChildren)
    {
        // Worksheet.Paste returns void and the PIA does not expose Shapes.Paste, so the
        // pasted range is identified deterministically by the shape-count delta.
        Excel.Range destination = scope.Track(sheet.Range["A60"]);
        destination.Select();
        int before = sheet.Shapes.Count;

        sheet.Paste();

        int after = sheet.Shapes.Count;
        int shapeCount = after - before;
        Assert.True(
            shapeCount > 0,
            string.Create(
                CultureInfo.InvariantCulture,
                $"Worksheet.Paste produced no shapes (before {before}, after {after})."));

        object[] pastedIndices = [.. Enumerable.Range(before + 1, shapeCount).Cast<object>()];
        Excel.ShapeRange pasted = scope.Track(sheet.Shapes.Range[pastedIndices]);
        double widthPt = pasted.Width;
        double heightPt = pasted.Height;

        // The same usable-range invariant the spike used: a non-throwing COM call is never
        // treated as a successful materialisation by itself.
        Assert.True(widthPt > 0, string.Create(CultureInfo.InvariantCulture, $"Pasted width {widthPt} is not positive."));
        Assert.True(heightPt > 0, string.Create(CultureInfo.InvariantCulture, $"Pasted height {heightPt} is not positive."));

        int groupItemCount = 0;
        if (shapeCount == 1)
        {
            Excel.Shape first = scope.Track(pasted.Item(1));
            if ((int)first.Type == _msoGroupShapeType)
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

        Assert.Equal(expectedChildren, childCount);
        return new ExcelPasteOutcome(shapeCount, groupItemCount, widthPt, heightPt, childCount);
    }

    /// <summary>What one paste materialisation produced.</summary>
    /// <param name="ShapeCount">The shapes the paste added.</param>
    /// <param name="GroupItemCount">The group's children when the paste produced one group.</param>
    /// <param name="WidthPt">The pasted range's width in points.</param>
    /// <param name="HeightPt">The pasted range's height in points.</param>
    /// <param name="ChildCount">The editable children after ungrouping.</param>
    internal sealed record ExcelPasteOutcome(
        int ShapeCount,
        int GroupItemCount,
        double WidthPt,
        double HeightPt,
        int ChildCount);
}
