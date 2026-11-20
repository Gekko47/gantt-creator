using System.Diagnostics;
using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// Live gate for <see cref="IShapeWritePort.ApplyGroupAppearance"/>: one
/// <c>ShapeRange</c> write must carry the shared fill, stroke, line width, and
/// typography to every shape in an appearance group, while the string stays a
/// per-shape write through <see cref="IShapeWritePort.UpdateContent"/>.
/// </summary>
/// <remarks>
/// <para>
/// The expected host values are literals transcribed from the entity guide's
/// <c>#RRGGBB</c> tokens and packed by hand into the host's
/// <c>blue * 65536 + green * 256 + red</c> order. A live test that computed its
/// expectation through <see cref="OfficeStyleMapper"/> would prove only that the
/// adapter called the mapper; it could not catch a swapped byte order or a host
/// that normalised the value.
/// </para>
/// <para>
/// The grouped pass is exercised with three rectangles of one appearance and two
/// text boxes of another, so both the multi-shape <c>ShapeRange</c> path and the
/// per-shape content path run against the real host. Every shape is read back
/// after the write, so a host that applied the range style to only the first
/// shape, or a content write that landed on the wrong label, fails here.
/// </para>
/// </remarks>
public sealed class ShapeGroupUpdateIntegrationTests(ITestOutputHelper output)
{
    private const string _tokenFont = "Aptos";

    // Entity-guide tokens for the grouped update, and the host's packed values.
    private const string _changedFillHex = "#FFC7CE";
    private const int _changedFillRgb = 0xCEC7FF;
    private const string _changedStrokeHex = "#C00000";
    private const int _changedStrokeRgb = 0x0000C0;
    private const double _changedWidthPt = 1.25d;

    private const string _changedTextHex = "#1F4E79";
    private const int _changedTextRgb = 0x794E1F;
    private const double _changedFontSizePt = 11d;

    private const int _measureShapeCount = 20;

    private readonly ITestOutputHelper _output = output;

    private static string XllPath => OfficeFixtureTests.ResolvePackedXllPath();

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task A_grouped_update_reaches_every_shape_and_its_own_text()
    {
        OfficeFixture fixture = new() { SuppressLeakSignal = true };

        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath),
                $"Application.RegisterXLL returned false for '{XllPath}'.");

            using OfficeFixture.ComScope scope = new();
            Excel.Workbook workbook = scope.Track(fixture.CreateWorkbook());
            WorkbookInitialiseOutcome initialised = new ExcelWorkbookInitialiser(fixture.Excel).Initialise();
            Assert.True(initialised.Succeeded, $"Initialise refused: {initialised.Refusal}");

            ExcelShapeWriter writer = new(fixture.Excel);

            OfficeShapeRequest[] created =
            [
                Bar("row-1:bar", 30d),
                Bar("row-2:bar", 60d),
                Bar("row-3:bar", 90d),
                Label("row-1:label", "Alpha", 30d),
                Label("row-2:label", "Beta", 60d),
            ];
            foreach (OfficeShapeRequest request in created)
            {
                ShapeWriteOutcome outcome = writer.Create(request);
                Assert.True(outcome.Succeeded, $"Create refused for '{request.PrimitiveId}': {outcome.Refusal}");
            }

            // One appearance per group: the rectangles share a changed fill, stroke,
            // and width; the labels share changed typography. Only the string differs
            // per label, which is why content follows on the per-shape call.
            OfficeShapeRequest[] restyledBars =
            [
                Bar("row-1:bar", 30d, _changedFillHex, _changedStrokeHex, _changedWidthPt),
                Bar("row-2:bar", 60d, _changedFillHex, _changedStrokeHex, _changedWidthPt),
                Bar("row-3:bar", 90d, _changedFillHex, _changedStrokeHex, _changedWidthPt),
            ];
            OfficeShapeRequest[] restyledLabels =
            [
                Label("row-1:label", "Alpha updated", 30d, _changedTextHex, _changedFontSizePt),
                Label("row-2:label", "Beta updated", 60d, _changedTextHex, _changedFontSizePt),
            ];

            ShapeGroupWriteOutcome bars = writer.ApplyGroupAppearance(restyledBars);
            Assert.True(bars.Succeeded, $"Rectangle group refused: {bars.Refusal} on '{bars.FailedPrimitiveId}'.");

            // Read-back verification, proved against the live host: a bar keeps its box, so
            // the grouped pass reports no per-shape work; a label's string changed, so the
            // pass names exactly the labels the caller must still dispatch.
            Assert.NotNull(bars.Mismatched);
            Assert.Empty(bars.Mismatched);

            ShapeGroupWriteOutcome labels = writer.ApplyGroupAppearance(restyledLabels);
            Assert.True(labels.Succeeded, $"Label group refused: {labels.Refusal} on '{labels.FailedPrimitiveId}'.");
            Assert.NotNull(labels.Mismatched);
            Assert.Equal(
                restyledLabels.Select(label => label.PrimitiveId),
                labels.Mismatched);

            foreach (OfficeShapeRequest restyledLabel in restyledLabels)
            {
                ShapeWriteOutcome content = writer.UpdateContent(restyledLabel);
                Assert.True(content.Succeeded, $"Content refused for '{restyledLabel.PrimitiveId}': {content.Refusal}");
            }

            // Once the content has landed, the SAME group pass converges: this is the
            // condition the reconciler uses to skip every per-shape dispatch on the next
            // refresh, so it must hold against the host and not only against a fake.
            ShapeGroupWriteOutcome convergedLabels = writer.ApplyGroupAppearance(restyledLabels);
            Assert.True(convergedLabels.Succeeded, $"Converged label pass refused: {convergedLabels.Refusal}.");
            Assert.NotNull(convergedLabels.Mismatched);
            Assert.Empty(convergedLabels.Mismatched);

            Excel.Worksheet sheet = (Excel.Worksheet)scope.Track(
                workbook.Sheets[GanttWorkbookContract.GanttSheetLabel]);

            foreach (OfficeShapeRequest bar in restyledBars)
            {
                AssertGroupedStyle(scope, sheet, bar.PrimitiveId);
            }

            foreach (OfficeShapeRequest label in restyledLabels)
            {
                AssertGroupedTypography(scope, sheet, label.PrimitiveId, label.Text ?? string.Empty);
            }

            _output.WriteLine(
                $"Excel build {fixture.Excel.Version}; grouped rectangles={restyledBars.Length}"
                + $" grouped labels={restyledLabels.Length} contentWrites={restyledLabels.Length}");

            foreach (OfficeShapeRequest request in created)
            {
                scope.Track(sheet.Shapes.Item(request.PrimitiveId)).Delete();
            }
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    /// <summary>
    /// Measures the grouped pass against the same appearance writes issued one shape at
    /// a time — the pre-grouping behaviour, since the single-shape path is
    /// <c>ApplyStyle</c> + <c>ApplyTextFormatting</c> per shape.
    /// </summary>
    /// <remarks>
    /// The timings are logged, not asserted: a busy host is not a deterministic subject,
    /// and the R6.5 D2 probe set the precedent of recording a measured answer instead of
    /// asserting an assumption. The two rounds use the cheaper round for each side so a
    /// single host stall cannot distort the comparison.
    /// </remarks>
    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task A_grouped_update_is_measured_against_the_same_writes_per_shape()
    {
        OfficeFixture fixture = new() { SuppressLeakSignal = true };

        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath),
                $"Application.RegisterXLL returned false for '{XllPath}'.");

            using OfficeFixture.ComScope scope = new();
            Excel.Workbook workbook = scope.Track(fixture.CreateWorkbook());
            WorkbookInitialiseOutcome initialised = new ExcelWorkbookInitialiser(fixture.Excel).Initialise();
            Assert.True(initialised.Succeeded, $"Initialise refused: {initialised.Refusal}");

            ExcelShapeWriter writer = new(fixture.Excel);
            OfficeShapeRequest[] created = LabelGrid("measure:label");
            foreach (OfficeShapeRequest request in created)
            {
                ShapeWriteOutcome outcome = writer.Create(request);
                Assert.True(outcome.Succeeded, $"Create refused for '{request.PrimitiveId}': {outcome.Refusal}");
            }

            OfficeShapeRequest[] restyled = LabelGrid(
                "measure:label",
                _changedTextHex,
                _changedFontSizePt);

            // The production reconciler writes inside a batch, which hoists the worksheet
            // and table resolution for the whole pass. The measurement runs inside a batch
            // for the same reason: outside one, the single-shape side would pay the
            // resolution cost twenty times and overstate the grouping win.
            double baselineMs;
            double groupedMs;
            using (writer.BeginBatch())
            {
                // The first write of a new appearance pays host caches, so both sides run
                // once before the clock starts.
                Assert.True(writer.ApplyGroupAppearance([restyled[0]]).Succeeded, "Warm-up single write refused.");
                Assert.True(writer.ApplyGroupAppearance(restyled).Succeeded, "Warm-up group write refused.");

                (baselineMs, groupedMs) = MeasureAppearanceWrites(writer, restyled);
            }
            _output.WriteLine(string.Create(
                CultureInfo.InvariantCulture,
                $"grouped-update probe: shapes={_measureShapeCount} "
                + $"per-shape={baselineMs:F1}ms grouped={groupedMs:F1}ms "
                + $"speed-up={baselineMs / groupedMs:F1}x"));

            Excel.Worksheet sheet = (Excel.Worksheet)scope.Track(
                workbook.Sheets[GanttWorkbookContract.GanttSheetLabel]);
            foreach (OfficeShapeRequest request in created)
            {
                scope.Track(sheet.Shapes.Item(request.PrimitiveId)).Delete();
            }
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    private static (double BaselineMs, double GroupedMs) MeasureAppearanceWrites(
        ExcelShapeWriter writer,
        OfficeShapeRequest[] restyled)
    {
        long baselineTicks = long.MaxValue;
        long groupedTicks = long.MaxValue;
        for (int round = 0; round < 2; round++)
        {
            long beforeBaseline = Stopwatch.GetTimestamp();
            foreach (OfficeShapeRequest request in restyled)
            {
                Assert.True(
                    writer.ApplyGroupAppearance([request]).Succeeded,
                    $"Single-shape appearance write refused for '{request.PrimitiveId}'.");
            }

            baselineTicks = Math.Min(baselineTicks, Stopwatch.GetTimestamp() - beforeBaseline);

            long beforeGrouped = Stopwatch.GetTimestamp();
            Assert.True(writer.ApplyGroupAppearance(restyled).Succeeded, "Grouped appearance write refused.");
            groupedTicks = Math.Min(groupedTicks, Stopwatch.GetTimestamp() - beforeGrouped);
        }

        double frequency = Stopwatch.Frequency;
        return (baselineTicks * 1000d / frequency, groupedTicks * 1000d / frequency);
    }

    private static OfficeShapeRequest[] LabelGrid(
        string prefix,
        string textHex = "#000000",
        double fontSizePt = 9d)
    {
        OfficeShapeRequest[] requests = new OfficeShapeRequest[_measureShapeCount];
        for (int index = 0; index < requests.Length; index++)
        {
            requests[index] = Label(
                FormattableString.Invariant($"{prefix}-{index}"),
                FormattableString.Invariant($"Row {index}"),
                20d + (index * 18d),
                textHex,
                fontSizePt);
        }

        return requests;
    }

    private static void AssertGroupedStyle(OfficeFixture.ComScope scope, Excel.Worksheet sheet, string id)
    {
        Excel.Shape shape = scope.Track(sheet.Shapes.Item(id));
        Excel.FillFormat fill = shape.Fill;
        Excel.LineFormat line = shape.Line;

        Assert.Equal(_changedFillRgb, fill.ForeColor?.RGB);
        Assert.Equal(MsoFillType.msoFillSolid, fill.Type);
        Assert.Equal(MsoTriState.msoTrue, fill.Visible);
        Assert.Equal(_changedStrokeRgb, line.ForeColor?.RGB);
        Assert.Equal(MsoTriState.msoTrue, line.Visible);
        AssertPoint($"{id} line weight", _changedWidthPt, line.Weight);

        // The name and ownership tag are the reconciliation key and proof; a
        // grouped style write must not disturb them.
        Assert.Equal(id, shape.Name);
        Assert.Equal(ShapeOwnershipTag.ForPrimitiveId(id), shape.AlternativeText);
    }

    private static void AssertGroupedTypography(
        OfficeFixture.ComScope scope,
        Excel.Worksheet sheet,
        string id,
        string expectedText)
    {
        Excel.Shape shape = scope.Track(sheet.Shapes.Item(id));
        TextRange2 range = scope.Track(shape.TextFrame2.TextRange);

        // The string is per shape: a range write cannot carry it, so the grouped
        // pass must leave this to UpdateContent.
        Assert.Equal(expectedText, range.Text ?? string.Empty);

        Font2 font = scope.Track(range.Font);
        ParagraphFormat2 paragraph = scope.Track(range.ParagraphFormat);
        AssertPoint($"{id} font size", _changedFontSizePt, font.Size);
        Assert.Equal(MsoParagraphAlignment.msoAlignCenter, paragraph.Alignment);
        Assert.Equal(_changedTextRgb, TextColourOf(scope, shape));

        Assert.Equal(id, shape.Name);
        Assert.Equal(ShapeOwnershipTag.ForPrimitiveId(id), shape.AlternativeText);
    }

    private static int TextColourOf(OfficeFixture.ComScope scope, Excel.Shape shape)
    {
        TextRange2 range = scope.Track(shape.TextFrame2.TextRange);
        Font2 font = scope.Track(range.Font);
        FillFormat fill = scope.Track(font.Fill);
        ColorFormat fore = scope.Track(fill.ForeColor);
        return fore.RGB;
    }

    private static OfficeShapeRequest Bar(
        string id,
        double y,
        string fillHex = "#92D050",
        string strokeHex = "#548235",
        double widthPt = 0.75d) =>
        new(
            id,
            OfficeShapeKind.Rectangle,
            new OfficeShapeGeometry(Bounds: new RectD(40d, y, 180d, 22d)),
            ZLayer.ActivityBody,
            FillColour: ColourHex.Parse(fillHex),
            StrokeColour: ColourHex.Parse(strokeHex),
            LineWidthPt: widthPt);

    private static OfficeShapeRequest Label(
        string id,
        string text,
        double y,
        string textHex = "#000000",
        double fontSizePt = 9d) =>
        new(
            id,
            OfficeShapeKind.TextBox,
            new OfficeShapeGeometry(Bounds: new RectD(40d, y, 120d, 16d)),
            ZLayer.Label,
            FontFamily: _tokenFont,
            FontSizePt: fontSizePt,
            Bold: false,
            Text: text,
            Alignment: GanttTextAlignment.Centre,
            TextColour: ColourHex.Parse(textHex));

    private static void AssertPoint(string what, double expected, double actual)
    {
        Assert.True(
            GeometryMath.ApproximatelyEqual(expected, actual),
            string.Create(
                CultureInfo.InvariantCulture,
                $"{what}: expected {expected:R} but the live host reported {actual:R}; "
                    + $"difference {Math.Abs(expected - actual):R} exceeds the documented "
                    + $"tolerance {GeometryMath.Epsilon:R}."));
    }
}
