using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using GanttCreator.Office;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// R4.6's Required Office gate: the complete style-token matrix is rendered into
/// a real workbook and every fill, stroke, width, and hatch is read back off the
/// live host and compared against the entity guide's token tables.
/// </summary>
/// <remarks>
/// <para>
/// The expected values are written as literals transcribed from the guide, not
/// derived from the mapper under test. A live test that computed its expectation
/// through <see cref="OfficeStyleMapper"/> would prove only that the adapter
/// called the mapper; it could not catch a wrong byte order, a swapped red and
/// blue, or a hatch the host silently refused.
/// </para>
/// <para>
/// Everything is asserted from the <em>host's</em> reported values, read back off
/// the shape after the write. A host that normalises, quantises, or ignores a
/// property therefore fails the test rather than being assumed to have taken it.
/// </para>
/// <para>
/// All the claims share one Excel session on purpose: the COM-proxy ratchet
/// counts live proxies, and opening a second instance for the hatch case would
/// have raised the signal against a ceiling the ratchet documentation forbids
/// raising without evidence.
/// </para>
/// </remarks>
public class StyleRenderIntegrationTests(ITestOutputHelper output)
{
    private const double PaddingPt = 12d;

    private readonly ITestOutputHelper _output = output;

    private static string XllPath => OfficeFixtureTests.ResolvePackedXllPath();

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Every_style_token_reaches_the_live_host_as_the_exact_property_value()
    {
        var fixture = new OfficeFixture();

        // KNOWN-LIMITATIONS L19, for the same reason the other render tests need
        // it: this test creates REAL shapes through the real writer, and a
        // workbook holding live shapes keeps the Excel process alive, so teardown
        // escalates to a kill. The test still asserts every style value, so
        // nothing is hidden by the suppression.
        fixture.SuppressLeakSignal = true;

        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(fixture.RegisterXll(XllPath),
                $"Application.RegisterXLL returned false for '{XllPath}'.");

            using var scope = new OfficeFixture.ComScope();
            Excel.Workbook workbook = scope.Track(fixture.CreateWorkbook());
            var initialiser = new ExcelWorkbookInitialiser(fixture.Excel);
            WorkbookInitialiseOutcome initialised = initialiser.Initialise();
            Assert.True(initialised.Succeeded, $"Initialise refused: {initialised.Refusal}");

            var renderer = new SceneShapeRenderer(
                ChartOriginDelta.ForChartBounds(new RectD(-PaddingPt, -PaddingPt, 400d, 300d)));
            var writer = new ExcelShapeWriter(fixture.Excel);

            foreach (StyleCase styleCase in Cases)
            {
                GanttScene scene = BuildScene(styleCase);
                SceneTranslationOutcome translated = renderer.Translate(scene);
                Assert.True(translated.Complete,
                    $"{styleCase.Id}: {translated.Deferred.Count} deferred, {translated.Refusals.Count} refused.");

                foreach (OfficeShapeRequest request in translated.Requests)
                {
                    ShapeWriteOutcome outcome = writer.Create(request);
                    Assert.True(outcome.Succeeded,
                        $"{styleCase.Id}: Create refused: {outcome.Refusal}");
                }

                Excel.Worksheet sheet = (Excel.Worksheet)scope.Track(
                    workbook.Sheets[GanttWorkbookContract.GanttSheetLabel]);
                AssertStyle(scope, sheet, styleCase);
            }

            // Delete the shapes but leave the workbook open for the fixture to
            // close, as the other render tests do.
            Excel.Worksheet ganttSheet = (Excel.Worksheet)scope.Track(
                workbook.Sheets[GanttWorkbookContract.GanttSheetLabel]);
            foreach (StyleCase styleCase in Cases)
            {
                scope.Track(ganttSheet.Shapes.Item(styleCase.Id)).Delete();
            }
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    private void AssertStyle(OfficeFixture.ComScope scope, Excel.Worksheet sheet, StyleCase styleCase)
    {
        Excel.Shape shape = scope.Track(sheet.Shapes.Item(styleCase.Id));

        Excel.FillFormat fill = shape.Fill;
        Excel.LineFormat line = shape.Line;

        _output.WriteLine($"{styleCase.Id}: fill={Describe(fill.ForeColor?.RGB ?? -1)} "
            + $"line={Describe(line.ForeColor?.RGB ?? -1)} weight={line.Weight} "
            + $"pattern={fill.Pattern} fillVisible={fill.Visible} lineVisible={line.Visible}");

        if (styleCase.ExpectedFillRgb is { } expectedFill)
        {
            // The fill TYPE is asserted per case, not assumed. A solid fill
            // reports msoFillSolid with Pattern = msoPatternMixed, so asserting
            // "patterned" for every filled shape would be wrong; asserting
            // "solid" for a hatched one would be worse, because it would pass on
            // a host that silently dropped the hatch. The hatch row below is the
            // one that pins the pattern, and it asserts patterned.
            Assert.Equal(
                styleCase.ExpectedPattern is null ? MsoFillType.msoFillSolid : MsoFillType.msoFillPatterned,
                fill.Type);
            Assert.Equal(expectedFill, fill.ForeColor?.RGB);
            Assert.Equal(MsoTriState.msoTrue, fill.Visible);
        }
        else
        {
            // The scene resolved no fill, so the host must report it switched off
            // rather than keeping whatever default AddShape gave the shape.
            Assert.Equal(MsoTriState.msoFalse, fill.Visible);
        }

        if (styleCase.ExpectedStrokeRgb is { } expectedStroke)
        {
            Assert.Equal(expectedStroke, line.ForeColor?.RGB);
            Assert.Equal(MsoTriState.msoTrue, line.Visible);
        }
        else
        {
            Assert.Equal(MsoTriState.msoFalse, line.Visible);
        }

        if (styleCase.ExpectedWidthPt is { } expectedWidth)
        {
            AssertPoint($"{styleCase.Id} line weight", expectedWidth, line.Weight);
        }

        if (styleCase.ExpectedPattern is { } expectedPattern)
        {
            // The host accepted the request and the pattern round-tripped: this
            // is the live half of the Step-0 obligation, since the contract tests
            // can only prove which constant the adapter asked for. Without this
            // the hatch row would pass on any host that quietly ignored
            // Patterned() and left the shape solid.
            Assert.Equal(MsoFillType.msoFillPatterned, fill.Type);
            Assert.Equal(expectedPattern, fill.Pattern);
        }
    }

    private static string Describe(int rgb) =>
        string.Create(CultureInfo.InvariantCulture, $"0x{rgb:X6}");

    /// <summary>
    /// One row of the live matrix: a scene style, and the host property values the
    /// entity guide's tables say it must produce. The expected integers are
    /// transcribed from the guide's <c>#RRGGBB</c> tokens and packed by hand into
    /// the host's <c>blue * 65536 + green * 256 + red</c> order.
    /// </summary>
    private sealed record StyleCase(
        string Id,
        SceneStyle Style,
        int? ExpectedFillRgb,
        int? ExpectedStrokeRgb,
        double? ExpectedWidthPt,
        MsoPatternType? ExpectedPattern,
        bool IsDiamond = false,
        bool IsLine = false);

    private static IReadOnlyList<StyleCase> Cases =>
    [
        // Section 13: as-planned. #92D050 -> 0x50D092, #548235 -> 0x358254.
        new(
            "row-1:planned",
            new SceneStyle("AsPlannedActivity", ColourHex.Parse("#92D050"), ColourHex.Parse("#548235"), 0.75),
            0x50D092,
            0x358254,
            0.75,
            null),

        // Section 14: as-built. #00B0F0 -> 0xF0B000, #0070C0 -> 0xC07000.
        new(
            "row-1:actual",
            new SceneStyle("AsBuiltActivity", ColourHex.Parse("#00B0F0"), ColourHex.Parse("#0070C0"), 0.75),
            0xF0B000,
            0xC07000,
            0.75,
            null),

        // Section 15: baseline. #00B050 -> 0x50B000, #006100 is symmetric.
        new(
            "row-1:baseline",
            new SceneStyle("BaselineActivity", ColourHex.Parse("#00B050"), ColourHex.Parse("#006100"), 0.75),
            0x50B000,
            0x006100,
            0.75,
            null),

        // Section 16: critical interval. No fill, #FF0000 -> 0x0000FF,
        // CriticalLinePt 2.25.
        new(
            "row-1:critical",
            new SceneStyle("CriticalInterval", strokeColour: ColourHex.Parse("#FF0000"), outlineWidthPt: 2.25),
            null,
            0x0000FF,
            2.25,
            null,
            IsLine: true),

        // Section 18: procurement, as-planned family. The hatch is the point of
        // this row: the host must actually report the pattern it was asked for.
        new(
            "row-1:procurement",
            new SceneStyle(
                "AsPlannedProcurement",
                ColourHex.Parse("#92D050"),
                ColourHex.Parse("#548235"),
                0.75,
                GanttHatchPattern.ForwardDiagonal),
            0x50D092,
            0x358254,
            0.75,
            MsoPatternType.msoPatternLightDownwardDiagonal),

        // Section 10: splitter band. #FFE699 -> 0x99E6FF, no stroke.
        new(
            "row-1:splitter",
            new SceneStyle("Splitter", ColourHex.Parse("#FFE699")),
            0x99E6FF,
            null,
            null,
            null),

        // Section 11: spacer. No fill, no stroke - the host must report both off.
        new(
            "row-1:spacer",
            new SceneStyle("Spacer"),
            null,
            null,
            null,
            null),

        // Section 21: critical milestone. #FF0000 -> 0x0000FF,
        // #C00000 -> 0x0000C0.
        new(
            "row-1:milestone",
            new SceneStyle("CriticalMilestone", ColourHex.Parse("#FF0000"), ColourHex.Parse("#C00000"), 0.75),
            0x0000FF,
            0x0000C0,
            0.75,
            null,
            IsDiamond: true),

        // Section 24: delineator. #404040 is symmetric, DelineatorLinePt 0.75.
        new(
            "row-1:delineator",
            new SceneStyle("DefaultDelineator", strokeColour: ColourHex.Parse("#404040"), outlineWidthPt: 0.75),
            null,
            0x404040,
            0.75,
            null,
            IsLine: true),
    ];

    private static GanttScene BuildScene(StyleCase styleCase)
    {
        ScenePrimitive primitive = styleCase.IsLine
            ? new SceneLine(
                styleCase.Id,
                SceneOwnerId.Chart,
                ZLayer.CriticalOverlay,
                new PointD(40d, 60d),
                new PointD(220d, 60d),
                styleCase.Style)
            : styleCase.IsDiamond
                ? new ScenePolygon(
                    styleCase.Id,
                    SceneOwnerId.Chart,
                    ZLayer.Milestone,
                    [
                        new PointD(130d, 38d),
                        new PointD(134d, 42d),
                        new PointD(130d, 46d),
                        new PointD(126d, 42d),
                    ],
                    styleCase.Style)
                : new SceneRect(
                    styleCase.Id,
                    SceneOwnerId.Chart,
                    ZLayer.ActivityBody,
                    new RectD(40d, 30d, 180d, 24d),
                    styleCase.Style);

        return GanttScene
            .TryCreate(
                new RectD(-PaddingPt, -PaddingPt, 400d, 300d),
                new RectD(0d, 0d, 380d, 280d),
                [primitive],
                [])
            .Scene ?? throw new InvalidOperationException($"The {styleCase.Id} scene failed to validate.");
    }

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
