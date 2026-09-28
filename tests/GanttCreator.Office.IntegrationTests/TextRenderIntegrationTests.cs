using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using GanttCreator.Office;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeCore = Microsoft.Office.Core;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// R4.4's Required Office gate: a scene label becomes a real text box whose
/// content, font, and alignment are read back off the live host.
/// </summary>
/// <remarks>
/// <para>
/// The expected values are hand-written from the scene constants, never derived
/// from <see cref="SceneShapeRenderer"/>. A test whose expectation comes from the
/// code under test is self-consistent by construction: mutating the translation
/// moves expectation and observation together and the test still passes. That
/// defect was found and fixed in the R4.3 gate and the same rule applies here.
/// </para>
/// <para>
/// The overflow case carries a string the scene already truncated. The gate
/// asserts it arrives byte-for-byte, so a second truncation by the renderer or a
/// silent host-side re-wrap cannot pass.
/// </para>
/// </remarks>
public class TextRenderIntegrationTests(ITestOutputHelper output)
{
    private const double ChartPaddingPt = 12d;
    private const string TokenFont = "Aptos";
    private const double FontSizePt = 9d;

    private readonly ITestOutputHelper _output = output;

    private static string XllPath => OfficeFixtureTests.ResolvePackedXllPath();

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task A_scene_label_renders_with_its_font_alignment_and_overflow_policy()
    {
        var fixture = new OfficeFixture();

        // L19: this test creates REAL text shapes through the real
        // ExcelShapeWriter, and a workbook holding a live shape keeps the Excel
        // process alive, so teardown legitimately escalates to a kill. The
        // harness provides this flag for exactly that case; it is not a licence
        // to suppress a test that simply leaked.
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

            const string Normal = "Site survey";
            const string Clipped = "Site survey and ground investigation phase one…";
            GanttScene scene = BuildScene(Normal, Clipped, out RectD chartBounds);

            var renderer = new SceneShapeRenderer(ChartOriginDelta.ForChartBounds(chartBounds));
            SceneTranslationOutcome translated = renderer.Translate(scene);
            Assert.True(translated.Complete, translated.Deferred.Count + " primitives were deferred.");

            var writer = new ExcelShapeWriter(fixture.Excel);
            foreach (OfficeShapeRequest request in translated.Requests)
            {
                ShapeWriteOutcome outcome = writer.Create(request);
                Assert.True(outcome.Succeeded, $"Create refused for '{request.PrimitiveId}': {outcome.Refusal}");
            }

            var sheet = (Excel.Worksheet)scope.Track(
                workbook.Sheets[GanttWorkbookContract.GanttSheetLabel]);

            Excel.Shape normal = scope.Track(sheet.Shapes.Item("row-1:label"));
            Excel.Shape clipped = scope.Track(sheet.Shapes.Item("row-2:clipped"));

            AssertGeometry(normal, clipped);
            AssertContent(normal, clipped, Normal, Clipped);
            AssertTypography(scope, normal, clipped);

            // Delete the shapes but leave the workbook open for the fixture to
            // close, as the R4.3 gate does.
            normal.Delete();
            clipped.Delete();
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    private static void AssertGeometry(Excel.Shape normal, Excel.Shape clipped)
    {
        // The same uniform delta every primitive family takes: scene (40, 30)
        // translated by the 12pt padding, and scene (40, 50) likewise.
        AssertPoint("normal label Left", 52d, normal.Left);
        AssertPoint("normal label Top", 42d, normal.Top);
        AssertPoint("clipped label Left", 52d, clipped.Left);
        AssertPoint("clipped label Top", 62d, clipped.Top);

        Assert.Equal("row-1:label", normal.Name);
        Assert.Equal(ShapeOwnershipTag.ForPrimitiveId("row-1:label"), normal.AlternativeText);
        Assert.Equal("row-2:clipped", clipped.Name);
        Assert.Equal(ShapeOwnershipTag.ForPrimitiveId("row-2:clipped"), clipped.AlternativeText);
    }

    private static void AssertContent(Excel.Shape normal, Excel.Shape clipped, string normalText, string clippedText)
    {
        // Verbatim. A second truncation by the renderer, or a host-side re-wrap
        // that dropped the tail, would show up as a different string here.
        Assert.Equal(normalText, TextOf(normal));
        Assert.Equal(clippedText, TextOf(clipped));
        Assert.EndsWith("…", TextOf(clipped), StringComparison.Ordinal);
    }

    private void AssertTypography(
        OfficeFixture.ComScope scope,
        Excel.Shape normal,
        Excel.Shape clipped)
    {
        OfficeCore.TextRange2 normalRange = scope.Track((OfficeCore.TextRange2)normal.TextFrame2.TextRange);
        OfficeCore.TextRange2 clippedRange = scope.Track((OfficeCore.TextRange2)clipped.TextFrame2.TextRange);

        _output.WriteLine("font family read back: " + normalRange.Font.Name);
        _output.WriteLine("font size read back  : " + normalRange.Font.Size);
        _output.WriteLine("normal alignment    : " + normalRange.ParagraphFormat.Alignment);
        _output.WriteLine("clipped alignment   : " + clippedRange.ParagraphFormat.Alignment);

        // R8.2a / D-G4 input: a host that substituted a missing font shows here.
        // It is REPORTED rather than silently accepted, because R4.4's stop
        // condition names an unreported substitution explicitly.
        string actualFamily = normalRange.Font.Name;
        if (!string.Equals(actualFamily, TokenFont, StringComparison.OrdinalIgnoreCase))
        {
            _output.WriteLine(
                "TOKEN FONT NOT PRESENT: the host reported '" + actualFamily + "' for requested '"
                    + TokenFont + "'. This is the R8.2a font-pinning input (D-G4), not a silent pass.");
        }

        AssertPoint("font size", FontSizePt, normalRange.Font.Size);
        Assert.Equal(
            Microsoft.Office.Core.MsoParagraphAlignment.msoAlignLeft,
            normalRange.ParagraphFormat.Alignment);
        Assert.Equal(
            Microsoft.Office.Core.MsoParagraphAlignment.msoAlignRight,
            clippedRange.ParagraphFormat.Alignment);
    }

    private static GanttScene BuildScene(string normal, string clipped, out RectD chartBounds)
    {
        chartBounds = new RectD(-ChartPaddingPt, -ChartPaddingPt, 400d, 300d);
        var style = new SceneStyle("Label", fontFamily: TokenFont, fontSizePt: FontSizePt, bold: false);

        return GanttScene.TryCreate(
                chartBounds,
                new RectD(0d, 0d, 380d, 280d),
                [
                    new SceneText(
                        "row-1:label",
                        SceneOwnerId.Chart,
                        ZLayer.Label,
                        normal,
                        new RectD(40d, 30d, 60d, 12d),
                        style,
                        GanttTextAlignment.Left),
                    new SceneText(
                        "row-2:clipped",
                        SceneOwnerId.Chart,
                        ZLayer.Label,
                        clipped,
                        new RectD(40d, 50d, 40d, 12d),
                        style,
                        GanttTextAlignment.Right),
                ],
                [])
            .Scene ?? throw new InvalidOperationException("The fixture scene failed to validate.");
    }

    private static string TextOf(Excel.Shape shape) =>
        ((OfficeCore.TextRange2)shape.TextFrame2.TextRange).Text ?? string.Empty;

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
