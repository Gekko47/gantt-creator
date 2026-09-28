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
            AssertUpdateReappliesContent(writer, scope, sheet);

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

    /// <summary>
    /// Proves on the live host that an update re-writes a label's content, not
    /// just its position. R4.7 reaches this path on every reconcile that finds
    /// an existing shape, so a create-only content path would leave a stale
    /// string sitting in the right place - a silent defect with no error.
    /// </summary>
    /// <param name="writer">The shape writer under test.</param>
    /// <param name="scope">The COM scope.</param>
    /// <param name="sheet">The worksheet holding the shape.</param>
    private static void AssertUpdateReappliesContent(
        ExcelShapeWriter writer,
        OfficeFixture.ComScope scope,
        Excel.Worksheet sheet)
    {
        const string Revised = "Revised after refresh";
        const double RevisedSize = 12d;

        ShapeWriteOutcome outcome = writer.Update(
            new OfficeShapeRequest(
                "row-1:label",
                OfficeShapeKind.TextBox,
                new OfficeShapeGeometry(Bounds: new RectD(40d, 30d, 90d, 16d)),
                ZLayer.Label,
                FontFamily: TokenFont,
                FontSizePt: RevisedSize,
                Text: Revised,
                Alignment: GanttTextAlignment.Centre));

        Assert.True(outcome.Succeeded, $"Update refused: {outcome.Refusal}");

        // Re-read from the host. A separate fetch proves the write landed on the
        // live shape rather than on a cached proxy value.
        Excel.Shape reread = scope.Track(sheet.Shapes.Item("row-1:label"));
        string text = TextOf(reread);

        Assert.Equal(Revised, text);
        AssertPoint("updated label Width", 90d, reread.Width);
        Assert.Equal(
            Microsoft.Office.Core.MsoParagraphAlignment.msoAlignCenter,
            ((OfficeCore.TextRange2)reread.TextFrame2.TextRange).ParagraphFormat.Alignment);
        AssertPoint("updated font size", RevisedSize, ((OfficeCore.TextRange2)reread.TextFrame2.TextRange).Font.Size);

        // The name and the ownership tag survive an update untouched: they are
        // the reconciliation key and the ownership proof, and re-stamping them
        // would silently "repair" a user-edited shape.
        Assert.Equal("row-1:label", reread.Name);
        Assert.Equal(ShapeOwnershipTag.ForPrimitiveId("row-1:label"), reread.AlternativeText);
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

    /// <summary>
    /// Proves on the live host that a label's text colour is written to the FONT's
    /// fill and read back off the real shape, and that an update restyles it.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This is the gate that could not be written from documentation alone. The
    /// write goes through <c>TextRange2.Font.Fill.ForeColor.RGB</c>, and the
    /// reflection probe established only that those members EXIST - not that the
    /// host honours a colour written to them, or that the fill must be made solid
    /// first for the write to land. A host that quietly ignored the write would
    /// leave the label in Excel's default font colour, which is exactly the
    /// failure entity guide section 17 forbids.
    /// </para>
    /// <para>
    /// The two colours are #FFFFFF and #123456. White is the section 17 inside
    /// value but is a poor byte-order witness (its channels are identical), so
    /// #123456 - three distinct bytes, packed to 0x563412 - is what actually pins
    /// the channel order. Both are asserted.
    /// </para>
    /// <para>
    /// The expected integers were confirmed by a live probe across five tokens
    /// (#123456, #FF0000, #00FF00, #0000FF, #010203) on this host build, which
    /// showed the FONT fill reporting the same packed value as the shape fill for
    /// every one. So the text path and the fill path share one byte order, and the
    /// asymmetry that would have broken this test is a test-helper bug, not a
    /// production one - it is recorded here because the first run of this test did
    /// fail on exactly that, and the failure was in the test.
    /// </para>
    /// </remarks>
    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task A_label_text_colour_reaches_the_live_font_and_survives_an_update()
    {
        var fixture = new OfficeFixture();

        // L19, for the same reason the test above needs it: this creates real
        // shapes through the real writer, so teardown escalates to a kill.
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

            const string InsideHex = "#FFFFFF";
            const string DistinctHex = "#123456";
            const int InsideRgb = 0xFFFFFF;
            const int DistinctRgb = 0x563412;

            var writer = new ExcelShapeWriter(fixture.Excel);
            foreach ((string id, string hex) in new[] { ("row-1:inside", InsideHex), ("row-2:outside", DistinctHex) })
            {
                ShapeWriteOutcome created = writer.Create(LabelRequest(id, hex));
                Assert.True(created.Succeeded, $"Create refused for '{id}': {created.Refusal}");
            }

            var sheet = (Excel.Worksheet)scope.Track(
                workbook.Sheets[GanttWorkbookContract.GanttSheetLabel]);

            Excel.Shape inside = scope.Track(sheet.Shapes.Item("row-1:inside"));
            Excel.Shape outside = scope.Track(sheet.Shapes.Item("row-2:outside"));

            _output.WriteLine("inside font colour : " + Describe(TextColourOf(scope, inside)));
            _output.WriteLine("outside font colour: " + Describe(TextColourOf(scope, outside)));

            // Read back off the HOST, not from the request: a host that ignored
            // the write, or normalised it, fails here rather than being assumed
            // to have taken it.
            Assert.Equal(InsideRgb, TextColourOf(scope, inside));
            Assert.Equal(DistinctRgb, TextColourOf(scope, outside));

            // The shape's own fill must be unaffected. The text colour is a FONT
            // property; if it had leaked onto the shape fill these two shapes -
            // created with no fill at all - would have gained one.
            Assert.Equal(OfficeCore.MsoTriState.msoFalse, inside.Fill.Visible);
            Assert.Equal(OfficeCore.MsoTriState.msoFalse, outside.Fill.Visible);

            AssertTextColourSurvivesAnUpdate(writer, scope, sheet);

            inside.Delete();
            outside.Delete();
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    /// <summary>
    /// The update half of the text-colour gate, run against shapes the create half
    /// already placed. R4.7 reaches this path on every reconcile that finds an
    /// existing shape, so a create-only colour path would leave the old colour
    /// sitting in the right place - a silent defect with no error.
    /// </summary>
    private void AssertTextColourSurvivesAnUpdate(
        ExcelShapeWriter writer,
        OfficeFixture.ComScope scope,
        Excel.Worksheet sheet)
    {
        // The two shapes swap colours, so a create-only path cannot pass.
        Assert.True(
            writer.Update(LabelRequest("row-1:inside", "#123456")).Succeeded,
            "Update refused.");
        Assert.True(
            writer.Update(LabelRequest("row-2:outside", "#FFFFFF")).Succeeded,
            "Update refused.");

        Excel.Shape restyledInside = scope.Track(sheet.Shapes.Item("row-1:inside"));
        Excel.Shape restyledOutside = scope.Track(sheet.Shapes.Item("row-2:outside"));

        _output.WriteLine("restyled inside : " + Describe(TextColourOf(scope, restyledInside)));
        _output.WriteLine("restyled outside: " + Describe(TextColourOf(scope, restyledOutside)));

        Assert.Equal(0x563412, TextColourOf(scope, restyledInside));
        Assert.Equal(0xFFFFFF, TextColourOf(scope, restyledOutside));
    }

    /// <summary>
    /// Reads the live font colour through the same chain the writer writes it, so
    /// the assertion is against the host's own reported value.
    /// </summary>
    private static int TextColourOf(OfficeFixture.ComScope scope, Excel.Shape shape)
    {
        OfficeCore.TextRange2 range = scope.Track((OfficeCore.TextRange2)shape.TextFrame2.TextRange);
        OfficeCore.Font2 font = scope.Track(range.Font);
        OfficeCore.FillFormat fill = scope.Track(font.Fill);
        OfficeCore.ColorFormat fore = scope.Track(fill.ForeColor);
        return fore.RGB;
    }

    /// <summary>
    /// One label request carrying the authored <c>#RRGGBB</c> token, exactly as the
    /// scene would hand it over.
    /// </summary>
    /// <param name="primitiveId">The stable scene primitive identifier.</param>
    /// <param name="textHex">The authored colour token, for example <c>#123456</c>.</param>
    /// <returns>The request.</returns>
    private static OfficeShapeRequest LabelRequest(string primitiveId, string textHex) =>
        new(
            primitiveId,
            OfficeShapeKind.TextBox,
            new OfficeShapeGeometry(Bounds: new RectD(40d, 30d, 90d, 16d)),
            ZLayer.Label,
            FontFamily: TokenFont,
            FontSizePt: FontSizePt,
            Text: "Delay",
            Alignment: GanttTextAlignment.Centre,
            TextColour: ColourHex.Parse(textHex));

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

    /// <summary>
    /// Formats a host colour as 0xRRGGBB for the test log, invariantly so the
    /// output does not vary with the machine's locale.
    /// </summary>
    private static string Describe(int rgb) =>
        string.Create(CultureInfo.InvariantCulture, $"0x{rgb:X6}");

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
