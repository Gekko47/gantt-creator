using GanttCreator.Core;
using GanttCreator.Core.Scene;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Moq;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R4.6's writer half: the resolved style actually reaches the host's fill and
/// line members, on the create path and the update path, for every shape family.
/// </summary>
/// <remarks>
/// <para>
/// Unlike the rest of <see cref="ExcelShapeWriterTests"/>, these tests do
/// <em>not</em> override the <c>ApplyStyle</c> seam. The seam exists so a test can
/// observe that the call happened; these tests instead let the real implementation
/// run against a shape that models <c>Fill</c> and <c>Line</c>, because the
/// question here is what the host was <em>written</em>, not whether a method was
/// invoked. A test that only recorded the call would pass unchanged if the body
/// wrote nothing.
/// </para>
/// <para>
/// The fake models the format objects as state, not as fixed returns, so an
/// assertion that the writer set a colour reads the value back off the host the
/// way a live shape would report it.
/// </para>
/// </remarks>
public class StyleWriteTests
{
    /// <summary>
    /// One modelled host shape: its fill and line formats, and the writes the
    /// adapter made to them, recorded in order so the
    /// <c>Solid()</c>/<c>Patterned()</c> sequence is assertable.
    /// </summary>
    private sealed class HostShape
    {
        public HostShape(Excel.FillFormat? fillOverride = null, Excel.LineFormat? lineOverride = null)
        {
            var fill = fillOverride ?? new Mock<Excel.FillFormat>().Object;
            var line = lineOverride ?? new Mock<Excel.LineFormat>().Object;
            Mock<Excel.ColorFormat>? foreColour = null;
            Mock<Excel.ColorFormat>? backColour = null;
            Mock<Excel.ColorFormat>? lineColour = null;

            if (fillOverride is null)
            {
                foreColour = new Mock<Excel.ColorFormat>();
                backColour = new Mock<Excel.ColorFormat>();
                var fillMock = (Mock<Excel.FillFormat>)Mock.Get(fill);
                _ = fillMock.SetupGet(f => f.ForeColor).Returns(foreColour.Object);
                _ = fillMock.SetupGet(f => f.BackColor).Returns(backColour.Object);
                _ = fillMock.Setup(f => f.Solid()).Callback(() => FillCalls.Add("Solid"));
                _ = fillMock.Setup(f => f.Patterned(It.IsAny<MsoPatternType>()))
                    .Callback<MsoPatternType>(pattern => FillCalls.Add($"Patterned:{pattern}"));
                _ = fillMock.SetupGet(f => f.Visible).Returns(() => FillVisible);
                _ = fillMock.SetupSet(f => f.Visible = It.IsAny<MsoTriState>())
                    .Callback<MsoTriState>(value => FillVisible = value);
                _ = fillMock.SetupGet(f => f.Transparency).Returns(() => FillTransparency);
                _ = fillMock.SetupSet(f => f.Transparency = It.IsAny<float>())
                    .Callback<float>(value => FillTransparency = value);
                _ = foreColour.SetupGet(c => c.RGB).Returns(() => FillRgb);
                _ = foreColour.SetupSet(c => c.RGB = It.IsAny<int>())
                    .Callback<int>(value => FillRgb = value);
                _ = backColour.SetupGet(c => c.RGB).Returns(() => FillBackRgb);
                _ = backColour.SetupSet(c => c.RGB = It.IsAny<int>())
                    .Callback<int>(value => FillBackRgb = value);
            }

            if (lineOverride is null)
            {
                lineColour = new Mock<Excel.ColorFormat>();
                var lineMock = (Mock<Excel.LineFormat>)Mock.Get(line);
                _ = lineMock.SetupGet(l => l.ForeColor).Returns(lineColour.Object);
                _ = lineMock.SetupGet(l => l.Weight).Returns(() => LineWeight);
                _ = lineMock.SetupSet(l => l.Weight = It.IsAny<float>())
                    .Callback<float>(value => LineWeight = value);
                _ = lineMock.SetupGet(l => l.Visible).Returns(() => LineVisible);
                _ = lineMock.SetupSet(l => l.Visible = It.IsAny<MsoTriState>())
                    .Callback<MsoTriState>(value => LineVisible = value);
                _ = lineMock.SetupGet(l => l.Transparency).Returns(() => LineTransparency);
                _ = lineMock.SetupSet(l => l.Transparency = It.IsAny<float>())
                    .Callback<float>(value => LineTransparency = value);
                _ = lineColour.SetupGet(c => c.RGB).Returns(() => LineRgb);
                _ = lineColour.SetupSet(c => c.RGB = It.IsAny<int>())
                    .Callback<int>(value => LineRgb = value);
            }

            var shape = new Mock<Excel.Shape>();
            _ = shape.SetupGet(s => s.Fill).Returns(fill);
            _ = shape.SetupGet(s => s.Line).Returns(line);
            _ = shape.SetupGet(s => s.Name).Returns("row-1:bar");
            _ = shape.SetupGet(s => s.AlternativeText)
                .Returns(() => ShapeOwnershipTag.ForPrimitiveId("row-1:bar"));
            _ = shape.Setup(s => s.Delete()).Callback(() => Deleted = true);

            // TextFrame2 is the first member ApplyText reads, so a plain mock is
            // enough for the paths where the content write succeeds.
            var frame = new Mock<Excel.TextFrame2>();
            _ = shape.SetupGet(s => s.TextFrame2).Returns(frame.Object);

            Shape = shape.Object;
        }

        public Excel.Shape Shape { get; }

        /// <summary>
        /// Makes the text-content write throw, simulating a host that refuses it.
        /// The text box is the one kind whose content is applied inside the
        /// creation call, so this is the create-path content failure the adapter
        /// has to contain. A method rather than a flag because the mock setup has
        /// to be re-applied after the shape exists, and a later Moq setup
        /// replaces an earlier one for the same member.
        /// </summary>
        public void FailTextWrite()
        {
            var mock = (Mock<Excel.Shape>)Mock.Get(Shape);
            _ = mock.SetupGet(s => s.TextFrame2)
                .Throws(new InvalidOperationException("The host refused the text."));
        }

        public List<string> FillCalls { get; } = [];

        public bool Deleted { get; private set; }

        public int FillRgb { get; private set; }

        public int FillBackRgb { get; private set; }

        public float FillTransparency { get; private set; }

        public MsoTriState FillVisible { get; private set; }

        public int LineRgb { get; private set; }

        public float LineWeight { get; private set; }

        public float LineTransparency { get; private set; }

        public MsoTriState LineVisible { get; private set; }
    }

    /// <summary>
    /// A writer whose shape creation returns the modelled host shape. No styling
    /// seam is overridden, so the real <c>ApplyStyle</c> runs.
    /// </summary>
    /// <param name="host">The modelled host shape to return from the creation seam.</param>
    /// <param name="shapeAlreadyExists">
    /// Whether the lookup seam reports the shape as already present. <c>Create</c>
    /// needs it absent (otherwise it correctly refuses <c>AlreadyExists</c>);
    /// <c>Update</c> needs it present.
    /// </param>
    private sealed class StyleWriter(HostShape host, bool shapeAlreadyExists = false)
        : ExcelShapeWriter(ActiveApplication(), new ClearGuard())
    {
        /// <summary>The ownership members stamped, one entry per stamp.</summary>
        public List<string> OwnershipStamps { get; } = [];

        internal override Excel._Worksheet? FindGanttWorksheet(Excel.Sheets sheets) =>
            new Mock<Excel._Worksheet>().Object;

        internal override Excel.Shapes? GetShapes(Excel._Worksheet sheet) =>
            new Mock<Excel.Shapes>().Object;

        internal override Excel.Shape? AddShape(Excel.Shapes shapes, OfficeShapeRequest request) =>
            host.Shape;

        internal override Excel.Shape? FindShapeByName(Excel.Shapes shapes, string name) =>
            shapeAlreadyExists ? host.Shape : null;

        internal override void ApplyOwnershipStamp(Excel.Shape shape, string primitiveId)
        {
            OwnershipStamps.Add(primitiveId);
            base.ApplyOwnershipStamp(shape, primitiveId);
        }
    }

    /// <summary>
    /// A writer whose text box is created through the real <c>AddShape</c>, so
    /// the content write genuinely happens inside the creation call. Overriding
    /// <c>AddShape</c> to return the host shape directly, as
    /// <see cref="StyleWriter"/> does for the styling paths, would skip exactly
    /// the sequence under test.
    /// </summary>
    /// <param name="host">The modelled host shape the host hands back.</param>
    private sealed class TextBoxWriter : ExcelShapeWriter
    {
        private readonly Mock<Excel.Shapes> _shapes = new();

        /// <summary>Initialises a writer whose host hands back the modelled shape.</summary>
        /// <param name="host">The shape <c>AddTextbox</c> returns.</param>
        internal TextBoxWriter(HostShape host)
            : base(ActiveApplication(), new ClearGuard())
        {
            _ = _shapes.Setup(s => s.AddTextbox(
                    It.IsAny<MsoTextOrientation>(),
                    It.IsAny<float>(),
                    It.IsAny<float>(),
                    It.IsAny<float>(),
                    It.IsAny<float>()))
                .Returns(host.Shape);
        }

        /// <summary>The ownership members stamped, one entry per stamp.</summary>
        public List<string> OwnershipStamps { get; } = [];

        internal override Excel._Worksheet? FindGanttWorksheet(Excel.Sheets sheets) =>
            new Mock<Excel._Worksheet>().Object;

        internal override Excel.Shapes? GetShapes(Excel._Worksheet sheet) => _shapes.Object;

        internal override Excel.Shape? FindShapeByName(Excel.Shapes shapes, string name) => null;

        internal override void ApplyOwnershipStamp(Excel.Shape shape, string primitiveId)
        {
            OwnershipStamps.Add(primitiveId);
            base.ApplyOwnershipStamp(shape, primitiveId);
        }
    }

    private sealed class ClearGuard : IWorksheetProtectionGuard
    {
        public ProtectionGuardOutcome Query() => ProtectionGuardOutcome.NotProtected;

        public ProtectionGuardOutcome QueryTarget(object? target) => ProtectionGuardOutcome.NotProtected;
    }

    /// <summary>
    /// A minimal application the adapter can resolve a workbook from. The writer
    /// refuses with <c>NoActiveWorkbook</c> before it reaches any seam when there
    /// is none, so the styling paths under test need a resolvable one.
    /// </summary>
    private static object ActiveApplication()
    {
        var workbook = new Mock<Excel.Workbook>();
        _ = workbook.Setup(w => w.Sheets).Returns(new Mock<Excel.Sheets>().Object);
        var application = new Mock<Excel.Application>();
        _ = application.Setup(a => a.ActiveWorkbook).Returns(workbook.Object);
        return application.Object;
    }

    [Fact]
    public void Creating_a_planned_activity_writes_its_fill_outline_and_width_to_the_host()
    {
        var host = new HostShape();
        var writer = new StyleWriter(host);

        // Section 13's default style, transcribed from the guide's token table.
        ShapeWriteOutcome outcome = writer.Create(Request(
            fill: "#92D050",
            stroke: "#548235",
            width: 0.75));

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal(0x50D092, host.FillRgb);
        Assert.Equal(0x358254, host.LineRgb);
        Assert.Equal(0.75f, host.LineWeight);
        Assert.Equal(MsoTriState.msoTrue, host.FillVisible);
        Assert.Equal(MsoTriState.msoTrue, host.LineVisible);
        Assert.Equal("Solid", Assert.Single(host.FillCalls));
    }

    [Fact]
    public void Creating_a_critical_interval_writes_the_critical_stroke_at_its_token_width()
    {
        var host = new HostShape();
        var writer = new StyleWriter(host);

        // Section 16: solid CriticalStroke, CriticalLinePt thick, and no fill.
        ShapeWriteOutcome outcome = writer.Create(Request(
            fill: null,
            stroke: "#FF0000",
            width: 2.25,
            kind: OfficeShapeKind.Line));

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal(0x0000FF, host.LineRgb);
        Assert.Equal(2.25f, host.LineWeight);
        Assert.Equal(MsoTriState.msoFalse, host.FillVisible);
    }

    [Fact]
    public void A_procurement_activity_asks_the_host_for_the_diagonal_pattern()
    {
        var host = new HostShape();
        var writer = new StyleWriter(host);

        ShapeWriteOutcome outcome = writer.Create(Request(
            fill: "#00B0F0",
            stroke: "#0070C0",
            width: 0.75,
            hatch: GanttHatchPattern.ForwardDiagonal));

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal("Patterned:msoPatternLightDownwardDiagonal", Assert.Single(host.FillCalls));

        // The family colours survive; the pattern only changes how the fill draws.
        Assert.Equal(0xF0B000, host.FillRgb);
        Assert.Equal(0xC07000, host.LineRgb);
    }

    [Fact]
    public void A_shape_with_no_resolved_fill_is_switched_off_rather_than_left_at_the_host_default()
    {
        var host = new HostShape();
        var writer = new StyleWriter(host);

        // A spacer: section 11 draws no fill. The host would otherwise apply its
        // own default fill, and "no fill resolved" has to be able to mean "no fill".
        ShapeWriteOutcome outcome = writer.Create(Request(fill: null, stroke: null));

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal(MsoTriState.msoFalse, host.FillVisible);
        Assert.Equal(MsoTriState.msoFalse, host.LineVisible);
    }

    [Fact]
    public void An_explicit_alpha_token_reaches_the_host_as_a_transparency_value()
    {
        var host = new HostShape();
        var writer = new StyleWriter(host);

        ShapeWriteOutcome outcome = writer.Create(Request(fill: "#92D05080", stroke: null));

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal(1f - (128f / 255f), host.FillTransparency, 5);
    }

    [Fact]
    public void An_opaque_token_writes_zero_transparency_rather_than_leaving_the_host_value()
    {
        var host = new HostShape();
        var writer = new StyleWriter(host);

        writer.Create(Request(fill: "#92D050", stroke: "#548235", width: 0.75));

        Assert.Equal(0f, host.FillTransparency);
        Assert.Equal(0f, host.LineTransparency);
    }

    [Fact]
    public void An_update_reapplies_a_changed_colour_rather_than_only_moving_the_shape()
    {
        var host = new HostShape();
        var writer = new StyleWriter(host, shapeAlreadyExists: true);

        // The staleness class R4.4 closed for text and R4.6 closes for style: a
        // refresh that changed only a colour must not leave the old one behind.
        ShapeWriteOutcome outcome = writer.Update(Request(
            id: "row-1:bar",
            fill: "#00B050",
            stroke: "#006100",
            width: 0.75));

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal(0x50B000, host.FillRgb);
        Assert.Equal(0x006100, host.LineRgb);
    }

    [Fact]
    public void An_update_reapplies_a_changed_width_and_hatch()
    {
        var host = new HostShape();
        var writer = new StyleWriter(host, shapeAlreadyExists: true);

        ShapeWriteOutcome outcome = writer.Update(Request(
            id: "row-1:bar",
            fill: "#00B0F0",
            stroke: "#0070C0",
            width: 2.25,
            hatch: GanttHatchPattern.Cross));

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal(2.25f, host.LineWeight);
        Assert.Equal("Patterned:msoPatternDiagonalCross", Assert.Single(host.FillCalls));
    }

    [Fact]
    public void A_milestone_diamond_reaches_the_host_with_its_own_subtype_colours()
    {
        var host = new HostShape();
        var writer = new StyleWriter(host);

        // Section 21's critical milestone: CriticalStroke fill, CriticalOutline
        // stroke. Without the mapping this renders in the host's accent colour,
        // which erases the planned/actual/baseline/critical distinction entirely.
        ShapeWriteOutcome outcome = writer.Create(Request(
            id: "row-1:bar",
            fill: "#FF0000",
            stroke: "#C00000",
            width: 0.75,
            kind: OfficeShapeKind.Diamond,
            points: true));

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal(0x0000FF, host.FillRgb);
        Assert.Equal(0x0000C0, host.LineRgb);
    }

    [Fact]
    public void A_host_that_refuses_the_style_write_leaves_no_orphaned_shape()
    {
        // The same reasoning as the ownership stamp: a shape the caller is told
        // was not created must not remain on the sheet. The adapter discards it
        // and Create still reports the typed refusal the caller acts on.
        var fill = new Mock<Excel.FillFormat>();
        _ = fill.SetupGet(f => f.ForeColor)
            .Throws(new InvalidOperationException("The host refused the fill."));
        var host = new HostShape(fill.Object);
        var writer = new StyleWriter(host);

        ShapeWriteOutcome outcome = writer.Create(Request(fill: "#92D050", stroke: "#548235", width: 0.75));

        Assert.True(host.Deleted);
        Assert.Equal(ShapeWriteRefusal.HostRejected, outcome.Refusal);
    }

    [Fact]
    public void A_style_that_could_not_be_written_is_never_stamped_with_ownership()
    {
        // The discard alone is not enough. If Create carried on to stamp the
        // ownership members, it would write Name and AlternativeText onto a shape
        // that is no longer on the sheet - a write that either throws (surfacing
        // as a second, misleading refusal) or succeeds and reports Ok for a shape
        // that does not exist. So the failed style write must end the create, and
        // the stamp must never be attempted.
        var fill = new Mock<Excel.FillFormat>();
        _ = fill.SetupGet(f => f.ForeColor)
            .Throws(new InvalidOperationException("The host refused the fill."));
        var host = new HostShape(fill.Object);
        var writer = new StyleWriter(host);

        ShapeWriteOutcome outcome = writer.Create(Request(fill: "#92D050", stroke: "#548235", width: 0.75));

        Assert.False(outcome.Succeeded);
        Assert.Equal(ShapeWriteRefusal.HostRejected, outcome.Refusal);
        // The seam records the stamp; an empty list is what proves it was skipped
        // rather than merely not observed.
        Assert.Empty(writer.OwnershipStamps);
    }

    [Fact]
    public void A_host_that_refuses_the_text_write_leaves_no_untagged_text_box()
    {
        // A text box is the one kind whose content is written inside AddShape,
        // before Create has any handle on the shape. A throwing ApplyText there
        // used to escape as an unhandled COM error and leave an untagged text box
        // on the sheet that nothing could ever update or remove. Create must
        // discard it and report the same typed refusal it uses for every other
        // create failure.
        var host = new HostShape();
        host.FailTextWrite();

        var writer = new TextBoxWriter(host);

        ShapeWriteOutcome outcome = writer.Create(TextBoxRequest());

        Assert.Equal(ShapeWriteRefusal.HostRejected, outcome.Refusal);
        Assert.True(host.Deleted, "The unowned text box must be discarded.");
        Assert.Empty(writer.OwnershipStamps);
    }

    [Fact]
    public void A_solid_fill_asks_the_host_for_solid_even_when_the_shape_is_reused_from_a_hatch()
    {
        // The host keeps whichever fill type it last had, so an update that
        // removes the hatch must restore a solid fill explicitly. Without the
        // Solid() branch the colour would be written and the hatch would remain.
        var host = new HostShape();
        var writer = new StyleWriter(host, shapeAlreadyExists: true);

        writer.Update(Request(
            id: "row-1:bar",
            fill: "#00B0F0",
            stroke: "#0070C0",
            width: 0.75,
            hatch: GanttHatchPattern.ForwardDiagonal));
        Assert.Equal("Patterned:msoPatternLightDownwardDiagonal", Assert.Single(host.FillCalls));

        var second = new HostShape();
        var secondWriter = new StyleWriter(second, shapeAlreadyExists: true);
        secondWriter.Update(Request(
            id: "row-1:bar",
            fill: "#00B0F0",
            stroke: "#0070C0",
            width: 0.75,
            hatch: GanttHatchPattern.None));

        Assert.Equal("Solid", Assert.Single(second.FillCalls));
    }

    private static OfficeShapeRequest TextBoxRequest() =>
        new(
            "row-1:label",
            OfficeShapeKind.TextBox,
            new OfficeShapeGeometry(Bounds: new RectD(10, 20, 100, 30)),
            ZLayer.Label,
            Text: "Site survey");

    private static OfficeShapeRequest Request(
        string id = "row-1:bar",
        string? fill = null,
        string? stroke = null,
        double? width = null,
        GanttHatchPattern hatch = GanttHatchPattern.None,
        OfficeShapeKind kind = OfficeShapeKind.Rectangle,
        bool points = false) =>
        new(
            id,
            kind,
            points
                ? new OfficeShapeGeometry(
                    Bounds: new RectD(10, 20, 8, 8),
                    Points:
                    [
                        new PointD(14, 20),
                        new PointD(18, 24),
                        new PointD(14, 28),
                        new PointD(10, 24),
                    ])
                : kind == OfficeShapeKind.Line
                    ? new OfficeShapeGeometry(From: new PointD(10, 20), To: new PointD(10, 50))
                    : new OfficeShapeGeometry(Bounds: new RectD(10, 20, 100, 30)),
            ZLayer.ActivityBody,
            FillColour: fill is null ? null : ColourHex.Parse(fill),
            StrokeColour: stroke is null ? null : ColourHex.Parse(stroke),
            LineWidthPt: width,
            HatchPattern: hatch);
}
