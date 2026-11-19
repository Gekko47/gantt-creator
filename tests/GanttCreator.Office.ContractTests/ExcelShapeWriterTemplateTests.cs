using GanttCreator.Core;
using GanttCreator.Core.Scene;
using Microsoft.Office.Core;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// The opt-in batch template reuse: a later shape with the same appearance duplicates the
/// first one instead of paying <c>AddShape</c> and the style writes again.
/// </summary>
/// <remarks>
/// Measured 2026-11-19 on the 20-activity editable copy (Excel 16.0 x64): the style block
/// was 4.2 s of a 9.9 s lifecycle, and duplicate-plus-per-shape-writes costs 2.3 ms for a
/// rectangle and 8.6 ms for a text box against 24.9 / 34.6 ms to create them. These facts
/// pin the four decisions that make the optimisation safe: only an identical appearance
/// signature may duplicate, a line's direction is part of that signature, a refused
/// duplicate falls back to a normal create, and a create outside a batch never templates.
/// </remarks>
public sealed class ExcelShapeWriterTemplateTests
{
    [Fact]
    public void A_repeated_appearance_duplicates_the_batch_template()
    {
        TemplateWriter writer = NewWriter();
        using (writer.BeginBatch())
        {
            Assert.True(writer.Create(Rectangle("row-1:bar", y: 10)).Succeeded);
            Assert.True(writer.Create(Rectangle("row-2:bar", y: 30)).Succeeded);
            Assert.True(writer.Create(Rectangle("row-3:bar", y: 50)).Succeeded);
        }

        Assert.Equal(1, writer.AddCalls);
        Assert.Equal(2, writer.DuplicateCalls);
        _ = Assert.Single(writer.StyledRequests);
        Assert.Equal(["row-1:bar", "row-2:bar", "row-3:bar"], writer.Created.Select(shape => shape.Name));
        Assert.All(writer.Created, shape => Assert.Equal(Excel.XlPlacement.xlMove, shape.Placement));
        Assert.All(writer.Created, shape => Assert.True(shape.Locked));
    }

    [Fact]
    public void An_appearance_change_creates_instead_of_duplicating()
    {
        TemplateWriter writer = NewWriter();
        using (writer.BeginBatch())
        {
            Assert.True(writer.Create(Rectangle("row-1:bar", y: 10, fill: "#92D050")).Succeeded);
            Assert.True(writer.Create(Rectangle("row-2:bar", y: 30, fill: "#FFC7CE")).Succeeded);
        }

        Assert.Equal(2, writer.AddCalls);
        Assert.Equal(0, writer.DuplicateCalls);
    }

    [Fact]
    public void A_duplicated_text_box_writes_only_its_string()
    {
        TemplateWriter writer = NewWriter();
        using (writer.BeginBatch())
        {
            Assert.True(writer.Create(TextLabel("row-1:label", "Alpha")).Succeeded);
            Assert.True(writer.Create(TextLabel("row-2:label", "Beta")).Succeeded);
            Assert.True(writer.Create(TextLabel("row-3:label", "Gamma")).Succeeded);
        }

        // The template takes the full text path; the duplicates inherit its formatting
        // and write only the string.
        Assert.Equal(1, writer.AddCalls);
        Assert.Equal(2, writer.DuplicateCalls);
        Assert.Equal(1, writer.FullTextCalls);
        Assert.Equal(2, writer.ContentTextCalls);
        Assert.Equal(["Beta", "Gamma"], writer.ContentRequests.Select(request => request.Text));
    }

    [Fact]
    public void A_child_indentation_change_creates_instead_of_duplicating()
    {
        // The left inset is the copied cell's indentation. It is part of the appearance
        // signature, so a grouped child must never duplicate a non-indented parent.
        TemplateWriter writer = NewWriter();
        using (writer.BeginBatch())
        {
            Assert.True(writer.Create(TextLabel("row-1:label", "Parent", inset: 0)).Succeeded);
            Assert.True(writer.Create(TextLabel("row-2:label", "Child", inset: 9)).Succeeded);
        }

        Assert.Equal(2, writer.AddCalls);
        Assert.Equal(0, writer.DuplicateCalls);
    }

    [Fact]
    public void An_alignment_change_creates_instead_of_duplicating()
    {
        TemplateWriter writer = NewWriter();
        using (writer.BeginBatch())
        {
            Assert.True(writer.Create(TextLabel("row-1:label", "Left", alignment: GanttTextAlignment.Left)).Succeeded);
            Assert.True(writer.Create(TextLabel("row-2:label", "Centre", alignment: GanttTextAlignment.Centre)).Succeeded);
        }

        Assert.Equal(2, writer.AddCalls);
        Assert.Equal(0, writer.DuplicateCalls);
    }

    [Fact]
    public void A_duplicate_inherits_the_template_placement_and_lock()
    {
        TemplateWriter writer = NewWriter();
        using (writer.BeginBatch())
        {
            Assert.True(writer.Create(Rectangle("row-1:bar", y: 10)).Succeeded);
            Assert.True(writer.Create(Rectangle("row-2:bar", y: 30)).Succeeded);
        }

        Assert.Equal(1, writer.DuplicateCalls);
        Assert.Equal(Excel.XlPlacement.xlMove, writer.Created[1].Placement);
        Assert.True(writer.Created[1].Locked);
    }

    [Fact]
    public void A_duplicate_that_lost_its_placement_is_healed()
    {
        // Defensive coverage for a host whose Duplicate does not copy Placement: the
        // duplicate must still end on xlMove, or a row resize could stretch it.
        TemplateWriter writer = NewWriter();
        writer.LosePlacementOnDuplicate = true;
        using (writer.BeginBatch())
        {
            Assert.True(writer.Create(Rectangle("row-1:bar", y: 10)).Succeeded);
            Assert.True(writer.Create(Rectangle("row-2:bar", y: 30)).Succeeded);
        }

        Assert.Equal(Excel.XlPlacement.xlMove, writer.Created[1].Placement);
    }

    [Fact]
    public void A_duplicate_that_lost_its_lock_is_healed()
    {
        // Defensive coverage for a host whose Duplicate does not copy Locked: an
        // unlocked generated shape must not reach the worksheet.
        TemplateWriter writer = NewWriter();
        writer.LoseLockOnDuplicate = true;
        using (writer.BeginBatch())
        {
            Assert.True(writer.Create(Rectangle("row-1:bar", y: 10)).Succeeded);
            Assert.True(writer.Create(Rectangle("row-2:bar", y: 30)).Succeeded);
        }

        Assert.True(writer.Created[1].Locked);
    }

    [Fact]
    public void A_host_owned_extent_change_creates_instead_of_duplicating()
    {
        // HostOwnsExtent drives Placement; it is part of the key, so a host-owned band
        // must never duplicate a scene-owned bar.
        TemplateWriter writer = NewWriter();
        using (writer.BeginBatch())
        {
            Assert.True(writer.Create(Rectangle("row-1:bar", y: 10, hostOwnsExtent: false)).Succeeded);
            Assert.True(writer.Create(Rectangle("row-2:bar", y: 30, hostOwnsExtent: true)).Succeeded);
        }

        Assert.Equal(2, writer.AddCalls);
        Assert.Equal(0, writer.DuplicateCalls);
    }

    [Fact]
    public void A_line_with_the_opposite_direction_is_not_duplicated()
    {
        TemplateWriter writer = NewWriter();
        using (writer.BeginBatch())
        {
            Assert.True(writer.Create(Line("chart:grid:1", fromX: 10, toX: 100)).Succeeded);
            Assert.True(writer.Create(Line("chart:grid:2", fromX: 100, toX: 10)).Succeeded);
        }

        Assert.Equal(2, writer.AddCalls);
        Assert.Equal(0, writer.DuplicateCalls);
    }

    [Fact]
    public void A_refused_duplicate_falls_back_to_a_normal_create()
    {
        TemplateWriter writer = NewWriter(failDuplication: true);
        using (writer.BeginBatch())
        {
            Assert.True(writer.Create(Rectangle("row-1:bar", y: 10)).Succeeded);
            Assert.True(writer.Create(Rectangle("row-2:bar", y: 30)).Succeeded);
        }

        Assert.Equal(2, writer.AddCalls);
        Assert.Equal(1, writer.DuplicateCalls);
    }

    [Fact]
    public void A_duplicate_whose_identity_stamp_fails_is_discarded()
    {
        // The duplicate exists on the sheet by the time the identity is written. A host
        // refusal there must delete it and report HostRejected, never leave an untagged
        // temporary shape for the cleanup registry to miss.
        TemplateWriter writer = NewWriter();
        using (writer.BeginBatch())
        {
            Assert.True(writer.Create(Rectangle("row-1:bar", y: 10)).Succeeded);
            writer.FailDuplicateOwnership = true;
            ShapeWriteOutcome refused = writer.Create(Rectangle("row-2:bar", y: 30));
            Assert.False(refused.Succeeded);
            Assert.Equal(ShapeWriteRefusal.HostRejected, refused.Refusal);
        }

        Assert.Equal(1, writer.DuplicateCalls);
        Assert.True(writer.DeleteCalls > 0);
    }

    [Fact]
    public void A_create_outside_a_batch_never_reuses_a_template()
    {
        TemplateWriter writer = NewWriter();

        Assert.True(writer.Create(Rectangle("row-1:bar", y: 10)).Succeeded);
        Assert.True(writer.Create(Rectangle("row-2:bar", y: 30)).Succeeded);

        Assert.Equal(2, writer.AddCalls);
        Assert.Equal(0, writer.DuplicateCalls);
    }

    private static TemplateWriter NewWriter(bool failDuplication = false) =>
        new(ActiveApplication().Object, ClearGuard().Object, failDuplication);

    private static Mock<IWorksheetProtectionGuard> ClearGuard()
    {
        Mock<IWorksheetProtectionGuard> guard = new();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        _ = guard.Setup(g => g.QueryTarget(It.IsAny<object>())).Returns(ProtectionGuardOutcome.NotProtected);
        return guard;
    }

    private static Mock<Excel.Application> ActiveApplication()
    {
        Mock<Excel.Workbook> workbook = new();
        _ = workbook.Setup(w => w.Sheets).Returns(new Mock<Excel.Sheets>().Object);
        Mock<Excel.Application> application = new();
        _ = application.Setup(a => a.ActiveWorkbook).Returns(workbook.Object);
        return application;
    }

    private static OfficeShapeRequest Rectangle(
        string id,
        double y,
        string fill = "#92D050",
        bool hostOwnsExtent = false) =>
        new(
            id,
            OfficeShapeKind.Rectangle,
            new OfficeShapeGeometry(Bounds: new RectD(10, y, 100, 18)),
            ZLayer.ActivityBody,
            FillColour: ColourHex.Parse(fill),
            StrokeColour: ColourHex.Parse("#375623"),
            LineWidthPt: 0.5,
            HostOwnsExtent: hostOwnsExtent);

    private static OfficeShapeRequest TextLabel(
        string id,
        string text,
        double y = 10,
        double inset = 0,
        GanttTextAlignment alignment = GanttTextAlignment.Left) =>
        new(
            id,
            OfficeShapeKind.TextBox,
            new OfficeShapeGeometry(Bounds: new RectD(10, y, 100, 16)),
            ZLayer.Label,
            FontFamily: "Aptos",
            FontSizePt: 8,
            Bold: false,
            Text: text,
            Alignment: alignment,
            TextColour: ColourHex.Parse("#000000"),
            LeftInsetPt: inset);

    private static OfficeShapeRequest Line(string id, double fromX, double toX) =>
        new(
            id,
            OfficeShapeKind.Line,
            new OfficeShapeGeometry(
                From: new PointD(fromX, 40),
                To: new PointD(toX, 40)),
            ZLayer.Grid,
            StrokeColour: ColourHex.Parse("#808080"),
            LineWidthPt: 0.5,
            StrokeDash: GanttStrokeDash.Solid);

    /// <summary>A writer whose host is a list of shape mocks and which records the create path.</summary>
    private sealed class TemplateWriter(
        object? application,
        IWorksheetProtectionGuard guard,
        bool failDuplication)
        : ExcelShapeWriter(application, guard, reuseTemplates: true)
    {
        private readonly List<Excel.Shape> _shapes = [];

        public List<Excel.Shape> Created { get; } = [];

        public List<OfficeShapeRequest> StyledRequests { get; } = [];

        public List<OfficeShapeRequest> ContentRequests { get; } = [];

        public int AddCalls { get; private set; }

        public int FullTextCalls { get; private set; }

        public int ContentTextCalls { get; private set; }

        public int DuplicateCalls { get; private set; }

        /// <summary>Models a host whose Duplicate does not copy Placement.</summary>
        public bool LosePlacementOnDuplicate { get; set; }

        /// <summary>Models a host whose Duplicate does not copy Locked.</summary>
        public bool LoseLockOnDuplicate { get; set; }

        /// <summary>Models a host that refuses the duplicate's identity stamp.</summary>
        public bool FailDuplicateOwnership { get; set; }

        /// <summary>How many modelled shapes were deleted (the discard path).</summary>
        public int DeleteCalls { get; private set; }

        internal override Excel._Worksheet? FindGanttWorksheet(Excel.Sheets sheets) =>
            new Mock<Excel._Worksheet>().Object;

        internal override Excel.Shapes? GetShapes(Excel._Worksheet sheet) =>
            new Mock<Excel.Shapes>().Object;

        internal override Excel.Shape? FindShapeByName(Excel.Shapes shapes, string name) =>
            _shapes.FirstOrDefault(shape => string.Equals(shape.Name, name, StringComparison.Ordinal));

        internal override Excel.Shape? AddShape(Excel.Shapes shapes, OfficeShapeRequest request)
        {
            AddCalls++;
            Excel.Shape created = NewHostShape();
            _shapes.Add(created);
            Created.Add(created);
            if (request.Kind == OfficeShapeKind.TextBox)
            {
                ApplyText(created, request);
            }

            return created;
        }

        internal override Excel.Shape? DuplicateShape(Excel.Shape template)
        {
            DuplicateCalls++;
            if (failDuplication)
            {
                return null;
            }

            Excel.XlPlacement placement = LosePlacementOnDuplicate
                ? Excel.XlPlacement.xlMoveAndSize
                : template.Placement;
            bool locked = !LoseLockOnDuplicate && template.Locked;
            Excel.Shape duplicate = NewHostShape(placement, locked);
            _shapes.Add(duplicate);
            Created.Add(duplicate);
            return duplicate;
        }

        internal override void ApplyStyle(Excel.Shape shape, OfficeShapeRequest request) =>
            StyledRequests.Add(request);

        internal override void ApplyText(Excel.Shape shape, OfficeShapeRequest request) =>
            FullTextCalls++;

        internal override void ApplyTextContent(Excel.Shape shape, OfficeShapeRequest request)
        {
            ContentTextCalls++;
            ContentRequests.Add(request);
        }

        internal override void ApplyDuplicateOwnershipStamp(Excel.Shape shape, string primitiveId)
        {
            if (FailDuplicateOwnership)
            {
                throw new InvalidOperationException("The host refused the duplicate identity.");
            }

            base.ApplyDuplicateOwnershipStamp(shape, primitiveId);
        }

        private Excel.Shape NewHostShape(
            Excel.XlPlacement placement = Excel.XlPlacement.xlMoveAndSize,
            bool locked = false)
        {
            Mock<Excel.Shape> mock = new();
            string name = string.Empty;
            string alternativeText = string.Empty;
            _ = mock.SetupProperty(shape => shape.Placement, placement);
            _ = mock.SetupProperty(shape => shape.Locked, locked);
            _ = mock.SetupSet(shape => shape.Name = It.IsAny<string>()).Callback<string>(value => name = value);
            _ = mock.SetupGet(shape => shape.Name).Returns(() => name);
            _ = mock.SetupSet(shape => shape.AlternativeText = It.IsAny<string>())
                .Callback<string>(value => alternativeText = value);
            _ = mock.SetupGet(shape => shape.AlternativeText).Returns(() => alternativeText);
            _ = mock.Setup(shape => shape.Delete()).Callback(() => DeleteCalls++);
            _ = mock.Setup(shape => shape.ZOrder(It.IsAny<MsoZOrderCmd>()));
            return mock.Object;
        }
    }
}
