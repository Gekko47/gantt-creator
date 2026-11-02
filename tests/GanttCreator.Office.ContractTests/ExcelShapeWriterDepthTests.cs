using GanttCreator.Core;
using GanttCreator.Core.Scene;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Moq;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Depth tests for <see cref="ExcelShapeWriter"/>: the refusal arms of every
/// entry point, the worksheet/shape traversal the adapter performs for itself,
/// and the create-path bodies that a live host would execute.
/// </summary>
/// <remarks>
/// <para>
/// Deliberately a different cut from <see cref="ExcelShapeWriterTests"/>. That
/// file substitutes <c>FindGanttWorksheet</c>, <c>HasGanttTable</c>,
/// <c>FindShapeByName</c> and <c>AddShape</c> so its assertions are about the
/// rules the adapter enforces. This file substitutes only the <em>lowest</em>
/// seams — the COM indexers and the shape-collection reads — so those four
/// members run for real against a modelled collection. They contain loops and
/// name comparisons, which are product decisions, not host plumbing.
/// </para>
/// <para>
/// The two members that stay out of reach are recorded rather than chased:
/// <c>ApplyText</c>'s body walks a <c>TextFrame2</c> chain whose members Castle
/// DynamicProxy cannot proxy, and the <c>GetSheetAt</c>/<c>GetShapeAt</c>-style
/// one-line indexer seams need a live host. Both are what
/// <c>verify-office.ps1</c> exercises.
/// </para>
/// </remarks>
public sealed class ExcelShapeWriterDepthTests
{
    /// <summary>
    /// A modelled host shape: writable geometry backing fields, an optional
    /// modelled fill and line, and the failure switches the discard paths need.
    /// </summary>
    private sealed class FakeShape
    {
        public FakeShape(
            string name = "row-1:bar",
            string? alternativeText = null,
            bool withFillAndLine = false)
        {
            Name = name;
            AlternativeText = alternativeText ?? string.Empty;

            Mock<Excel.Shape> shape = new();
            _ = shape.SetupGet(s => s.Name).Returns(() => Name);
            _ = shape.SetupSet(s => s.Name = It.IsAny<string>()).Callback<string>(value =>
            {
                if (NameWriteThrows)
                {
                    throw new InvalidOperationException("The host refused the name.");
                }

                Name = value;
            });
            _ = shape.SetupGet(s => s.AlternativeText).Returns(() => AlternativeText);
            _ = shape.SetupSet(s => s.AlternativeText = It.IsAny<string>()).Callback<string>(value =>
            {
                if (TagWriteThrows)
                {
                    throw new InvalidOperationException("The host refused the tag.");
                }

                AlternativeText = value;
            });
            _ = shape.SetupGet(s => s.Left).Returns(() => Left);
            _ = shape.SetupSet(s => s.Left = It.IsAny<float>()).Callback<float>(value => Left = value);
            _ = shape.SetupGet(s => s.Top).Returns(() => Top);
            _ = shape.SetupSet(s => s.Top = It.IsAny<float>()).Callback<float>(value => Top = value);
            _ = shape.SetupGet(s => s.Width).Returns(() => Width);
            _ = shape.SetupSet(s => s.Width = It.IsAny<float>()).Callback<float>(value => Width = value);
            _ = shape.SetupGet(s => s.Height).Returns(() => Height);
            _ = shape.SetupSet(s => s.Height = It.IsAny<float>()).Callback<float>(value => Height = value);
            _ = shape.Setup(s => s.Delete()).Callback(() =>
            {
                if (DeleteThrows)
                {
                    throw new InvalidOperationException("The host refused the delete.");
                }

                Deleted = true;
            });

            if (withFillAndLine)
            {
                Mock<Excel.ColorFormat> fillColour = new();
                _ = fillColour.SetupGet(c => c.RGB).Returns(() => FillRgb);
                _ = fillColour.SetupSet(c => c.RGB = It.IsAny<int>()).Callback<int>(value => FillRgb = value);
                Mock<Excel.FillFormat> fill = new();
                _ = fill.SetupGet(f => f.ForeColor).Returns(fillColour.Object);
                _ = fill.SetupGet(f => f.BackColor).Returns(fillColour.Object);
                _ = fill.SetupGet(f => f.Visible).Returns(() => FillVisible);
                _ = fill.SetupSet(f => f.Visible = It.IsAny<MsoTriState>())
                    .Callback<MsoTriState>(value => FillVisible = value);
                _ = fill.SetupGet(f => f.Transparency).Returns(() => FillTransparency);
                _ = fill.SetupSet(f => f.Transparency = It.IsAny<float>())
                    .Callback<float>(value => FillTransparency = value);
                _ = fill.Setup(f => f.Solid()).Callback(() => FillCalls.Add("Solid"));
                _ = fill.Setup(f => f.Patterned(It.IsAny<MsoPatternType>()))
                    .Callback<MsoPatternType>(pattern => FillCalls.Add($"Patterned:{pattern}"));
                Fill = fill.Object;
                _ = shape.SetupGet(s => s.Fill).Returns(Fill);

                Mock<Excel.ColorFormat> lineColour = new();
                _ = lineColour.SetupGet(c => c.RGB).Returns(() => LineRgb);
                _ = lineColour.SetupSet(c => c.RGB = It.IsAny<int>()).Callback<int>(value => LineRgb = value);
                Mock<Excel.LineFormat> line = new();
                _ = line.SetupGet(l => l.ForeColor).Returns(lineColour.Object);
                _ = line.SetupGet(l => l.Visible).Returns(() => LineVisible);
                _ = line.SetupSet(l => l.Visible = It.IsAny<MsoTriState>())
                    .Callback<MsoTriState>(value => LineVisible = value);
                _ = line.SetupGet(l => l.Weight).Returns(() => LineWeight);
                _ = line.SetupSet(l => l.Weight = It.IsAny<float>())
                    .Callback<float>(value => LineWeight = value);
                _ = line.SetupGet(l => l.DashStyle).Returns(() => LineDash);
                _ = line.SetupSet(l => l.DashStyle = It.IsAny<MsoLineDashStyle>())
                    .Callback<MsoLineDashStyle>(value => LineDash = value);
                _ = line.SetupGet(l => l.Transparency).Returns(() => LineTransparency);
                _ = line.SetupSet(l => l.Transparency = It.IsAny<float>())
                    .Callback<float>(value => LineTransparency = value);
                Line = line.Object;
                _ = shape.SetupGet(s => s.Line).Returns(Line);
            }

            Mock = shape;
        }

        public Mock<Excel.Shape> Mock { get; }

        public Excel.Shape Shape => Mock.Object;

        public string Name { get; private set; }

        public string AlternativeText { get; private set; }

        public float Left { get; private set; }

        public float Top { get; private set; }

        public float Width { get; private set; }

        public float Height { get; private set; }

        public bool Deleted { get; private set; }

        public bool DeleteThrows { get; set; }

        public bool NameWriteThrows { get; set; }

        public bool TagWriteThrows { get; set; }

        public Excel.FillFormat? Fill { get; }

        public Excel.LineFormat? Line { get; }

        public List<string> FillCalls { get; } = [];

        public int FillRgb { get; private set; }

        public MsoTriState FillVisible { get; private set; }

        public float FillTransparency { get; private set; }

        public int LineRgb { get; private set; }

        public MsoTriState LineVisible { get; private set; }

        public float LineWeight { get; private set; }

        public MsoLineDashStyle LineDash { get; private set; }

        public float LineTransparency { get; private set; }
    }

    /// <summary>
    /// The writer with only the lowest COM seams substituted, so
    /// <c>FindGanttWorksheet</c>, <c>HasGanttTable</c>, <c>FindShapeByName</c>,
    /// <c>AddShape</c> and <c>AddDiamond</c> all run.
    /// </summary>
    private sealed class DepthWriter(
        object? application,
        IWorksheetProtectionGuard guard,
        Excel.Shapes? shapes = null,
        params FakeShape[] existing) : ExcelShapeWriter(application, guard)
    {
        private readonly List<Excel._Worksheet?> _sheets = [];
        private readonly List<Excel.ListObject?> _tables = [];
        private readonly Excel.Shapes? _shapes = shapes;

        public List<FakeShape> Existing { get; } = [.. existing];

        public List<OfficeShapeRequest> TextWrites { get; } = [];

        public List<OfficeShapeRequest> StyleWrites { get; } = [];

        public Exception? TextWriteFailure { get; set; }

        public bool CallBaseStyle { get; set; }

        public DepthWriter WithSheets(params Excel._Worksheet?[] sheets)
        {
            _sheets.AddRange(sheets);
            return this;
        }

        public DepthWriter WithTables(params Excel.ListObject?[] tables)
        {
            _tables.AddRange(tables);
            return this;
        }

        internal override int GetSheetCount(Excel.Sheets sheets) => _sheets.Count;

        internal override Excel._Worksheet? GetSheetAt(Excel.Sheets sheets, int index) =>
            index >= 1 && index <= _sheets.Count ? _sheets[index - 1] : null;

        internal override int GetListObjectCount(Excel.ListObjects objects) => _tables.Count;

        internal override Excel.ListObject? GetListObjectAt(Excel.ListObjects objects, int index) =>
            index >= 1 && index <= _tables.Count ? _tables[index - 1] : null;

        internal override Excel.Shapes? GetShapes(Excel._Worksheet sheet) => _shapes;

        internal override int GetShapeCount(Excel.Shapes shapes) => Existing.Count;

        internal override Excel.Shape? GetShapeAt(Excel.Shapes shapes, int index) =>
            index >= 1 && index <= Existing.Count ? Existing[index - 1].Shape : null;

        internal override void ApplyText(Excel.Shape shape, OfficeShapeRequest request)
        {
            if (TextWriteFailure is not null)
            {
                throw TextWriteFailure;
            }

            TextWrites.Add(request);
        }

        internal override void ApplyStyle(Excel.Shape shape, OfficeShapeRequest request)
        {
            StyleWrites.Add(request);
            if (CallBaseStyle)
            {
                base.ApplyStyle(shape, request);
            }
        }
    }

    private static Mock<IWorksheetProtectionGuard> Guard(
        ProtectionGuardOutcome query = ProtectionGuardOutcome.NotProtected,
        ProtectionGuardOutcome target = ProtectionGuardOutcome.NotProtected)
    {
        Mock<IWorksheetProtectionGuard> guard = new();
        _ = guard.Setup(g => g.Query()).Returns(query);
        _ = guard.Setup(g => g.QueryTarget(It.IsAny<object>())).Returns(target);
        return guard;
    }

    private static Excel.Sheets Sheets() => new Mock<Excel.Sheets>().Object;

    private static Excel.ListObjects TableCollection() => new Mock<Excel.ListObjects>().Object;

    private static Excel.Shapes ShapesCollection() => new Mock<Excel.Shapes>().Object;

    private static Excel.ListObject Table(string name)
    {
        Mock<Excel.ListObject> table = new();
        _ = table.SetupGet(t => t.Name).Returns(name);
        return table.Object;
    }

    private static Excel._Worksheet Sheet(Excel.ListObjects? tables)
    {
        Mock<Excel._Worksheet> sheet = new();
        if (tables is not null)
        {
            _ = sheet.SetupGet(s => s.ListObjects).Returns(tables);
        }

        return sheet.Object;
    }

    /// <summary>An application whose active workbook reports the given sheets.</summary>
    private static Excel.Application Application(Excel.Sheets? sheets)
    {
        Mock<Excel.Workbook> workbook = new();
        if (sheets is not null)
        {
            _ = workbook.SetupGet(w => w.Sheets).Returns(sheets);
        }

        Mock<Excel.Application> application = new();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        return application.Object;
    }

    private static OfficeShapeRequest Rectangle(
        string id = "row-1:bar",
        double x = 10,
        double y = 20,
        double width = 100,
        double height = 30,
        double? lineWidthPt = null,
        double? fontSizePt = null) =>
        new(
            id,
            OfficeShapeKind.Rectangle,
            new OfficeShapeGeometry(Bounds: new RectD(x, y, width, height)),
            ZLayer.ActivityBody,
            LineWidthPt: lineWidthPt,
            FontSizePt: fontSizePt);

    private static OfficeShapeRequest LineRequest(string id = "row-1:critical") =>
        new(
            id,
            OfficeShapeKind.Line,
            new OfficeShapeGeometry(From: new PointD(100, 50), To: new PointD(20, 30)),
            ZLayer.CriticalOverlay);

    private static OfficeShapeRequest DiamondRequest(string id = "row-1:marker") =>
        new(
            id,
            OfficeShapeKind.Diamond,
            new OfficeShapeGeometry(
                Bounds: new RectD(40, 40, 20, 20),
                Points: [new PointD(50, 40), new PointD(60, 50), new PointD(50, 60), new PointD(40, 50)]),
            ZLayer.Milestone);

    /// <summary>
    /// A writer wired to a workbook that resolves the Gantt sheet and a shapes
    /// collection, so the traversal succeeds and the caller reaches its own logic.
    /// </summary>
    private static DepthWriter Writer(
        Excel.Shapes? shapes = null,
        ProtectionGuardOutcome query = ProtectionGuardOutcome.NotProtected,
        ProtectionGuardOutcome target = ProtectionGuardOutcome.NotProtected,
        params FakeShape[] existing)
    {
        DepthWriter writer = new(Application(Sheets()), Guard(query, target).Object, shapes, existing);
        return writer
            .WithSheets(Sheet(TableCollection()))
            .WithTables(Table(GanttTableSchema.TableName));
    }

    [Fact]
    public void Every_entry_point_reports_target_protected_when_the_active_sheet_is_protected()
    {
        DepthWriter writer = new(
            Application(Sheets()),
            Guard(ProtectionGuardOutcome.SheetProtected).Object,
            ShapesCollection());

        Assert.Equal(ShapeWriteRefusal.TargetProtected, writer.Create(Rectangle()).Refusal);
        Assert.Equal(ShapeWriteRefusal.TargetProtected, writer.Update(Rectangle()).Refusal);
        Assert.Equal(ShapeWriteRefusal.TargetProtected, writer.Delete("row-1:bar").Refusal);
        Assert.Equal(ShapeWriteRefusal.TargetProtected, writer.ApplyZOrder(["row-1:bar"]).Refusal);
    }

    [Fact]
    public void Every_entry_point_reports_no_active_workbook_when_the_guard_cannot_see_one()
    {
        DepthWriter writer = new(
            Application(Sheets()),
            Guard(ProtectionGuardOutcome.NoActiveWorkbook).Object,
            ShapesCollection());

        Assert.Equal(ShapeWriteRefusal.NoActiveWorkbook, writer.Create(Rectangle()).Refusal);
        Assert.Equal(ShapeWriteRefusal.NoActiveWorkbook, writer.Update(Rectangle()).Refusal);
        Assert.Equal(ShapeWriteRefusal.NoActiveWorkbook, writer.Delete("row-1:bar").Refusal);
        Assert.Equal(ShapeWriteRefusal.NoActiveWorkbook, writer.ApplyZOrder(["row-1:bar"]).Refusal);
    }

    [Fact]
    public void An_update_of_an_identifier_that_does_not_exist_reports_not_found()
    {
        DepthWriter writer = Writer(ShapesCollection());

        Assert.Equal(ShapeWriteRefusal.NotFound, writer.Update(Rectangle()).Refusal);
        Assert.Empty(writer.StyleWrites);
    }

    [Fact]
    public void A_delete_of_an_identifier_that_does_not_exist_reports_not_found()
    {
        DepthWriter writer = Writer(ShapesCollection());

        Assert.Equal(ShapeWriteRefusal.NotFound, writer.Delete("row-1:bar").Refusal);
    }

    [Fact]
    public void A_delete_of_an_owned_shape_deletes_exactly_that_shape()
    {
        FakeShape owned = new("row-1:bar", ShapeOwnershipTag.ForPrimitiveId("row-1:bar"));
        FakeShape other = new("row-2:bar", ShapeOwnershipTag.ForPrimitiveId("row-2:bar"));
        DepthWriter writer = Writer(ShapesCollection(), existing: [owned, other]);

        ShapeWriteOutcome outcome = writer.Delete("row-1:bar");

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.True(owned.Deleted);
        Assert.False(other.Deleted);
    }

    [Fact]
    public void The_z_order_pass_with_nothing_to_order_is_a_no_op_success()
    {
        DepthWriter writer = Writer(ShapesCollection());

        Assert.True(writer.ApplyZOrder([]).Succeeded);
    }

    [Fact]
    public void The_z_order_pass_reports_worksheet_missing_when_the_shapes_collection_is_absent()
    {
        DepthWriter writer = Writer(shapes: null);

        Assert.Equal(ShapeWriteRefusal.WorksheetMissing, writer.ApplyZOrder(["row-1:bar"]).Refusal);
    }

    [Fact]
    public void An_absent_active_workbook_reports_no_active_workbook()
    {
        DepthWriter writer = new(new Mock<Excel.Application>().Object, Guard().Object, ShapesCollection());

        Assert.Equal(ShapeWriteRefusal.NoActiveWorkbook, writer.Create(Rectangle()).Refusal);
    }

    [Fact]
    public void An_absent_sheet_collection_reports_worksheet_missing()
    {
        DepthWriter writer = new(Application(sheets: null), Guard().Object, ShapesCollection());

        Assert.Equal(ShapeWriteRefusal.WorksheetMissing, writer.Create(Rectangle()).Refusal);
    }

    [Fact]
    public void A_workbook_without_a_gantt_sheet_reports_worksheet_missing()
    {
        // Two sheets, neither carrying the table: the traversal must walk the
        // whole collection and then refuse, not throw on the last one.
        DepthWriter writer = new(Application(Sheets()), Guard().Object, ShapesCollection());
        _ = writer.WithSheets(Sheet(TableCollection()), Sheet(tables: null));

        Assert.Equal(ShapeWriteRefusal.WorksheetMissing, writer.Create(Rectangle()).Refusal);
    }

    [Fact]
    public void The_gantt_sheet_is_found_behind_a_sheet_that_does_not_carry_the_table()
    {
        // The first sheet has tables but none named for the schema; the second
        // carries it. FindGanttWorksheet must not stop at the first table it sees,
        // and HasGanttTable must compare the table name rather than accept any.
        Excel.ListObject decoy = Table("tblSomethingElse");
        DepthWriter writer = new(Application(Sheets()), Guard().Object, ShapesCollection());
        _ = writer
            .WithSheets(Sheet(TableCollection()), Sheet(TableCollection()))
            .WithTables(decoy, Table(GanttTableSchema.TableName));

        // An empty z-order list succeeds only after ResolveShapes has found the
        // worksheet and its shapes, so success here proves the traversal reached
        // the second sheet rather than refusing on the first.
        Assert.True(writer.ApplyZOrder([]).Succeeded);
    }

    [Fact]
    public void A_sheet_whose_table_collection_is_absent_does_not_carry_the_gantt_table()
    {
        DepthWriter writer = new(Application(Sheets()), Guard().Object, ShapesCollection());
        _ = writer.WithSheets(Sheet(tables: null));

        Assert.Null(Sheet(tables: null).ListObjects);
        Assert.Equal(ShapeWriteRefusal.WorksheetMissing, writer.Create(Rectangle()).Refusal);
    }

    [Fact]
    public void An_identifier_present_as_a_name_but_owned_by_nothing_is_not_found()
    {
        // A user shape can share the name. Ownership is the tag, not the name.
        FakeShape userShape = new("row-1:bar", "my own note");
        DepthWriter writer = Writer(ShapesCollection(), existing: [userShape]);

        Assert.Equal(ShapeWriteRefusal.NotFound, writer.Update(Rectangle()).Refusal);
        Assert.Equal(ShapeWriteRefusal.NotFound, writer.Delete("row-1:bar").Refusal);
        Assert.False(userShape.Deleted);
    }

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(-1d)]
    public void A_negative_or_non_finite_line_width_is_refused_before_the_host_is_touched(double width)
    {
        DepthWriter writer = Writer(ShapesCollection());

        Assert.Equal(ShapeWriteRefusal.InvalidGeometry, writer.Create(Rectangle(lineWidthPt: width)).Refusal);
    }

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(0d)]
    public void A_zero_or_non_finite_font_size_is_refused_before_the_host_is_touched(double size)
    {
        DepthWriter writer = Writer(ShapesCollection());

        Assert.Equal(ShapeWriteRefusal.InvalidGeometry, writer.Create(Rectangle(fontSizePt: size)).Refusal);
    }

    [Fact]
    public void An_undefined_shape_kind_is_refused_as_invalid_geometry()
    {
        DepthWriter writer = Writer(ShapesCollection());
        OfficeShapeRequest request = new(
            "row-1:bar",
            (OfficeShapeKind)99,
            new OfficeShapeGeometry(Bounds: new RectD(0, 0, 10, 10)),
            ZLayer.ActivityBody);

        Assert.Equal(ShapeWriteRefusal.InvalidGeometry, writer.Create(request).Refusal);
    }

    [Fact]
    public void A_line_update_normalises_a_right_to_left_request_into_positive_box_members()
    {
        // The host's Shape surface exposes only a bounding box, so a line drawn
        // right-to-left has to arrive as left/top plus positive width/height.
        FakeShape owned = new("row-1:critical", ShapeOwnershipTag.ForPrimitiveId("row-1:critical"));
        DepthWriter writer = Writer(ShapesCollection(), existing: [owned]);

        ShapeWriteOutcome outcome = writer.Update(LineRequest());

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal(20f, owned.Left);
        Assert.Equal(30f, owned.Top);
        Assert.Equal(80f, owned.Width);
        Assert.Equal(20f, owned.Height);
    }

    [Fact]
    public void A_diamond_update_moves_the_shape_to_its_bounding_box()
    {
        FakeShape owned = new("row-1:marker", ShapeOwnershipTag.ForPrimitiveId("row-1:marker"));
        DepthWriter writer = Writer(ShapesCollection(), existing: [owned]);

        ShapeWriteOutcome outcome = writer.Update(DiamondRequest());

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal(40f, owned.Left);
        Assert.Equal(40f, owned.Top);
        Assert.Equal(20f, owned.Width);
        Assert.Equal(20f, owned.Height);
    }

    [Fact]
    public void An_update_whose_content_write_is_a_contract_fault_propagates_instead_of_refusing()
    {
        // ArgumentOutOfRangeException is MapAlignment's drift signal, not a host
        // failure: swallowing it into a refusal would hide the drift.
        FakeShape owned = new("row-1:label", ShapeOwnershipTag.ForPrimitiveId("row-1:label"));
        DepthWriter writer = Writer(ShapesCollection(), existing: [owned]);
        writer.TextWriteFailure = new ArgumentOutOfRangeException("alignment", "no host mapping");

        OfficeShapeRequest request = new(
            "row-1:label",
            OfficeShapeKind.TextBox,
            new OfficeShapeGeometry(Bounds: new RectD(10, 10, 40, 12)),
            ZLayer.Label,
            Text: "label");

        _ = Assert.Throws<ArgumentOutOfRangeException>(() => writer.Update(request));
    }

    [Fact]
    public void A_rectangle_is_created_through_the_real_add_shape_body()
    {
        Mock<Excel.Shapes> shapes = new();
        FakeShape created = new();
        MsoAutoShapeType kind = MsoAutoShapeType.msoShapeMixed;
        float left = 0f, top = 0f, width = 0f, height = 0f;
        _ = shapes
            .Setup(s => s.AddShape(
                It.IsAny<MsoAutoShapeType>(),
                It.IsAny<float>(),
                It.IsAny<float>(),
                It.IsAny<float>(),
                It.IsAny<float>()))
            .Callback<MsoAutoShapeType, float, float, float, float>((type, l, t, w, h) =>
            {
                kind = type;
                left = l;
                top = t;
                width = w;
                height = h;
            })
            .Returns(created.Shape);

        DepthWriter writer = Writer(shapes.Object);

        ShapeWriteOutcome outcome = writer.Create(Rectangle(x: 10.5, y: 20, width: 100, height: 30));

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal(MsoAutoShapeType.msoShapeRectangle, kind);
        Assert.Equal(10.5f, left);
        Assert.Equal(20f, top);
        Assert.Equal(100f, width);
        Assert.Equal(30f, height);
        Assert.Equal("row-1:bar", created.Name);
    }

    [Fact]
    public void A_line_is_created_through_the_real_add_line_body()
    {
        Mock<Excel.Shapes> shapes = new();
        FakeShape created = new("row-1:critical");
        float beginX = 0f, beginY = 0f, endX = 0f, endY = 0f;
        _ = shapes
            .Setup(s => s.AddLine(
                It.IsAny<float>(), It.IsAny<float>(), It.IsAny<float>(), It.IsAny<float>()))
            .Callback<float, float, float, float>((bx, by, ex, ey) =>
            {
                beginX = bx;
                beginY = by;
                endX = ex;
                endY = ey;
            })
            .Returns(created.Shape);

        DepthWriter writer = Writer(shapes.Object);

        Assert.True(writer.Create(LineRequest()).Succeeded);
        Assert.Equal(100f, beginX);
        Assert.Equal(50f, beginY);
        Assert.Equal(20f, endX);
        Assert.Equal(30f, endY);
    }

    [Fact]
    public void A_textbox_is_created_as_a_textbox_and_its_content_applied()
    {
        Mock<Excel.Shapes> shapes = new();
        FakeShape created = new("row-1:label");
        MsoTextOrientation orientation = MsoTextOrientation.msoTextOrientationMixed;
        _ = shapes
            .Setup(s => s.AddTextbox(
                It.IsAny<MsoTextOrientation>(),
                It.IsAny<float>(),
                It.IsAny<float>(),
                It.IsAny<float>(),
                It.IsAny<float>()))
            .Callback<MsoTextOrientation, float, float, float, float>(
                (value, _, _, _, _) => orientation = value)
            .Returns(created.Shape);

        DepthWriter writer = Writer(shapes.Object);
        OfficeShapeRequest request = new(
            "row-1:label",
            OfficeShapeKind.TextBox,
            new OfficeShapeGeometry(Bounds: new RectD(40, 60, 80, 14)),
            ZLayer.Label,
            Text: "label");

        Assert.True(writer.Create(request).Succeeded);
        Assert.Equal(MsoTextOrientation.msoTextOrientationHorizontal, orientation);
        OfficeShapeRequest applied = Assert.Single(writer.TextWrites);
        Assert.Equal("label", applied.Text);
    }

    [Fact]
    public void A_textbox_whose_content_write_the_host_refuses_is_discarded_and_refused()
    {
        Mock<Excel.Shapes> shapes = new();
        FakeShape created = new("row-1:label");
        _ = shapes
            .Setup(s => s.AddTextbox(
                It.IsAny<MsoTextOrientation>(),
                It.IsAny<float>(),
                It.IsAny<float>(),
                It.IsAny<float>(),
                It.IsAny<float>()))
            .Returns(created.Shape);

        DepthWriter writer = Writer(shapes.Object);
        writer.TextWriteFailure = new InvalidOperationException("The host refused the text.");
        OfficeShapeRequest request = new(
            "row-1:label",
            OfficeShapeKind.TextBox,
            new OfficeShapeGeometry(Bounds: new RectD(40, 60, 80, 14)),
            ZLayer.Label,
            Text: "label");

        ShapeWriteOutcome outcome = writer.Create(request);

        Assert.Equal(ShapeWriteRefusal.HostRejected, outcome.Refusal);
        Assert.True(created.Deleted);
        // The stamp is never attempted on a shape that was just deleted.
        Assert.NotEqual("row-1:label", created.AlternativeText);
    }

    [Fact]
    public void A_diamond_is_created_as_an_auto_shape_at_its_bounding_box()
    {
        Mock<Excel.Shapes> shapes = new();
        FakeShape created = new("row-1:marker");
        MsoAutoShapeType kind = MsoAutoShapeType.msoShapeMixed;
        _ = shapes
            .Setup(s => s.AddShape(
                It.IsAny<MsoAutoShapeType>(),
                It.IsAny<float>(),
                It.IsAny<float>(),
                It.IsAny<float>(),
                It.IsAny<float>()))
            .Callback<MsoAutoShapeType, float, float, float, float>((type, _, _, _, _) => kind = type)
            .Returns(created.Shape);

        DepthWriter writer = Writer(shapes.Object);

        Assert.True(writer.Create(DiamondRequest()).Succeeded);
        Assert.Equal(MsoAutoShapeType.msoShapeDiamond, kind);
    }

    [Fact]
    public void A_host_that_refuses_the_create_call_reports_host_rejected()
    {
        // A loose Shapes mock returns null from AddShape, which is how the host
        // reports a refused creation; no setup is needed to model it.
        DepthWriter writer = Writer(ShapesCollection());

        Assert.Equal(ShapeWriteRefusal.HostRejected, writer.Create(Rectangle()).Refusal);
    }

    [Fact]
    public void A_created_shape_whose_tag_write_and_delete_both_fail_still_reports_a_refusal()
    {
        // DiscardUnownedShape is best-effort: a delete that itself fails must not
        // replace the typed refusal with an escaping exception.
        Mock<Excel.Shapes> shapes = new();
        FakeShape created = new() { TagWriteThrows = true, DeleteThrows = true };
        _ = shapes
            .Setup(s => s.AddShape(
                It.IsAny<MsoAutoShapeType>(),
                It.IsAny<float>(),
                It.IsAny<float>(),
                It.IsAny<float>(),
                It.IsAny<float>()))
            .Returns(created.Shape);

        DepthWriter writer = Writer(shapes.Object);

        ShapeWriteOutcome outcome = writer.Create(Rectangle());

        Assert.Equal(ShapeWriteRefusal.HostRejected, outcome.Refusal);
        Assert.False(created.Deleted);
    }

    [Fact]
    public void The_create_path_writes_a_fill_and_a_stroke_onto_the_shape()
    {
        Mock<Excel.Shapes> shapes = new();
        FakeShape created = new(withFillAndLine: true);
        _ = shapes
            .Setup(s => s.AddShape(
                It.IsAny<MsoAutoShapeType>(),
                It.IsAny<float>(),
                It.IsAny<float>(),
                It.IsAny<float>(),
                It.IsAny<float>()))
            .Returns(created.Shape);

        DepthWriter writer = Writer(shapes.Object);
        writer.CallBaseStyle = true;
        OfficeShapeRequest request = new(
            "row-1:bar",
            OfficeShapeKind.Rectangle,
            new OfficeShapeGeometry(Bounds: new RectD(10, 20, 100, 30)),
            ZLayer.ActivityBody,
            FillColour: ColourHex.Parse("#3366CC"),
            StrokeColour: ColourHex.Parse("#112233"),
            LineWidthPt: 1.5d,
            StrokeDash: GanttStrokeDash.Dash);

        ShapeWriteOutcome outcome = writer.Create(request);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.NotEqual(0, created.FillRgb);
        Assert.Equal(MsoTriState.msoTrue, created.FillVisible);
        Assert.Contains("Solid", created.FillCalls);
        Assert.NotEqual(0, created.LineRgb);
        Assert.Equal(1.5f, created.LineWeight);
        Assert.Equal(MsoLineDashStyle.msoLineDash, created.LineDash);
        Assert.Equal(MsoTriState.msoTrue, created.LineVisible);
    }

    [Fact]
    public void A_hatched_fill_is_written_as_a_pattern_rather_than_solid()
    {
        Mock<Excel.Shapes> shapes = new();
        FakeShape created = new(withFillAndLine: true);
        _ = shapes
            .Setup(s => s.AddShape(
                It.IsAny<MsoAutoShapeType>(),
                It.IsAny<float>(),
                It.IsAny<float>(),
                It.IsAny<float>(),
                It.IsAny<float>()))
            .Returns(created.Shape);

        DepthWriter writer = Writer(shapes.Object);
        writer.CallBaseStyle = true;
        OfficeShapeRequest request = new(
            "row-1:bar",
            OfficeShapeKind.Rectangle,
            new OfficeShapeGeometry(Bounds: new RectD(10, 20, 100, 30)),
            ZLayer.ActivityBody,
            FillColour: ColourHex.Parse("#3366CC"),
            HatchPattern: GanttHatchPattern.ForwardDiagonal);

        Assert.True(writer.Create(request).Succeeded);
        Assert.Contains(created.FillCalls, call => call.StartsWith("Patterned:", StringComparison.Ordinal));
    }

    [Fact]
    public void An_invisible_fill_and_stroke_are_written_off_explicitly()
    {
        Mock<Excel.Shapes> shapes = new();
        FakeShape created = new(withFillAndLine: true);
        _ = shapes
            .Setup(s => s.AddShape(
                It.IsAny<MsoAutoShapeType>(),
                It.IsAny<float>(),
                It.IsAny<float>(),
                It.IsAny<float>(),
                It.IsAny<float>()))
            .Returns(created.Shape);

        DepthWriter writer = Writer(shapes.Object);
        writer.CallBaseStyle = true;

        ShapeWriteOutcome outcome = writer.Create(Rectangle());

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal(MsoTriState.msoFalse, created.FillVisible);
        Assert.Equal(MsoTriState.msoFalse, created.LineVisible);
    }
}
