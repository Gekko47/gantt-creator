using GanttCreator.Core;
using GanttCreator.Core.Scene;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Moq;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for the R4.1 shape-write port. No live Excel: the adapter's
/// <c>internal virtual</c> seams are derived so the observable operation
/// sequence and the ownership members ADR-0019 defines are both asserted.
/// </summary>
/// <remarks>
/// The fake is deliberately narrow — it records what the adapter asked the host
/// to do, not a fake Excel object model (docs/04-TEST-STRATEGY.md "Excel shape
/// contract tests").
/// </remarks>
public class ExcelShapeWriterTests
{
    /// <summary>
    /// The writer with its COM seams substituted. Every host call the adapter
    /// makes is recorded so a test can assert sequence and values rather than
    /// inspect a proxy.
    /// </summary>
    private sealed class TestableWriter(
        object? application,
        IWorksheetProtectionGuard guard,
        params Excel.Shape[] existingShapes)
        : ExcelShapeWriter(application, guard)
    {
        private readonly List<Excel.Shape> _shapes = [.. existingShapes];

        /// <summary>The shapes the adapter created, in order.</summary>
        public List<Excel.Shape> Created { get; } = [];

        /// <summary>The z-order commands the adapter issued.</summary>
        public List<string> ZOrderCommands { get; } = [];

        /// <summary>
        /// The requests the adapter was asked to create, in order. R4.4 asserts
        /// the text, font, and alignment members the host was handed, so the
        /// observable operation is recorded rather than the shape's post-write
        /// state, which a fake would have to model to be worth anything.
        /// </summary>
        public List<OfficeShapeRequest> Requests { get; } = [];

        internal override Excel._Worksheet? FindGanttWorksheet(Excel.Sheets sheets) =>
            new Mock<Excel._Worksheet>().Object;

        internal override Excel.Shapes? GetShapes(Excel._Worksheet sheet) =>
            new Mock<Excel.Shapes>().Object;

        internal override int GetShapeCount(Excel.Shapes shapes) => _shapes.Count;

        internal override Excel.Shape? GetShapeAt(Excel.Shapes shapes, int index) =>
            index >= 1 && index <= _shapes.Count ? _shapes[index - 1] : null;

        internal override Excel.Shape? FindShapeByName(Excel.Shapes shapes, string name) =>
            _shapes.FirstOrDefault(shape => string.Equals(shape.Name, name, StringComparison.Ordinal));

        internal override Excel.Shape? AddShape(Excel.Shapes shapes, OfficeShapeRequest request)
        {
            Requests.Add(request);

            // The real adapter now implements the polygon kind through
            // BuildFreeform, which is exercised by the tagged live-Office
            // integration test. The fake records the request and returns a
            // shape so count/type/order pins can be asserted without a host.
            var mock = new Mock<Excel.Shape>();
            TestableWriter owner = this;
            string name = string.Empty;
            string alternativeText = string.Empty;
            _ = mock.SetupSet(s => s.Name = It.IsAny<string>()).Callback<string>(value => name = value);
            _ = mock.SetupGet(s => s.Name).Returns(() => name);
            _ = mock.SetupSet(s => s.AlternativeText = It.IsAny<string>())
                .Callback<string>(value => alternativeText = value);
            _ = mock.SetupGet(s => s.AlternativeText).Returns(() => alternativeText);

            // ZOrder is a real COM call on the shape; the fake records the
            // command instead, so D2's "no opportunistic BringToFront" pin
            // asserts the actual operation sequence rather than a proxy.
            _ = mock.Setup(s => s.ZOrder(It.IsAny<MsoZOrderCmd>()))
                .Callback<MsoZOrderCmd>(command => owner.ZOrderCommands.Add(command.ToString()));

            var created = mock.Object;
            Created.Add(created);
            _shapes.Add(created);
            return created;
        }
    }

    private static Mock<IWorksheetProtectionGuard> ClearGuard()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        _ = guard.Setup(g => g.QueryTarget(It.IsAny<object>())).Returns(ProtectionGuardOutcome.NotProtected);
        return guard;
    }

    private static Mock<Excel.Application> ActiveApplication()
    {
        var workbook = new Mock<Excel.Workbook>();
        _ = workbook.Setup(w => w.Sheets).Returns(new Mock<Excel.Sheets>().Object);
        var application = new Mock<Excel.Application>();
        _ = application.Setup(a => a.ActiveWorkbook).Returns(workbook.Object);
        return application;
    }

    private static OfficeShapeRequest RectangleRequest(
        string id = "row-1:bar",
        double x = 10,
        double y = 20,
        double width = 100,
        double height = 30) =>
        new(id, OfficeShapeKind.Rectangle, new OfficeShapeGeometry(Bounds: new RectD(x, y, width, height)), ZLayer.ActivityBody);

    [Fact]
    public void Create_writes_the_primitive_id_as_the_shape_name_and_the_tag_as_the_alternative_text()
    {
        var writer = new TestableWriter(ActiveApplication().Object, ClearGuard().Object);

        ShapeWriteOutcome outcome = writer.Create(RectangleRequest());

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Excel.Shape created = Assert.Single(writer.Created);
        Assert.Equal("row-1:bar", created.Name);
        Assert.Equal(ShapeOwnershipTag.ForPrimitiveId("row-1:bar"), created.AlternativeText);
        Assert.StartsWith(ShapeOwnershipTag.Prefix, created.AlternativeText, StringComparison.Ordinal);
    }

    [Fact]
    public void Create_refuses_a_blank_primitive_id_before_touching_the_host()
    {
        var writer = new TestableWriter(ActiveApplication().Object, ClearGuard().Object);

        ShapeWriteOutcome outcome = writer.Create(RectangleRequest(id: "   "));

        Assert.Equal(ShapeWriteRefusal.BlankIdentifier, outcome.Refusal);
        Assert.Empty(writer.Created);
    }

    [Fact]
    public void Create_refuses_a_rectangle_with_no_bounds_rather_than_defaulting_them_to_the_origin()
    {
        var writer = new TestableWriter(ActiveApplication().Object, ClearGuard().Object);
        var request = new OfficeShapeRequest(
            "row-1:bar", OfficeShapeKind.Rectangle, new OfficeShapeGeometry(), ZLayer.ActivityBody);

        ShapeWriteOutcome outcome = writer.Create(request);

        Assert.Equal(ShapeWriteRefusal.InvalidGeometry, outcome.Refusal);
        Assert.Empty(writer.Created);
    }

    [Fact]
    public void Create_refuses_a_line_with_a_non_finite_endpoint()
    {
        var writer = new TestableWriter(ActiveApplication().Object, ClearGuard().Object);
        var request = new OfficeShapeRequest(
            "row-1:critical",
            OfficeShapeKind.Line,
            new OfficeShapeGeometry(From: new PointD(double.NaN, 0), To: new PointD(10, 10)),
            ZLayer.CriticalOverlay);

        ShapeWriteOutcome outcome = writer.Create(request);

        Assert.Equal(ShapeWriteRefusal.InvalidGeometry, outcome.Refusal);
        Assert.Empty(writer.Created);
    }
    /// <summary>
    /// R4.5 D1: a four-point milestone polygon becomes a freeform request whose
    /// points survive in draw order, so the adapter draws the diamond the scene
    /// resolved rather than one it re-derived.
    /// </summary>
    [Fact]
    public void A_four_point_polygon_becomes_a_freeform_request_preserving_its_points()
    {
        var writer = new TestableWriter(ActiveApplication().Object, ClearGuard().Object);
        PointD[] diamond = [new(50, 40), new(60, 50), new(50, 60), new(40, 50)];
        var request = new OfficeShapeRequest(
            "row-1:marker",
            OfficeShapeKind.Polygon,
            new OfficeShapeGeometry(Points: diamond),
            ZLayer.Milestone);

        ShapeWriteOutcome outcome = writer.Create(request);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        OfficeShapeRequest created = Assert.Single(writer.Requests);
        Assert.Equal(OfficeShapeKind.Polygon, created.Kind);
        Assert.Equal(diamond, created.Geometry.Points);

        // The shape is owned and named, so a diamond is reconcilable exactly
        // like a bar or a label.
        Excel.Shape shape = Assert.Single(writer.Created);
        Assert.Equal("row-1:marker", shape.Name);
        Assert.Equal(ShapeOwnershipTag.ForPrimitiveId("row-1:marker"), shape.AlternativeText);
    }

    /// <summary>
    /// A polygon with fewer than three points cannot enclose an area, so it is
    /// refused before the host is touched rather than placed as a degenerate
    /// shape. This is the validator's positive test.
    /// </summary>
    [Fact]
    public void A_polygon_with_fewer_than_three_points_is_refused()
    {
        var writer = new TestableWriter(ActiveApplication().Object, ClearGuard().Object);
        var request = new OfficeShapeRequest(
            "row-1:marker",
            OfficeShapeKind.Polygon,
            new OfficeShapeGeometry(Points: [new PointD(50, 40), new PointD(60, 50)]),
            ZLayer.Milestone);

        ShapeWriteOutcome outcome = writer.Create(request);

        Assert.Equal(ShapeWriteRefusal.InvalidGeometry, outcome.Refusal);
        Assert.Empty(writer.Created);
    }

    /// <summary>
    /// R4.5 D2: the z-order pass moves the first named shape to the back and then
    /// brings each subsequent one forward exactly once. No per-shape
    /// <c>BringToFront</c> outside this single pass, and no host-default order.
    /// </summary>
    [Fact]
    public void The_z_order_pass_sends_the_first_shape_to_the_back_then_brings_each_forward_once()
    {
        var writer = new TestableWriter(ActiveApplication().Object, ClearGuard().Object);
        string[] ids = ["chart:background", "chart:grid:0", "row-1:bar", "row-1:label"];
        foreach (string id in ids)
        {
            _ = writer.Create(new OfficeShapeRequest(
                id,
                OfficeShapeKind.Rectangle,
                new OfficeShapeGeometry(Bounds: new RectD(0, 0, 10, 10)),
                ZLayer.Background));
        }

        writer.ZOrderCommands.Clear();

        ShapeWriteOutcome outcome = writer.ApplyZOrder(ids);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());

        // Exactly one send-to-back and one bring-forward per shape after the
        // first - no opportunistic repetition, no BringToFront at all.
        Assert.Equal(
            1 + (ids.Length - 1),
            writer.ZOrderCommands.Count(command => command != "BringToFront"));
        Assert.DoesNotContain("BringToFront", writer.ZOrderCommands);
    }

    /// <summary>
    /// The z-order pass refuses when a named shape does not exist, rather than
    /// applying a partial order that would leave the chart in a state neither the
    /// scene nor the user asked for.
    /// </summary>
    [Fact]
    public void The_z_order_pass_refuses_rather_than_applying_a_partial_order()
    {
        var writer = new TestableWriter(ActiveApplication().Object, ClearGuard().Object);
        _ = writer.Create(new OfficeShapeRequest(
            "chart:background",
            OfficeShapeKind.Rectangle,
            new OfficeShapeGeometry(Bounds: new RectD(0, 0, 10, 10)),
            ZLayer.Background));

        ShapeWriteOutcome outcome = writer.ApplyZOrder(["chart:background", "row-1:missing"]);

        Assert.Equal(ShapeWriteRefusal.NotFound, outcome.Refusal);
    }

    /// <summary>
    /// R4.4 D3: the alignment mapping is a closed three-member table, and an
    /// alignment outside the enum is REFUSED rather than defaulted. A silent
    /// fallback would render a label with an alignment nobody chose, which is
    /// the same class of unrequested visual change the entity guide forbids.
    /// </summary>
    [Fact]
    public void An_alignment_outside_the_scene_enum_is_refused_rather_than_defaulted()
    {
        var method = typeof(ExcelShapeWriter).GetMethod(
            "MapAlignment",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);

        Assert.NotNull(method);

        var undefined = (GanttTextAlignment)99;
        var thrown = Assert.Throws<System.Reflection.TargetInvocationException>(
            () => method!.Invoke(null, [undefined]));

        Assert.IsType<ArgumentOutOfRangeException>(thrown.InnerException);
    }

    /// <summary>
    /// The three defined alignments each map to a distinct host constant, and
    /// none of them collapses onto another's value.
    /// </summary>
    [Fact]
    public void Each_defined_alignment_maps_to_a_distinct_host_constant()
    {
        var method = typeof(ExcelShapeWriter).GetMethod(
            "MapAlignment",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);

        Assert.NotNull(method);

        var mapped = Enum.GetValues<GanttTextAlignment>()
            .Select(alignment => method!.Invoke(null, [alignment]))
            .ToArray();

        Assert.Equal(Enum.GetValues<GanttTextAlignment>().Length, mapped.Length);
        Assert.Equal(mapped.Length, mapped.Distinct().Count());
        Assert.DoesNotContain(null, mapped);
    }

    /// <summary>
    /// R4.4 D4: an over-long label the scene already truncated is written to the
    /// shape verbatim, ellipsis included. The adapter applies no truncation of
    /// its own, so what the host receives is exactly what the scene resolved.
    /// </summary>
    [Fact]
    public void A_clipped_label_reaches_the_host_verbatim_with_no_second_truncation()
    {
        var writer = new TestableWriter(ActiveApplication().Object, ClearGuard().Object);
        const string Clipped = "Site survey and ground investigation phase one…";
        var request = new OfficeShapeRequest(
            "row-1:label",
            OfficeShapeKind.TextBox,
            new OfficeShapeGeometry(Bounds: new RectD(10, 20, 40, 12)),
            ZLayer.Label,
            Text: Clipped,
            Alignment: GanttTextAlignment.Right);

        ShapeWriteOutcome outcome = writer.Create(request);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        OfficeShapeRequest created = Assert.Single(writer.Requests);
        Assert.Equal(Clipped, created.Text);
        Assert.Equal(GanttTextAlignment.Right, created.Alignment);
    }

    [Fact]
    public void Create_refuses_a_duplicate_name_rather_than_creating_a_second_shape()
    {
        var existing = NewShape("row-1:bar", ShapeOwnershipTag.ForPrimitiveId("row-1:bar"));
        var writer = new TestableWriter(ActiveApplication().Object, ClearGuard().Object, existing);

        ShapeWriteOutcome outcome = writer.Create(RectangleRequest());

        Assert.Equal(ShapeWriteRefusal.AlreadyExists, outcome.Refusal);
        Assert.Empty(writer.Created);
    }

    [Fact]
    public void Create_refuses_before_any_mutation_when_the_worksheet_is_protected()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.SheetProtected);
        var writer = new TestableWriter(ActiveApplication().Object, guard.Object);

        ShapeWriteOutcome outcome = writer.Create(RectangleRequest());

        Assert.Equal(ShapeWriteRefusal.TargetProtected, outcome.Refusal);
        Assert.Empty(writer.Created);
    }

    [Fact]
    public void Create_refuses_when_the_target_worksheet_alone_is_protected()
    {
        // The active sheet is clear but the resolved Gantt target is protected.
        // A guard that only read the active sheet would authorise this write.
        var guard = ClearGuard();
        _ = guard.Setup(g => g.QueryTarget(It.IsAny<object>())).Returns(ProtectionGuardOutcome.SheetProtected);
        var writer = new TestableWriter(ActiveApplication().Object, guard.Object);

        ShapeWriteOutcome outcome = writer.Create(RectangleRequest());

        Assert.Equal(ShapeWriteRefusal.TargetProtected, outcome.Refusal);
        Assert.Empty(writer.Created);
    }

    [Fact]
    public void Every_entry_point_refuses_without_an_application_object()
    {
        var writer = new TestableWriter(null, ClearGuard().Object);

        Assert.Equal(ShapeWriteRefusal.NoActiveWorkbook, writer.Create(RectangleRequest()).Refusal);
        Assert.Equal(ShapeWriteRefusal.NoActiveWorkbook, writer.Update(RectangleRequest()).Refusal);
        Assert.Equal(ShapeWriteRefusal.NoActiveWorkbook, writer.Delete("row-1:bar").Refusal);
        Assert.Empty(writer.ListOwned());
    }

    [Fact]
    public void Delete_never_deletes_a_shape_whose_alternative_text_is_not_a_valid_tag()
    {
        // R4.8 D1: the ownership filter authorises the delete, not the name. The
        // sentinel deliberately carries the requested name.
        bool deleted = false;
        var foreign = NewShape("row-1:bar", "a note the user typed", onDelete: () => deleted = true);
        var writer = new TestableWriter(ActiveApplication().Object, ClearGuard().Object, foreign);

        ShapeWriteOutcome outcome = writer.Delete("row-1:bar");

        Assert.Equal(ShapeWriteRefusal.NotFound, outcome.Refusal);
        Assert.False(deleted, "An unowned shape must never be deleted.");
    }

    [Fact]
    public void ListOwned_returns_only_validly_tagged_shapes_in_ordinal_name_order()
    {
        var writer = new TestableWriter(
            ActiveApplication().Object,
            ClearGuard().Object,
            NewShape("z-last", ShapeOwnershipTag.ForPrimitiveId("z-last")),
            NewShape("user-shape", null),
            NewShape("a-first", ShapeOwnershipTag.ForPrimitiveId("a-first")),
            NewShape("broken", ShapeOwnershipTag.Prefix + "not-a-hash"));

        Assert.Equal(["a-first", "z-last"], writer.ListOwned());
    }

    private static Excel.Shape NewShape(string name, string? alternativeText, Action? onDelete = null)
    {
        var shape = new Mock<Excel.Shape>();
        _ = shape.SetupGet(s => s.Name).Returns(name);
        _ = shape.SetupGet(s => s.AlternativeText).Returns(alternativeText ?? string.Empty);
        _ = shape.Setup(s => s.Delete()).Callback(() => onDelete?.Invoke());
        return shape.Object;
    }
}