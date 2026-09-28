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
        /// The shape names the z-order commands were issued against, in call order.
        /// A command sequence alone proves <em>how many</em> operations ran, not
        /// <em>which</em> shape each one moved; the reverse pass is only correct if
        /// the front-most shape is the first one sent back, so the shape identity
        /// is recorded too.
        /// </summary>
        public List<string> ZOrderShapes { get; } = [];

        /// <summary>
        /// The requests the adapter was asked to create, in order. R4.4 asserts
        /// the text, font, and alignment members the host was handed, so the
        /// observable operation is recorded rather than the shape's post-write
        /// state, which a fake would have to model to be worth anything.
        /// </summary>
        public List<OfficeShapeRequest> Requests { get; } = [];

        /// <summary>
        /// Whether writing the alternative-text ownership tag throws, simulating a
        /// host that refuses the stamp (KNOWN-LIMITATIONS L18's over-long shape name
        /// is the live instance of this).
        /// </summary>
        public bool FailOwnershipTagWrite { get; set; }

        /// <summary>Whether a shape the adapter created was subsequently deleted.</summary>
        public bool CreatedShapeDeleted { get; private set; }

        internal override Excel._Worksheet? FindGanttWorksheet(Excel.Sheets sheets) =>
            new Mock<Excel._Worksheet>().Object;

        internal override Excel.Shapes? GetShapes(Excel._Worksheet sheet) =>
            new Mock<Excel.Shapes>().Object;

        internal override int GetShapeCount(Excel.Shapes shapes) => _shapes.Count;

        internal override Excel.Shape? GetShapeAt(Excel.Shapes shapes, int index) =>
            index >= 1 && index <= _shapes.Count ? _shapes[index - 1] : null;

        internal override Excel.Shape? FindShapeByName(Excel.Shapes shapes, string name) =>
            _shapes.FirstOrDefault(shape => string.Equals(shape.Name, name, StringComparison.Ordinal));

        /// <summary>
        /// Records the content write instead of walking the live
        /// <c>TextFrame2</c> COM chain, so an update's content application is
        /// observable without a host.
        /// </summary>
        /// <param name="shape">The shape whose content would be written.</param>
        /// <param name="request">The content that would be written.</param>
        internal override void ApplyText(Excel.Shape shape, OfficeShapeRequest request) =>
            Requests.Add(request);

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
            if (FailOwnershipTagWrite)
            {
                // The host rejects the stamp (L18's over-long name is the live
                // case), so the discard path in the adapter must run.
                _ = mock.SetupSet(s => s.AlternativeText = It.IsAny<string>())
                    .Throws(new ArgumentException("The specified value is out of range."));
            }

            _ = mock.SetupGet(s => s.AlternativeText).Returns(() => alternativeText);

            _ = mock.Setup(s => s.Delete()).Callback(() => CreatedShapeDeleted = true);

            // ZOrder is a real COM call on the shape; the fake records the
            // command and the shape it moved instead, so D2's "no opportunistic
            // BringToFront" pin and the reverse-order pin both assert the actual
            // operation sequence rather than a proxy.
            _ = mock.Setup(s => s.ZOrder(It.IsAny<MsoZOrderCmd>()))
                .Callback<MsoZOrderCmd>(command =>
                {
                    owner.ZOrderCommands.Add(command.ToString());
                    owner.ZOrderShapes.Add(name);
                });

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
    /// R4.5 D1: a milestone is placed as a diamond auto-shape at the bounding
    /// box of its points. The points stay on the request so the translation is
    /// auditable, and the box is what the adapter places the shape from.
    /// </summary>
    [Fact]
    public void A_milestone_becomes_a_diamond_request_carrying_its_points_and_their_box()
    {
        var writer = new TestableWriter(ActiveApplication().Object, ClearGuard().Object);
        PointD[] diamond = [new(50, 40), new(60, 50), new(50, 60), new(40, 50)];
        var request = new OfficeShapeRequest(
            "row-1:marker",
            OfficeShapeKind.Diamond,
            new OfficeShapeGeometry(Bounds: new RectD(40, 40, 20, 20), Points: diamond),
            ZLayer.Milestone);

        ShapeWriteOutcome outcome = writer.Create(request);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        OfficeShapeRequest created = Assert.Single(writer.Requests);
        Assert.Equal(OfficeShapeKind.Diamond, created.Kind);
        Assert.Equal(new RectD(40, 40, 20, 20), created.Geometry.Bounds);
        Assert.Equal(diamond, created.Geometry.Points);

        // The shape is owned and named, so a diamond is reconcilable exactly
        // like a bar or a label.
        Excel.Shape shape = Assert.Single(writer.Created);
        Assert.Equal("row-1:marker", shape.Name);
        Assert.Equal(ShapeOwnershipTag.ForPrimitiveId("row-1:marker"), shape.AlternativeText);
    }

    /// <summary>
    /// A diamond needs its box, so a request carrying points but no box is
    /// refused rather than placed at the origin. This is the validator's
    /// positive test.
    /// </summary>
    [Fact]
    public void A_diamond_without_bounds_is_refused_rather_than_placed_at_the_origin()
    {
        var writer = new TestableWriter(ActiveApplication().Object, ClearGuard().Object);
        var request = new OfficeShapeRequest(
            "row-1:marker",
            OfficeShapeKind.Diamond,
            new OfficeShapeGeometry(Points: [new PointD(50, 40), new PointD(60, 50), new PointD(50, 60)]),
            ZLayer.Milestone);

        ShapeWriteOutcome outcome = writer.Create(request);

        Assert.Equal(ShapeWriteRefusal.InvalidGeometry, outcome.Refusal);
        Assert.Empty(writer.Created);
    }

    /// <summary>
    /// R4.4: an update re-applies a text box's content, not just its geometry.
    /// A refresh that changed a label's string, font, or alignment would
    /// otherwise move the box and leave the stale text on the sheet.
    /// </summary>
    [Fact]
    public void An_update_re_applies_a_labels_text_font_and_alignment()
    {
        var existing = NewShape("row-1:label", ShapeOwnershipTag.ForPrimitiveId("row-1:label"));
        var writer = new TestableWriter(ActiveApplication().Object, ClearGuard().Object, existing);

        var request = new OfficeShapeRequest(
            "row-1:label",
            OfficeShapeKind.TextBox,
            new OfficeShapeGeometry(Bounds: new RectD(40, 60, 80, 14)),
            ZLayer.Label,
            FontFamily: "Aptos",
            FontSizePt: 11d,
            Bold: true,
            Text: "Revised label text",
            Alignment: GanttTextAlignment.Centre);

        ShapeWriteOutcome outcome = writer.Update(request);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());

        // The request reached the host as a whole: the update path must not
        // silently drop the members the create path writes.
        OfficeShapeRequest applied = Assert.Single(writer.Requests);
        Assert.Equal("Revised label text", applied.Text);
        Assert.Equal("Aptos", applied.FontFamily);
        Assert.Equal(11d, applied.FontSizePt);
        Assert.True(applied.Bold);
        Assert.Equal(GanttTextAlignment.Centre, applied.Alignment);
    }

    /// <summary>
    /// The R4.1 guarantee still holds on the path just changed: an update
    /// rewrites content but never the name or the ownership tag, because those
    /// are the reconciliation key and the ownership proof.
    /// </summary>
    /// <remarks>
    /// A user-edited tag - a shape that keeps the name but whose alternative text
    /// the user replaced - is now <em>refused</em> rather than updated. That is
    /// strictly stronger than the original "do not re-stamp" rule, and it is what
    /// the entity guide requires: refresh touches only shapes whose alternative
    /// text carries a valid ownership tag. Both members are still left alone.
    /// </remarks>
    [Fact]
    public void An_update_never_rewrites_the_name_or_the_ownership_tag()
    {
        var existing = NewShape("row-1:label", ShapeOwnershipTag.ForPrimitiveId("row-1:label"));
        var writer = new TestableWriter(ActiveApplication().Object, ClearGuard().Object, existing);

        // A user-edited shape: the alternative text no longer matches the name.
        const string UserEditedTag = "my own note about this shape";
        existing.AlternativeText = UserEditedTag;

        ShapeWriteOutcome outcome = writer.Update(
            new OfficeShapeRequest(
                "row-1:label",
                OfficeShapeKind.TextBox,
                new OfficeShapeGeometry(Bounds: new RectD(10, 10, 40, 12)),
                ZLayer.Label,
                Text: "New text"));

        Assert.False(outcome.Succeeded);
        Assert.Equal("row-1:label", existing.Name);
        Assert.Equal(UserEditedTag, existing.AlternativeText);
    }

    /// <summary>
    /// A text box with no bounds is refused by the geometry validator before the
    /// content path runs, so the re-application can never dereference a missing
    /// box. This is the positive test for that ordering.
    /// </summary>
    [Fact]
    public void A_text_box_with_no_bounds_is_refused_before_its_content_is_applied()
    {
        var existing = NewShape("row-1:label", ShapeOwnershipTag.ForPrimitiveId("row-1:label"));
        var writer = new TestableWriter(ActiveApplication().Object, ClearGuard().Object, existing);

        ShapeWriteOutcome outcome = writer.Update(
            new OfficeShapeRequest(
                "row-1:label",
                OfficeShapeKind.TextBox,
                new OfficeShapeGeometry(),
                ZLayer.Label,
                Text: "New text"));

        Assert.Equal(ShapeWriteRefusal.InvalidGeometry, outcome.Refusal);

        // Nothing was requested of the host, so the content path did not run.
        Assert.Empty(writer.Requests);
    }

    /// <summary>
    /// A non-text update is unaffected: only a text box has content, so a
    /// rectangle update must not reach the text path at all.
    /// </summary>
    [Fact]
    public void A_rectangle_update_does_not_reach_the_text_path()
    {
        var existing = NewShape("row-1:bar", ShapeOwnershipTag.ForPrimitiveId("row-1:bar"));
        var writer = new TestableWriter(ActiveApplication().Object, ClearGuard().Object, existing);

        ShapeWriteOutcome outcome = writer.Update(
            new OfficeShapeRequest(
                "row-1:bar",
                OfficeShapeKind.Rectangle,
                new OfficeShapeGeometry(Bounds: new RectD(10, 20, 100, 30)),
                ZLayer.ActivityBody,
                Text: "text a rectangle must never receive"));

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Empty(writer.Requests);
    }

    /// <summary>
    /// R4.5 D2: the z-order pass sends each named shape to the back in REVERSE
    /// order, which is what makes the host's final order equal the scene's
    /// back-to-front list whatever order the host started in. <c>msoSendToBack</c>
    /// is absolute, so walking front-to-back and pushing every shape behind
    /// everything else leaves the list in order; walking it the other way would
    /// produce its exact reverse. No per-shape <c>BringToFront</c> outside this
    /// single pass, and no host-default order.
    /// </summary>
    [Fact]
    public void The_z_order_pass_sends_each_shape_to_the_back_in_reverse_so_the_final_order_matches_the_scene()
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

        // Create issues no z-order command of its own: insertion order already is
        // the scene's order, and ApplyZOrder is the single authoritative pass.
        Assert.Empty(writer.ZOrderCommands);

        ShapeWriteOutcome outcome = writer.ApplyZOrder(ids);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());

        // One send-to-back per shape, issued front-to-back (i.e. the reverse of the
        // backToFront list). Anything else - a bring-forward pass, a BringToFront,
        // or fewer calls than shapes - would leave the host's order wrong. The
        // literal is the enum member's own name, as the recorder writes it.
        Assert.Equal(ids.Length, writer.ZOrderCommands.Count);
        Assert.All(
            writer.ZOrderCommands,
            command => Assert.Equal(nameof(MsoZOrderCmd.msoSendToBack), command));

        // The shapes are recorded in the order they were sent to the back, so the
        // sequence itself is the assertion: the front-most scene shape goes first.
        Assert.Equal(
            [.. Enumerable.Reverse(ids)],
            writer.ZOrderShapes);
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

        // The preflight resolves everything before issuing the first command, so
        // the one shape that *does* exist was not moved. Asserting the absence of
        // commands is what distinguishes an atomic refusal from a partial one.
        Assert.Empty(writer.ZOrderCommands);
    }

    /// <summary>
    /// A user-drawn shape that happens to carry a requested scene name is NOT
    /// reordered, because <c>ApplyZOrder</c> now authorises on the ownership tag
    /// exactly as <c>Update</c>, <c>Delete</c>, and <c>ListOwned</c> already do.
    /// </summary>
    /// <remarks>
    /// The unowned shape is deliberately <em>first</em> in the list, so a check
    /// placed inside the mutation loop would already have started issuing commands
    /// before reaching it. The refusal is only correct because the whole list is
    /// preflighted. The command log is attached to the seeded shapes rather than
    /// taken from <c>writer.ZOrderCommands</c>, because that recorder is wired only
    /// into shapes <c>AddShape</c> creates - asserting against it here would pass
    /// whether or not any command was issued. This test was mutation-checked:
    /// moving the ownership test into the mutation loop makes it fail.
    /// </remarks>
    [Fact]
    public void The_z_order_pass_refuses_an_unowned_same_name_shape_and_moves_nothing()
    {
        // A user shape: the name matches the scene, the alternative text does not
        // carry a GanttCreator ownership tag.
        List<string> moved = [];
        var userShape = NewShape("row-1:bar", "a rectangle the user drew themselves", onZOrder: (n, _) => moved.Add(n));
        var owned = NewShape("row-1:label", ShapeOwnershipTag.ForPrimitiveId("row-1:label"), onZOrder: (n, _) => moved.Add(n));
        var writer = new TestableWriter(
            ActiveApplication().Object,
            ClearGuard().Object,
            userShape,
            owned);

        ShapeWriteOutcome outcome = writer.ApplyZOrder(["row-1:bar", "row-1:label"]);

        Assert.False(outcome.Succeeded);
        Assert.Equal(ShapeWriteRefusal.NotFound, outcome.Refusal);
        // Zero, not one: the unowned shape is first, so a per-shape check inside
        // the loop would already have moved the legitimately-owned shape.
        Assert.Empty(moved);
    }

    /// <summary>
    /// The atomicity case the existence preflight already provided, now also
    /// required of the ownership check: an unowned shape refused <em>later</em> in
    /// the list leaves the earlier, legitimately-owned shapes untouched.
    /// </summary>
    /// <summary>
    /// The atomicity case. An unowned shape in the **middle** of the list must
    /// leave the shapes on either side of it untouched.
    /// </summary>
    /// <remarks>
    /// Middle, not last, and that is deliberate. The approved pass walks the list in
    /// reverse, so a per-shape check inside the mutation loop processes the *last*
    /// entry first: an unowned shape placed last would refuse before issuing
    /// anything, and the test would pass whether or not the check was preflighted.
    /// With the unowned shape in the middle, a per-shape check has already moved
    /// the third shape by the time it refuses, so this test fails unless the whole
    /// list is preflighted. Mutation-checked in both directions.
    /// </remarks>
    [Fact]
    public void An_unowned_shape_in_the_middle_leaves_the_surrounding_owned_shapes_untouched()
    {
        List<string> moved = [];
        var first = NewShape("chart:background", ShapeOwnershipTag.ForPrimitiveId("chart:background"), onZOrder: (n, _) => moved.Add(n));
        var userShape = NewShape("row-1:bar", "a rectangle the user drew themselves", onZOrder: (n, _) => moved.Add(n));
        var last = NewShape("row-1:label", ShapeOwnershipTag.ForPrimitiveId("row-1:label"), onZOrder: (n, _) => moved.Add(n));
        var writer = new TestableWriter(
            ActiveApplication().Object,
            ClearGuard().Object,
            first,
            userShape,
            last);

        ShapeWriteOutcome outcome = writer.ApplyZOrder(["chart:background", "row-1:bar", "row-1:label"]);

        Assert.False(outcome.Succeeded);
        Assert.Equal(ShapeWriteRefusal.NotFound, outcome.Refusal);
        // Nothing at all, including the two legitimately-owned shapes: a partial
        // order would leave the chart in a state neither the scene nor the user
        // asked for.
        Assert.Empty(moved);
    }

    /// <summary>
    /// A shape carrying <em>another</em> primitive's valid tag is equally unowned
    /// for this request. <c>IsOwnedTag</c> would accept it; only
    /// <c>CarriesOwnershipTagFor</c> refuses it, which is the same distinction
    /// <c>Update</c> and <c>Delete</c> make.
    /// </summary>
    [Fact]
    public void A_shape_carrying_another_primitives_valid_tag_is_still_refused_by_the_z_order_pass()
    {
        List<string> moved = [];
        var impostor = NewShape(
            "row-1:bar",
            ShapeOwnershipTag.ForPrimitiveId("row-9:something-else"),
            onZOrder: (n, _) => moved.Add(n));
        var writer = new TestableWriter(ActiveApplication().Object, ClearGuard().Object, impostor);

        ShapeWriteOutcome outcome = writer.ApplyZOrder(["row-1:bar"]);

        Assert.False(outcome.Succeeded);
        Assert.Equal(ShapeWriteRefusal.NotFound, outcome.Refusal);
        Assert.Empty(moved);
    }

    /// <summary>
    /// The all-owned path still succeeds, and the approved single reverse-order
    /// <c>msoSendToBack</c> pass is untouched: the ownership preflight changes
    /// <em>who</em> may be reordered, never <em>how</em>.
    /// </summary>
    [Fact]
    public void The_z_order_pass_still_applies_the_approved_sequence_when_every_shape_is_owned()
    {
        List<(string Name, MsoZOrderCmd Command)> log = [];
        var background = NewShape(
            "chart:background",
            ShapeOwnershipTag.ForPrimitiveId("chart:background"),
            onZOrder: (name, command) => log.Add((name, command)));
        var bar = NewShape(
            "row-1:bar",
            ShapeOwnershipTag.ForPrimitiveId("row-1:bar"),
            onZOrder: (name, command) => log.Add((name, command)));
        var writer = new TestableWriter(ActiveApplication().Object, ClearGuard().Object, background, bar);

        ShapeWriteOutcome outcome = writer.ApplyZOrder(["chart:background", "row-1:bar"]);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        // Still one send-to-back per shape, in reverse: the ownership preflight
        // changes *who* may be reordered, never *how*.
        Assert.Equal(2, log.Count);
        Assert.All(log, entry => Assert.Equal(nameof(MsoZOrderCmd.msoSendToBack), entry.Command.ToString()));
        Assert.Equal(["row-1:bar", "chart:background"], [.. log.Select(entry => entry.Name)]);
    }

    /// <summary>
    /// A user shape the scene never names is untouched, which is the R4.8
    /// preservation guarantee the ownership filter exists to protect. Re-asserted
    /// here because <c>ApplyZOrder</c> is now a mutating path that could reach it.
    /// </summary>
    [Fact]
    public void A_user_shape_the_scene_never_names_is_untouched_by_the_z_order_pass()
    {
        List<string> moved = [];
        var owned = NewShape(
            "row-1:bar",
            ShapeOwnershipTag.ForPrimitiveId("row-1:bar"),
            onZOrder: (name, _) => moved.Add(name));
        var unrelated = NewShape("User drawing", string.Empty, onZOrder: (name, _) => moved.Add(name));
        var writer = new TestableWriter(ActiveApplication().Object, ClearGuard().Object, owned, unrelated);

        ShapeWriteOutcome outcome = writer.ApplyZOrder(["row-1:bar"]);

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal(["row-1:bar"], moved);
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

    /// <summary>
    /// The ownership filter authorises the update, exactly as it authorises the
    /// delete. A user shape that happens to carry the requested name must not be
    /// moved or rewritten, because a refresh reaching it is precisely the R4.8
    /// preservation failure. The sentinel deliberately carries the requested name.
    /// </summary>
    [Fact]
    public void Update_never_rewrites_a_shape_whose_alternative_text_is_not_a_valid_tag()
    {
        var foreign = NewShape("row-1:bar", "a note the user typed");
        var writer = new TestableWriter(ActiveApplication().Object, ClearGuard().Object, foreign);

        ShapeWriteOutcome outcome = writer.Update(RectangleRequest());

        Assert.Equal(ShapeWriteRefusal.NotFound, outcome.Refusal);
        Assert.Empty(writer.Requests);
    }

    /// <summary>
    /// The positive half of the same rule: a validly tagged shape is still
    /// updated, so the ownership filter refuses unowned shapes rather than
    /// refusing updates in general.
    /// </summary>
    [Fact]
    public void Update_still_rewrites_a_shape_carrying_a_valid_ownership_tag()
    {
        var existing = NewShape("row-1:bar", ShapeOwnershipTag.ForPrimitiveId("row-1:bar"));
        var writer = new TestableWriter(ActiveApplication().Object, ClearGuard().Object, existing);

        ShapeWriteOutcome outcome = writer.Update(RectangleRequest(x: 55, y: 66));

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Equal(55f, existing.Left);
        Assert.Equal(66f, existing.Top);
    }

    /// <summary>
    /// L18: Excel refuses a shape name beyond roughly 255 characters, and the
    /// name carries the full primitive identifier. Create must refuse it BEFORE
    /// <c>AddShape</c>, or the shape is placed on the sheet and only then fails on
    /// the name write - leaving debris the caller was told nothing about.
    /// </summary>
    [Fact]
    public void Create_refuses_an_identifier_longer_than_the_host_name_limit_before_creating_anything()
    {
        var writer = new TestableWriter(ActiveApplication().Object, ClearGuard().Object);
        var tooLong = new string('r', ExcelShapeWriter.MaxShapeNameLength + 1);

        ShapeWriteOutcome outcome = writer.Create(RectangleRequest(id: tooLong));

        Assert.Equal(ShapeWriteRefusal.HostRejected, outcome.Refusal);
        Assert.Empty(writer.Created);
    }

    /// <summary>
    /// The positive counterpart: an identifier exactly at the cap is still
    /// created, so the guard is a bound and not an accidental blanket refusal.
    /// </summary>
    [Fact]
    public void Create_accepts_an_identifier_exactly_at_the_host_name_limit()
    {
        var writer = new TestableWriter(ActiveApplication().Object, ClearGuard().Object);
        var atLimit = new string('r', ExcelShapeWriter.MaxShapeNameLength);

        ShapeWriteOutcome outcome = writer.Create(RectangleRequest(id: atLimit));

        Assert.True(outcome.Succeeded, outcome.Refusal?.ToString());
        Assert.Single(writer.Created);
    }

    /// <summary>
    /// A shape the host will not stamp must not be left behind. An untagged shape
    /// is invisible to <c>ListOwned</c>, so neither R4.7 nor R4.8 can ever find it
    /// again, and it would sit on the user's sheet permanently - the one outcome
    /// worse than failing to render.
    /// </summary>
    [Fact]
    public void Create_deletes_the_shape_it_created_when_the_ownership_tag_cannot_be_written()
    {
        var writer = new TestableWriter(ActiveApplication().Object, ClearGuard().Object)
        {
            FailOwnershipTagWrite = true,
        };

        ShapeWriteOutcome outcome = writer.Create(RectangleRequest());

        Assert.Equal(ShapeWriteRefusal.HostRejected, outcome.Refusal);
        Assert.True(writer.CreatedShapeDeleted, "The unstamped shape must be deleted, not orphaned.");
    }

    /// <summary>
    /// R4.8 D1: the ownership filter authorises the delete, not the name. The
    /// sentinel deliberately carries the requested name.
    /// </summary>
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

    /// <summary>
    /// A shape carrying ANOTHER identifier's valid tag is not this primitive's
    /// shape. <c>IsOwnedTag</c> alone accepts it - the prefix and hash are both
    /// well formed - so a format check would let a refresh move and rewrite a shape
    /// the caller never named. The ownership proof is the tag for <em>this</em>
    /// identifier, so all three entry points refuse the mismatch.
    /// </summary>
    [Fact]
    public void A_shape_carrying_another_identifier_s_valid_tag_is_never_treated_as_owned()
    {
        // The sentinel deliberately carries the REQUESTED name and a tag that is
        // perfectly valid - just for a different primitive.
        const string OtherId = "row-9:bar";
        var impostor = NewShape("row-1:bar", ShapeOwnershipTag.ForPrimitiveId(OtherId));
        var writer = new TestableWriter(ActiveApplication().Object, ClearGuard().Object, impostor);

        // Update must not move it.
        Assert.Equal(ShapeWriteRefusal.NotFound, writer.Update(RectangleRequest()).Refusal);
        Assert.Equal(0f, impostor.Left);

        // Delete must not remove it. The alternative text is unchanged, so a
        // delete that ran would have taken the shape with it.
        Assert.Equal(ShapeWriteRefusal.NotFound, writer.Delete("row-1:bar").Refusal);
        Assert.Equal(ShapeOwnershipTag.ForPrimitiveId(OtherId), impostor.AlternativeText);
        Assert.Equal("row-1:bar", impostor.Name);

        // And ListOwned must not report it as this add-in's shape.
        Assert.Empty(writer.ListOwned());
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

    private static Excel.Shape NewShape(
        string name,
        string? alternativeText,
        Action? onDelete = null,
        Action<string, MsoZOrderCmd>? onZOrder = null)
    {
        // Backing fields, not fixed Returns(name): a test that simulates a
        // user editing the alternative text has to be able to write it, or the
        // "update never re-stamps the tag" assertion would pass vacuously
        // against a shape whose tag could never change in the first place.
        string currentName = name;
        string currentAlternativeText = alternativeText ?? string.Empty;
        float currentLeft = 0f;
        float currentTop = 0f;
        var shape = new Mock<Excel.Shape>();
        _ = shape.SetupGet(s => s.Name).Returns(() => currentName);
        _ = shape.SetupSet(s => s.Name = It.IsAny<string>()).Callback<string>(value => currentName = value);
        _ = shape.SetupGet(s => s.AlternativeText).Returns(() => currentAlternativeText);
        _ = shape.SetupSet(s => s.AlternativeText = It.IsAny<string>())
            .Callback<string>(value => currentAlternativeText = value);
        // Geometry is recorded through backing fields so a test can assert that an
        // update actually MOVED the shape, rather than inferring it from a success
        // flag. A mock with no setup silently returns 0 for every property. The PIA
        // types Left/Top as Single, which is the host side of the port's own
        // double -> float rounding boundary.
        _ = shape.SetupGet(s => s.Left).Returns(() => currentLeft);
        _ = shape.SetupSet(s => s.Left = It.IsAny<float>()).Callback<float>(value => currentLeft = value);
        _ = shape.SetupGet(s => s.Top).Returns(() => currentTop);
        _ = shape.SetupSet(s => s.Top = It.IsAny<float>()).Callback<float>(value => currentTop = value);
        _ = shape.Setup(s => s.Delete()).Callback(() => onDelete?.Invoke());
        // A seeded shape needs its own z-order recorder: the writer's built-in
        // `ZOrderShapes` is wired only into the shapes `AddShape` creates, so a
        // z-order test built from seeded shapes would otherwise observe zero
        // commands and pass vacuously. The shape name is recorded alongside the
        // command so the reverse-order sequence is assertable from here too.
        _ = shape.Setup(s => s.ZOrder(It.IsAny<MsoZOrderCmd>()))
            .Callback<MsoZOrderCmd>(command => onZOrder?.Invoke(currentName, command));
        return shape.Object;
    }
}