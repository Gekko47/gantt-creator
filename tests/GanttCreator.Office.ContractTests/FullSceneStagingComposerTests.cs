using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R6.3: the full chart scene composes at the deterministic staging origin, reproduces
/// the R6.1 export bounds within the documented tolerance, and never enters the
/// ADR-0052 host-owned resize class.
/// </summary>
/// <remarks>
/// <para>
/// The fixture is a hand-built scene that closes its own frame on <c>ChartBounds</c>, so
/// the union of its primitives IS the export bounds - the entity guide section 1 rule
/// that no whitespace outside the bounds is exported. It deliberately contains both
/// families the panel path never carries: a plot-spanning band and a full-height
/// delineator (both host-owned by classification) and a milestone diamond.
/// </para>
/// <para>
/// The D5 scenario at the end runs the real <see cref="ShapeReconciler.Plan"/> against a
/// live-scheme writer while a staging-scheme composer's shapes exist on the same modelled
/// worksheet. Its remarks state the simulation's limit: staging writes no cells, so it
/// cannot itself schedule a refresh; the test asserts what happens when one fires.
/// </para>
/// </remarks>
public sealed class FullSceneStagingComposerTests
{
    private const double _chartLeft = 520;
    private const double _chartTop = 60;
    private const double _chartWidth = 320;
    private const double _chartHeight = 280;
    private const double _plotLeft = 620;
    private const double _plotTop = 100;
    private const double _plotWidth = 200;
    private const double _plotHeight = 200;

    private static readonly SceneStyle _fill = new("Fill", fillColour: ColourHex.Parse("#EEEEEE"));
    private static readonly SceneStyle _stroke = new(
        "Stroke",
        strokeColour: ColourHex.Parse("#777777"),
        outlineWidthPt: 0.5);
    private static readonly SceneStyle _text = new(
        "Text",
        fontSizePt: 9,
        textColour: ColourHex.Parse("#111111"));

    [Fact]
    public void The_full_composition_stages_every_scene_primitive_in_scene_order()
    {
        Fixture fixture = Build();
        FakeShapeWritePort port = new();

        StagingCompositionOutcome outcome = EditablePanelStagingComposer.ComposeFullScene(fixture.Outcome, port);

        Assert.True(outcome.Succeeded, "Composition refused: " + outcome.Refusal);

        string[] expected = [.. fixture.Scene.Primitives.Select(primitive => primitive.PrimitiveId)];
        Assert.NotEmpty(expected);

        // Completeness: the count matches, and so does the set.
        Assert.Equal(expected.Length, outcome.StagedPrimitiveIds.Count);
        Assert.Equal(expected, outcome.StagedPrimitiveIds);

        // Creation order IS the scene's back-to-front order: a new host shape lands in
        // front of the existing ones, so this is what makes the staged composition's
        // z-order the scene's z-order without a second ordering pass (R4.5).
        string[] created =
        [
            .. port.Calls
                .Where(call => call.StartsWith("Create:", StringComparison.Ordinal))
                .Select(call => call["Create:".Length..]),
        ];
        Assert.Equal(expected, created);
    }

    [Fact]
    public void The_staged_envelope_equals_the_export_bounds_within_the_documented_tolerance()
    {
        Fixture fixture = Build();
        FakeShapeWritePort port = new();

        StagingCompositionOutcome outcome = EditablePanelStagingComposer.ComposeFullScene(fixture.Outcome, port);

        Assert.True(outcome.Succeeded, "Composition refused: " + outcome.Refusal);
        Assert.True(outcome.StagedBounds is not null);

        // D2: the composed set's overall bounds equal the R6.1 export bounds within
        // GeometryMath.Epsilon. The fixture's frame closes on ChartBounds, so the union
        // must reproduce the translated bounds exactly.
        RectD[] envelopes =
        [
            .. fixture.Scene.Primitives.Select(primitive => Envelope(port.Requests[primitive.PrimitiveId])),
        ];
        StagedEnvelopeResult check = StagedEnvelopeCheck.Evaluate(envelopes, outcome.StagedBounds!.Value);

        Assert.Empty(check.Strays);
        Assert.True(
            check.Matches(outcome.StagedBounds.Value),
            $"Union {check.Union} does not equal {outcome.StagedBounds}.");

        PointD origin = EditablePanelStagingComposer.StagingOriginFor(fixture.ChartBounds);
        Assert.Equal(origin.X, check.Union!.Value.X, 6);
        Assert.Equal(origin.Y, check.Union.Value.Y, 6);
    }

    [Fact]
    public void Staged_shapes_leave_the_host_owned_resize_class()
    {
        // D1: staged shapes are temporary and deleted, never reconciled, so the host
        // must not stretch them when a user changes a row height or column width while
        // the set exists. The fixture contains two primitives the classification WOULD
        // call host-owned, so the re-stamp is not vacuous.
        Fixture fixture = Build();
        Assert.True(
            HostOwnedGeometry.IsHostOwned(fixture.Band, fixture.Scene.PlotBounds),
            "The fixture band must be host-owned before staging, or this test proves nothing.");
        Assert.True(
            HostOwnedGeometry.IsHostOwned(fixture.Delineator, fixture.Scene.PlotBounds),
            "The fixture delineator must be host-owned before staging, or this test proves nothing.");

        FakeShapeWritePort port = new();
        StagingCompositionOutcome outcome = EditablePanelStagingComposer.ComposeFullScene(fixture.Outcome, port);
        Assert.True(outcome.Succeeded, "Composition refused: " + outcome.Refusal);

        Assert.All(
            port.Requests.Values,
            request => Assert.False(
                request.HostOwnsExtent,
                $"Staged shape '{request.PrimitiveId}' entered the host-owned resize class."));
    }

    [Fact]
    public void A_chart_only_composition_still_stages_every_chart_primitive()
    {
        // The full-scene path has no zero-column early return: with no panel primitive
        // the chart still stages beside the chart-only bounds R6.1 resolved.
        Fixture fixture = Build(includePanel: false);
        FakeShapeWritePort port = new();

        StagingCompositionOutcome outcome = EditablePanelStagingComposer.ComposeFullScene(fixture.Outcome, port);

        Assert.True(outcome.Succeeded, "Composition refused: " + outcome.Refusal);
        Assert.NotEmpty(outcome.StagedPrimitiveIds);
        Assert.Equal(
            fixture.Scene.Primitives.Select(primitive => primitive.PrimitiveId),
            outcome.StagedPrimitiveIds);
        Assert.True(outcome.StagedBounds is not null);
    }

    [Fact]
    public void A_group_primitive_is_refused_before_any_create()
    {
        // A family the renderer cannot translate is surfaced as a typed refusal rather
        // than a silently partial composition; the sheet is untouched.
        Fixture fixture = Build(includeGroup: true);
        FakeShapeWritePort port = new();

        StagingCompositionOutcome outcome = EditablePanelStagingComposer.ComposeFullScene(fixture.Outcome, port);

        Assert.Equal(StagingCompositionRefusal.UnsupportedPrimitive, outcome.Refusal);
        Assert.DoesNotContain(
            port.Calls,
            call => call.StartsWith("Create:", StringComparison.Ordinal));
    }

    [Fact]
    public void An_existing_staging_collection_refuses_and_stages_nothing()
    {
        Fixture fixture = Build();
        FakeShapeWritePort port = new();
        _ = port.Owned(ShapeStagingTag.NamePrefix + "chart:bar:0");

        StagingCompositionOutcome outcome = EditablePanelStagingComposer.ComposeFullScene(fixture.Outcome, port);

        Assert.Equal(StagingCompositionRefusal.StagingAlreadyPopulated, outcome.Refusal);
        Assert.DoesNotContain(
            port.Calls,
            call => call.StartsWith("Create:", StringComparison.Ordinal));
    }

    [Fact]
    public void A_refused_resolution_is_surfaced_and_stages_nothing()
    {
        ExportBoundsOutcome refused = new(null, ExportBoundsRefusal.EmptyVisibleTable, null);
        FakeShapeWritePort port = new();

        StagingCompositionOutcome outcome = EditablePanelStagingComposer.ComposeFullScene(refused, port);

        Assert.Equal(StagingCompositionRefusal.ExportRefused, outcome.Refusal);
        Assert.Same(refused, outcome.Resolution);
        Assert.Empty(port.Calls);
    }

    [Fact]
    public void A_successful_resolution_without_a_scene_is_refused()
    {
        Fixture fixture = Build();
        ExportBoundsOutcome sceneLess = new(fixture.Outcome.Bounds, null, null);
        FakeShapeWritePort port = new();

        StagingCompositionOutcome outcome = EditablePanelStagingComposer.ComposeFullScene(sceneLess, port);

        Assert.Equal(StagingCompositionRefusal.MissingScene, outcome.Refusal);
        Assert.Empty(port.Calls);
    }

    [Fact]
    public void Null_inputs_are_refused()
    {
        Fixture fixture = Build();
        FakeShapeWritePort port = new();

        Assert.Equal(
            StagingCompositionRefusal.NullResolution,
            EditablePanelStagingComposer.ComposeFullScene(null, port).Refusal);
        Assert.Equal(
            StagingCompositionRefusal.NullPort,
            EditablePanelStagingComposer.ComposeFullScene(fixture.Outcome, null).Refusal);
    }

    [Fact]
    public void A_host_refusal_after_two_creates_leaves_no_staged_shape()
    {
        Fixture fixture = Build();
        FakeShapeWritePort port = new();
        port.Refusals["Create#3"] = ShapeWriteRefusal.HostRejected;

        StagingCompositionOutcome outcome = EditablePanelStagingComposer.ComposeFullScene(fixture.Outcome, port);

        Assert.Equal(StagingCompositionRefusal.HostRejected, outcome.Refusal);
        Assert.Equal(ShapeWriteRefusal.HostRejected, outcome.WriteRefusal);
        Assert.Empty(outcome.StagedPrimitiveIds);
        Assert.Empty(port.AllShapeNames);
        Assert.Equal(
            2,
            port.Calls.Count(call => call.StartsWith("Delete:", StringComparison.Ordinal)));
    }

    [Fact]
    public void The_composition_issues_only_shape_port_calls_and_one_batch()
    {
        Fixture fixture = Build();
        FakeShapeWritePort port = new();

        StagingCompositionOutcome outcome = EditablePanelStagingComposer.ComposeFullScene(fixture.Outcome, port);
        Assert.True(outcome.Succeeded, "Composition refused: " + outcome.Refusal);

        Assert.All(
            port.Calls,
            call => Assert.True(
                call == "ListOwned"
                    || call.StartsWith("Create:", StringComparison.Ordinal)
                    || call.StartsWith("Delete:", StringComparison.Ordinal),
                "A non-shape port call was issued: " + call));
        Assert.Equal(1, port.BatchesStarted);
        Assert.Equal(1, port.BatchesDisposed);
    }

    [Fact]
    public void Repeated_composition_after_cleanup_is_deterministic()
    {
        Fixture fixture = Build();
        FakeShapeWritePort port = new();

        StagingCompositionOutcome first = EditablePanelStagingComposer.ComposeFullScene(fixture.Outcome, port);
        Assert.True(first.Succeeded, "First composition refused: " + first.Refusal);
        foreach (string id in first.StagedPrimitiveIds)
        {
            Assert.True(port.Delete(id).Succeeded);
        }

        StagingCompositionOutcome second = EditablePanelStagingComposer.ComposeFullScene(fixture.Outcome, port);

        Assert.True(second.Succeeded, "Second composition refused: " + second.Refusal);
        Assert.Equal(first.StagedPrimitiveIds, second.StagedPrimitiveIds);
        Assert.Equal(first.StagedBounds, second.StagedBounds);
    }

    [Fact]
    public void A_coalesced_refresh_during_staging_reconciles_the_live_chart_and_leaves_staged_shapes_alone()
    {
        // D5 positive: a refresh that fires while a composition is staged reconciles the
        // live chart and is structurally blind to the staged set. The compositing
        // operation here writes no cells, so it cannot itself schedule a refresh - this
        // simulates the concurrent user edit that does, by running the real reconciler
        // plan against the live scheme while both families occupy one modelled sheet.
        Fixture fixture = Build();
        SharedShapes sheet = new();
        SchemeWriter live = new(sheet, ShapeOwnershipScheme.Live);
        SchemeWriter staging = new(sheet, ShapeOwnershipScheme.Staging);
        Assert.True(live.Create(StagingSchemeFixtures.RectangleRequest("live:bar-1")).Succeeded);
        Assert.True(live.Create(StagingSchemeFixtures.RectangleRequest("live:bar-2")).Succeeded);

        StagingCompositionOutcome staged = EditablePanelStagingComposer.ComposeFullScene(fixture.Outcome, staging);
        Assert.True(staged.Succeeded, "Composition refused: " + staged.Refusal);
        string[] stagedNames =
        [
            .. staged.StagedPrimitiveIds.Select(id => ShapeStagingTag.NamePrefix + id),
        ];
        int shapesWhileStaged = sheet.Shapes.Count;
        Assert.Equal(2 + stagedNames.Length, shapesWhileStaged);

        // The coalesced pass: the reconciler plans against the LIVE owned set only, then
        // executes through the live writer. bar-2 disappears, bar-3 appears, bar-1 stays.
        ShapeReconcilePlan plan = ShapeReconciler.Plan(
            [StagingSchemeFixtures.RectangleRequest("live:bar-1"), StagingSchemeFixtures.RectangleRequest("live:bar-3")],
            live.ListOwned());
        foreach (ShapeReconcileOperation operation in plan.Operations)
        {
            ShapeWriteOutcome write = operation.Action switch
            {
                ShapeReconcileAction.Update => live.Update(operation.Request!),
                ShapeReconcileAction.Create => live.Create(operation.Request!),
                ShapeReconcileAction.Delete => live.Delete(operation.PrimitiveId),
                _ => throw new InvalidOperationException(
                    "An unhandled reconcile action: " + operation.Action),
            };
            Assert.True(write.Succeeded, $"Reconcile operation refused: {operation.Action} {operation.PrimitiveId} {write.Refusal}");
        }

        // The live chart reconciled.
        Assert.Equal(["live:bar-1", "live:bar-3"], live.ListOwned());

        // Every staged shape survived, and the staged set was not counted as live work.
        Assert.Equal(shapesWhileStaged, sheet.Shapes.Count);
        foreach (string name in stagedNames)
        {
            Assert.Contains(sheet.Shapes, shape => string.Equals(shape.Name, name, StringComparison.Ordinal));
        }
    }

    private static RectD Envelope(OfficeShapeRequest request) =>
        request.Geometry.Bounds
        ?? new RectD(
            Math.Min(request.Geometry.From!.Value.X, request.Geometry.To!.Value.X),
            Math.Min(request.Geometry.From.Value.Y, request.Geometry.To.Value.Y),
            Math.Abs(request.Geometry.To.Value.X - request.Geometry.From.Value.X),
            Math.Abs(request.Geometry.To.Value.Y - request.Geometry.From.Value.Y));

    private sealed record Fixture(
        GanttScene Scene,
        ExportBoundsOutcome Outcome,
        RectD ChartBounds,
        SceneRect Band,
        SceneLine Delineator);

    private static Fixture Build(bool includePanel = true, bool includeGroup = false)
    {
        RectD chart = new(_chartLeft, _chartTop, _chartWidth, _chartHeight);
        RectD plot = new(_plotLeft, _plotTop, _plotWidth, _plotHeight);
        SceneOwnerId rowOne = SceneOwnerId.ForRow(GanttRowId.New());
        SceneOwnerId rowTwo = SceneOwnerId.ForRow(GanttRowId.New());

        SceneRect bar = new(
            "row-1:bar",
            rowOne,
            ZLayer.ActivityBody,
            new RectD(640, 140, 80, 10),
            _fill);
        ScenePolygon milestone = new(
            "row-2:milestone",
            rowTwo,
            ZLayer.ActivityBody,
            [new PointD(770, 140), new PointD(780, 150), new PointD(770, 160), new PointD(760, 150)],
            _fill);
        SceneRect band = new("chart:background", SceneOwnerId.Chart, ZLayer.Background, plot, _fill);
        SceneLine delineator = new(
            "row-3:delineator",
            SceneOwnerId.Chart,
            ZLayer.Grid,
            new PointD(700, _plotTop),
            new PointD(700, _plotTop + _plotHeight),
            _stroke);

        List<ScenePrimitive> primitives = [bar, milestone, band, delineator];
        if (includePanel)
        {
            primitives.Add(new SceneRect(
                "chart:panel-cell:0",
                SceneOwnerId.Chart,
                ZLayer.Background,
                new RectD(540, 120, 60, 14),
                _fill));
            primitives.Add(new SceneText(
                "chart:panel-text:0",
                SceneOwnerId.Chart,
                ZLayer.Label,
                "Type",
                new RectD(542, 121, 40, 12),
                _text,
                GanttTextAlignment.Left));
        }

        // The frame closes the outline on ChartBounds, which is what makes the union of
        // the primitives equal the export bounds (entity guide section 1).
        primitives.Add(new SceneLine(
            "chart:frame:left", SceneOwnerId.Chart, ZLayer.Frame,
            new PointD(_chartLeft, _chartTop), new PointD(_chartLeft, _chartTop + _chartHeight), _stroke));
        primitives.Add(new SceneLine(
            "chart:frame:right", SceneOwnerId.Chart, ZLayer.Frame,
            new PointD(_chartLeft + _chartWidth, _chartTop),
            new PointD(_chartLeft + _chartWidth, _chartTop + _chartHeight), _stroke));
        primitives.Add(new SceneLine(
            "chart:frame:top", SceneOwnerId.Chart, ZLayer.Frame,
            new PointD(_chartLeft, _chartTop), new PointD(_chartLeft + _chartWidth, _chartTop), _stroke));
        primitives.Add(new SceneLine(
            "chart:frame:bottom", SceneOwnerId.Chart, ZLayer.Frame,
            new PointD(_chartLeft, _chartTop + _chartHeight),
            new PointD(_chartLeft + _chartWidth, _chartTop + _chartHeight), _stroke));

        if (includeGroup)
        {
            primitives.Add(new SceneGroup(
                "chart:group:0",
                SceneOwnerId.Chart,
                ZLayer.Frame,
                ["row-1:bar", "row-2:milestone"]));
        }

        SceneCreationOutcome created = GanttScene.TryCreate(chart, plot, primitives, []);
        Assert.NotNull(created.Scene);
        GanttScene scene = created.Scene;
        ExportBoundsOutcome outcome = new(new ExportBounds(scene.ChartBounds), null, null, scene);
        return new Fixture(scene, outcome, scene.ChartBounds, band, delineator);
    }
}
