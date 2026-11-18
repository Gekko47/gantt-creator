using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R6.2: the editable-panel staging composition reproduces the scene's R3.11 panel
/// primitives at the deterministic staging origin.
/// </summary>
/// <remarks>
/// The fixture builds its panel with the real <see cref="PanelBuilder"/> over a measured
/// grid, then wraps those primitives in a scene beside two non-panel chart primitives, so
/// the count, bounds, text, style, and order pins are asserted against genuine builder
/// output rather than hand-written stand-ins. The recording port is
/// <see cref="FakeShapeWritePort"/>, which records every call and can refuse the Nth
/// create for the cleanup positive.
/// </remarks>
public sealed class EditablePanelStagingComposerTests
{
    private const double _headerHeight = 12;
    private const double _rowHeight = 14;
    private const double _plotLeft = 600;
    private const double _plotTop = 80;
    private const double _plotWidth = 300;
    private const double _plotHeight = 240;
    private const double _yearBandHeight = 20;

    private static readonly PanelTheme _theme = new(
        new SceneStyle("DataPanelFill", fillColour: ColourHex.Parse("#EEEEEE")),
        new SceneStyle("DefaultText", fontSizePt: 8.5, textColour: ColourHex.Parse("#111111")),
        new SceneStyle("HeaderFill", fillColour: ColourHex.Parse("#DDDDDD")),
        new SceneStyle("HeaderText", fontSizePt: 9.0, bold: true, textColour: ColourHex.Parse("#222222")),
        new SceneStyle("Border", strokeColour: ColourHex.Parse("#777777"), outlineWidthPt: 0.5));

    [Fact]
    public void The_composition_stages_exactly_the_scene_s_panel_primitives_in_scene_order()
    {
        Fixture fixture = Build();
        FakeShapeWritePort port = new();

        StagingCompositionOutcome outcome = EditablePanelStagingComposer.Compose(fixture.Outcome, port);

        Assert.True(outcome.Succeeded, "Composition refused: " + outcome.Refusal);

        string[] expected =
        [
            .. fixture.Scene.Primitives
                .Where(PanelPrimitiveRoles.IsPanelPrimitive)
                .Select(primitive => primitive.PrimitiveId),
        ];
        Assert.NotEmpty(expected);
        Assert.Equal(expected, outcome.StagedPrimitiveIds);

        // Creation order IS the scene's back-to-front order: a new host shape lands in
        // front of the existing ones, so this is what makes the staged panel's z-order
        // the scene's z-order without a second ordering pass.
        string[] created =
        [
            .. port.Calls
                .Where(call => call.StartsWith("Create:", StringComparison.Ordinal))
                .Select(call => call["Create:".Length..]),
        ];
        Assert.Equal(expected, created);

        // Non-panel primitives never reach the port.
        Assert.DoesNotContain("chart:frame:left", port.Requests.Keys);
        Assert.DoesNotContain("chart:background", port.Requests.Keys);
    }

    [Fact]
    public void Every_staged_shape_lands_inside_the_staging_envelope_at_its_translated_bounds()
    {
        Fixture fixture = Build();
        FakeShapeWritePort port = new();

        StagingCompositionOutcome outcome = EditablePanelStagingComposer.Compose(fixture.Outcome, port);
        Assert.True(outcome.Succeeded, "Composition refused: " + outcome.Refusal);
        Assert.True(outcome.StagedBounds is not null);

        PointD origin = EditablePanelStagingComposer.StagingOriginFor(fixture.ChartBounds);
        double dx = origin.X - fixture.ChartBounds.X;
        double dy = origin.Y - fixture.ChartBounds.Y;
        RectD envelope = outcome.StagedBounds!.Value;

        foreach (ScenePrimitive primitive in fixture.Panel.Primitives)
        {
            OfficeShapeRequest request = port.Requests[primitive.PrimitiveId];

            switch (primitive)
            {
                case SceneRect rect:
                    AssertRect(rect.Bounds, dx, dy, request.Geometry.Bounds!.Value);
                    break;
                case SceneText text:
                    AssertRect(text.TextBounds, dx, dy, request.Geometry.Bounds!.Value);
                    break;
                case SceneLine line:
                    Assert.Equal(line.From.X + dx, request.Geometry.From!.Value.X, 6);
                    Assert.Equal(line.From.Y + dy, request.Geometry.From.Value.Y, 6);
                    Assert.Equal(line.To.X + dx, request.Geometry.To!.Value.X, 6);
                    Assert.Equal(line.To.Y + dy, request.Geometry.To.Value.Y, 6);
                    break;
                default:
                    Assert.Fail("Unexpected panel primitive family: " + primitive.GetType().Name);
                    break;
            }

            // "Nothing outside the staging origin": every staged shape lies inside the
            // envelope the outcome reports, so the live inspection has one rectangle to
            // assert against.
            RectD bounds = request.Geometry.Bounds
                ?? new RectD(
                    Math.Min(request.Geometry.From!.Value.X, request.Geometry.To!.Value.X),
                    Math.Min(request.Geometry.From.Value.Y, request.Geometry.To.Value.Y),
                    Math.Abs(request.Geometry.To.Value.X - request.Geometry.From.Value.X),
                    Math.Abs(request.Geometry.To.Value.Y - request.Geometry.From.Value.Y));
            Assert.True(
                bounds.X >= envelope.X - 0.01
                && bounds.Y >= envelope.Y - 0.01
                && bounds.Right <= envelope.Right + 0.01
                && bounds.Bottom <= envelope.Bottom + 0.01,
                $"Staged shape escapes the envelope: {primitive.PrimitiveId} {bounds} vs {envelope}");
        }
    }

    [Fact]
    public void Cell_and_header_texts_are_verbatim_and_a_blank_cell_emits_no_text()
    {
        Fixture fixture = Build();
        FakeShapeWritePort port = new();
        StagingCompositionOutcome outcome = EditablePanelStagingComposer.Compose(fixture.Outcome, port);
        Assert.True(outcome.Succeeded, "Composition refused: " + outcome.Refusal);

        // Headers carry the schema display names verbatim; Description is the one
        // left-aligned header (section 4).
        foreach (PanelColumn column in fixture.Grid.Columns)
        {
            string id = ScenePrimitive.CreateId(SceneOwnerId.Chart, "header-text:" + column.Name);
            Assert.Equal(column.Name, port.Requests[id].Text);
            Assert.Equal(
                string.Equals(column.Name, "Description", StringComparison.Ordinal)
                    ? GanttTextAlignment.Left
                    : GanttTextAlignment.Centre,
                port.Requests[id].Alignment);
        }

        // The blank Description cell of the first row emits no text shape, exactly as
        // the builder omits it; the second row's text is reproduced verbatim.
        string blankId = ScenePrimitive.CreateId(
            SceneOwnerId.ForRow(fixture.Rows[0].RowId),
            "panel-text:Description");
        string presentId = ScenePrimitive.CreateId(
            SceneOwnerId.ForRow(fixture.Rows[1].RowId),
            "panel-text:Description");
        Assert.DoesNotContain(blankId, port.Requests.Keys);
        Assert.Equal("Build phase", port.Requests[presentId].Text);
    }

    [Fact]
    public void Styles_are_the_panel_theme_tokens_and_the_scene_layers()
    {
        Fixture fixture = Build();
        FakeShapeWritePort port = new();
        StagingCompositionOutcome outcome = EditablePanelStagingComposer.Compose(fixture.Outcome, port);
        Assert.True(outcome.Succeeded, "Composition refused: " + outcome.Refusal);

        // The style values are copied from the primitives, which carry the theme's own
        // tokens (R4.6) - the composer re-derives nothing.
        SceneRect bodyCell = fixture.Panel.Primitives
            .OfType<SceneRect>()
            .First(rect => rect.PrimitiveId.Contains("panel-cell:", StringComparison.Ordinal));
        OfficeShapeRequest body = port.Requests[bodyCell.PrimitiveId];
        Assert.Equal(_theme.BodyFill.FillColour, body.FillColour);

        SceneText bodyText = fixture.Panel.Primitives
            .OfType<SceneText>()
            .First(text => text.PrimitiveId.Contains("panel-text:", StringComparison.Ordinal));
        OfficeShapeRequest text = port.Requests[bodyText.PrimitiveId];
        Assert.Equal(_theme.BodyText.FontSizePt, text.FontSizePt);
        Assert.Equal(_theme.BodyText.TextColour, text.TextColour);

        SceneText header = fixture.Panel.Primitives
            .OfType<SceneText>()
            .First(candidate => candidate.PrimitiveId.Contains("header-text:", StringComparison.Ordinal));
        Assert.Equal(_theme.HeaderText.Bold, port.Requests[header.PrimitiveId].Bold);

        SceneLine border = fixture.Panel.Primitives
            .OfType<SceneLine>()
            .First(line => line.PrimitiveId.Contains("panel-border-v:", StringComparison.Ordinal));
        Assert.Equal(_theme.Border.StrokeColour, port.Requests[border.PrimitiveId].StrokeColour);
        Assert.Equal(_theme.Border.OutlineWidthPt, port.Requests[border.PrimitiveId].LineWidthPt);

        // Z-order contract: the scene's layers cross the port unchanged - backgrounds at
        // the back, text above its own cell, borders framing both.
        Assert.All(
            fixture.Panel.Primitives.OfType<SceneRect>()
                .Where(rect => rect.PrimitiveId.Contains("panel-cell:", StringComparison.Ordinal)),
            rect => Assert.Equal(ZLayer.Background, port.Requests[rect.PrimitiveId].ZLayer));
        Assert.All(
            fixture.Panel.Primitives.OfType<SceneText>()
                .Where(candidate => candidate.PrimitiveId.Contains("panel-text:", StringComparison.Ordinal)),
            candidate => Assert.Equal(ZLayer.Label, port.Requests[candidate.PrimitiveId].ZLayer));
        Assert.All(
            fixture.Panel.Primitives.OfType<SceneLine>(),
            line => Assert.Equal(ZLayer.Frame, port.Requests[line.PrimitiveId].ZLayer));
    }

    [Fact]
    public void Composing_no_panel_is_a_supported_empty_composition()
    {
        // D5 / ADR-0065 D4: zero visible columns means the scene carries no panel
        // primitive, and the composition is the chart alone. Nothing is staged, nothing
        // is refused, and the port is not touched at all - not even the pre-flight.
        Fixture fixture = Build(zeroColumns: true);
        FakeShapeWritePort port = new();

        StagingCompositionOutcome outcome = EditablePanelStagingComposer.Compose(fixture.Outcome, port);

        Assert.True(outcome.Succeeded, "Composition refused: " + outcome.Refusal);
        Assert.Empty(outcome.StagedPrimitiveIds);
        Assert.Null(outcome.StagedBounds);
        Assert.Empty(port.Calls);
    }

    [Fact]
    public void An_existing_staging_collection_refuses_and_stages_nothing()
    {
        Fixture fixture = Build();
        FakeShapeWritePort port = new();
        _ = port.Owned(ShapeStagingTag.NamePrefix + "chart:panel-border-v:0");

        StagingCompositionOutcome outcome = EditablePanelStagingComposer.Compose(fixture.Outcome, port);

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

        StagingCompositionOutcome outcome = EditablePanelStagingComposer.Compose(refused, port);

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

        StagingCompositionOutcome outcome = EditablePanelStagingComposer.Compose(sceneLess, port);

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
            EditablePanelStagingComposer.Compose(null, port).Refusal);
        Assert.Equal(
            StagingCompositionRefusal.NullPort,
            EditablePanelStagingComposer.Compose(fixture.Outcome, null).Refusal);
    }

    [Fact]
    public void A_host_refusal_after_two_creates_leaves_no_staged_shape()
    {
        // D3's cleanup rule: a refusal mid-composition deletes what was staged before
        // the failure, so a failed composition leaves the sheet as it found it.
        Fixture fixture = Build();
        FakeShapeWritePort port = new();
        port.Refusals["Create#3"] = ShapeWriteRefusal.HostRejected;

        StagingCompositionOutcome outcome = EditablePanelStagingComposer.Compose(fixture.Outcome, port);

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
        // The port's surface is shape-only by construction, so a cell write is
        // unrepresentable here; the recorded call log is the observation point, and the
        // integration body asserts the worksheet-level facts (sheet count and used
        // range) on the host. The batch is opened once and disposed on the success path.
        Fixture fixture = Build();
        FakeShapeWritePort port = new();

        StagingCompositionOutcome outcome = EditablePanelStagingComposer.Compose(fixture.Outcome, port);
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

        StagingCompositionOutcome first = EditablePanelStagingComposer.Compose(fixture.Outcome, port);
        Assert.True(first.Succeeded, "First composition refused: " + first.Refusal);
        foreach (string id in first.StagedPrimitiveIds)
        {
            Assert.True(port.Delete(id).Succeeded);
        }

        StagingCompositionOutcome second = EditablePanelStagingComposer.Compose(fixture.Outcome, port);

        Assert.True(second.Succeeded, "Second composition refused: " + second.Refusal);
        Assert.Equal(first.StagedPrimitiveIds, second.StagedPrimitiveIds);
        Assert.Equal(first.StagedBounds, second.StagedBounds);
    }

    private static void AssertRect(RectD expected, double dx, double dy, RectD actual)
    {
        Assert.Equal(expected.X + dx, actual.X, 6);
        Assert.Equal(expected.Y + dy, actual.Y, 6);
        Assert.Equal(expected.Width, actual.Width, 6);
        Assert.Equal(expected.Height, actual.Height, 6);
    }

    private sealed record Fixture(
        GanttScene Scene,
        ExportBoundsOutcome Outcome,
        PanelBuildResult Panel,
        PanelCellGrid Grid,
        IReadOnlyList<PanelRow> Rows,
        RectD ChartBounds);

    private static Fixture Build(bool zeroColumns = false)
    {
        PanelColumn[] columns = zeroColumns
            ? []
            : [.. GanttTableSchema.Default.Columns.Select((column, index) => new PanelColumn(column.Name, 20.0 + index))];

        string[] required = zeroColumns
            ? []
            : [.. GanttTableSchema.Default.Columns.Where(column => column.IsRequired).Select(column => column.Name)];

        PanelCellGridCreationOutcome grid = PanelCellGrid.TryCreate(
            columns,
            [_rowHeight, _rowHeight],
            _headerHeight,
            required);
        Assert.True(grid.Succeeded, "Panel grid refused: " + grid.Refusal);
        PanelCellGrid panelGrid = grid.Grid!;

        string?[] firstCells = [.. Enumerable.Repeat<string?>("x", GanttTableSchema.Default.Columns.Count)];
        firstCells[SchemaIndex("Description")] = null;
        string?[] secondCells = [.. firstCells];
        secondCells[SchemaIndex("Description")] = "Build phase";
        secondCells[SchemaIndex("Start")] = GanttDateFormatting.FormatDdMMyyyy(new DateOnly(2026, 1, 5));
        secondCells[SchemaIndex("Finish")] = GanttDateFormatting.FormatDdMMyyyy(new DateOnly(2026, 1, 9));

        PanelRow[] rows =
        [
            new(GanttRowId.New(), firstCells),
            new(GanttRowId.New(), secondCells),
        ];

        RectD plotBounds = new(_plotLeft, _plotTop, _plotWidth, _plotHeight);
        PanelBuildResult panelResult;
        if (zeroColumns)
        {
            // ADR-0065 D4: with no visible column the resolver supplies no panel, so the
            // builder never runs and the scene carries no panel primitive at all. The
            // empty result stands in for that absent panel.
            RectD absent = new(_plotLeft, _plotTop - _yearBandHeight - _headerHeight, 0, _headerHeight);
            panelResult = new PanelBuildResult([], absent, absent);
        }
        else
        {
            PanelBuildOutcome panel = PanelBuilder.TryBuild(
                new PanelBuildRequest(panelGrid, rows, plotBounds, _plotTop - _yearBandHeight, _theme));
            Assert.True(panel.Succeeded, "Panel build refused: " + panel.Refusal);
            panelResult = panel.Result!;
        }

        // Two non-panel primitives, interleaved around the panel, so filtering is proved
        // to be a role test rather than a position-of-first-panel assumption.
        SceneStyle frameStyle = new("MajorBoundaryPt", strokeColour: ColourHex.Parse("#000000"), outlineWidthPt: 0.25);
        ScenePrimitive[] chartPrimitives =
        [
            new SceneRect(
                ScenePrimitive.CreateId(SceneOwnerId.Chart, "chart:background"),
                SceneOwnerId.Chart,
                ZLayer.Background,
                new RectD(plotBounds.X - 10, plotBounds.Y - 10, 10, 10),
                new SceneStyle("ChartBackground")),
            new SceneLine(
                ScenePrimitive.CreateId(SceneOwnerId.Chart, "chart:frame:left"),
                SceneOwnerId.Chart,
                ZLayer.Frame,
                new PointD(plotBounds.X - 10, plotBounds.Y - 10),
                new PointD(plotBounds.X - 10, plotBounds.Bottom),
                frameStyle),
        ];

        List<ScenePrimitive> all = [.. chartPrimitives, .. panelResult.Primitives];
        RectD chartBounds = new(
            panelResult.PanelBounds.X,
            panelResult.HeaderRowBounds.Y - 10,
            plotBounds.Right + 10 - panelResult.PanelBounds.X,
            plotBounds.Bottom + 10 - (panelResult.HeaderRowBounds.Y - 10));

        SceneCreationOutcome scene = GanttScene.TryCreate(chartBounds, plotBounds, all, []);
        Assert.NotNull(scene.Scene);

        ExportBoundsOutcome outcome = new(new ExportBounds(chartBounds), null, null, scene.Scene);
        return new Fixture(scene.Scene!, outcome, panelResult, panelGrid, rows, chartBounds);
    }

    private static int SchemaIndex(string name)
    {
        for (int index = 0; index < GanttTableSchema.Default.Columns.Count; index++)
        {
            if (string.Equals(GanttTableSchema.Default.Columns[index].Name, name, StringComparison.Ordinal))
            {
                return index;
            }
        }

        throw new InvalidOperationException($"'{name}' is not a schema column.");
    }
}
