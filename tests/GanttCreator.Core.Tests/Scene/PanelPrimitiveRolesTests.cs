using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// R6.2: the panel-primitive predicate is the single authority the editable-export
/// composition uses to pick the data panel out of a mixed export scene.
/// </summary>
/// <remarks>
/// The coverage fact is the load-bearing one: it fails if <see cref="PanelBuilder"/>
/// ever emits a role the predicate does not know, which is exactly how a panel family
/// would silently vanish from an export. The negative facts prove the predicate is a
/// role test, not a substring search.
/// </remarks>
public sealed class PanelPrimitiveRolesTests
{
    private const double PlotLeft = 800.0;
    private const double PlotTop = 70.0;
    private const double PlotWidth = 400.0;
    private const double PlotHeight = 280.0;
    private const double YearBandHeight = 20.0;
    private const double HeaderHeight = 14.0;
    private const double RowHeight = 12.0;

    private static readonly PanelTheme _theme = new(
        new SceneStyle("DataPanelFill"),
        new SceneStyle("DefaultText"),
        new SceneStyle("HeaderFill"),
        new SceneStyle("DefaultText", fontSizePt: 9.0, bold: true),
        new SceneStyle("GridLine", strokeColour: ColourHex.Parse("#BFBFBF"), outlineWidthPt: 0.5));

    [Fact]
    public void Every_primitive_the_panel_builder_emits_is_a_panel_primitive()
    {
        PanelBuildResult result = Build();

        Assert.NotEmpty(result.Primitives);
        Assert.All(
            result.Primitives,
            primitive => Assert.True(
                PanelPrimitiveRoles.IsPanelPrimitive(primitive),
                $"The panel builder emitted a role the predicate does not classify: {primitive.PrimitiveId}"));

        // Every family is represented, so a predicate that matched only one prefix
        // could not pass this fact while the others silently vanished.
        string[] ids = [.. result.Primitives.Select(primitive => primitive.PrimitiveId)];
        Assert.Contains(ids, id => id.Contains("header-cell:", StringComparison.Ordinal));
        Assert.Contains(ids, id => id.Contains("header-text:", StringComparison.Ordinal));
        Assert.Contains(ids, id => id.Contains("panel-cell:", StringComparison.Ordinal));
        Assert.Contains(ids, id => id.Contains("panel-text:", StringComparison.Ordinal));
        Assert.Contains(ids, id => id.Contains("panel-border-h:", StringComparison.Ordinal));
        Assert.Contains(ids, id => id.Contains("panel-border-v:", StringComparison.Ordinal));
    }

    [Fact]
    public void A_chart_frame_primitive_is_not_a_panel_primitive()
    {
        ScenePrimitive frame = new SceneLine(
            ScenePrimitive.CreateId(SceneOwnerId.Chart, "chart:frame:left"),
            SceneOwnerId.Chart,
            ZLayer.Frame,
            new PointD(0, 0),
            new PointD(0, 100),
            new SceneStyle("MajorBoundaryPt"));

        Assert.False(PanelPrimitiveRoles.IsPanelPrimitive(frame));
    }

    [Fact]
    public void A_row_entity_primitive_is_not_a_panel_primitive()
    {
        SceneOwnerId owner = SceneOwnerId.ForRow(GanttRowId.New());
        ScenePrimitive bar = new SceneRect(
            ScenePrimitive.CreateId(owner, "activity-bar"),
            owner,
            ZLayer.Background,
            new RectD(0, 0, 10, 10),
            new SceneStyle("AsPlannedActivity"));

        Assert.False(PanelPrimitiveRoles.IsPanelPrimitive(bar));
    }

    [Fact]
    public void A_role_that_merely_contains_a_prefix_is_not_classified()
    {
        // The role is anchored at the start: a role that mentions a panel prefix in
        // the middle is not a panel primitive.
        ScenePrimitive embedded = new SceneRect(
            ScenePrimitive.CreateId(SceneOwnerId.Chart, "custom:panel-cell:Type"),
            SceneOwnerId.Chart,
            ZLayer.Background,
            new RectD(0, 0, 10, 10),
            new SceneStyle("DataPanelFill"));

        Assert.False(PanelPrimitiveRoles.IsPanelPrimitive(embedded));
    }

    [Fact]
    public void An_id_that_does_not_carry_its_owner_prefix_is_not_classified()
    {
        // PrimitiveId is caller-composed and the base constructor does not prove the
        // owner prefix, so a mismatched id is possible in principle. The predicate
        // must fail closed rather than scanning the identifier for a prefix.
        ScenePrimitive mismatched = new SceneRect(
            "chart:panel-cell:Type",
            SceneOwnerId.ForRow(GanttRowId.New()),
            ZLayer.Background,
            new RectD(0, 0, 10, 10),
            new SceneStyle("DataPanelFill"));

        Assert.False(PanelPrimitiveRoles.IsPanelPrimitive(mismatched));
    }

    [Fact]
    public void Null_is_not_a_panel_primitive() =>
        Assert.False(PanelPrimitiveRoles.IsPanelPrimitive(null));

    private static PanelBuildResult Build()
    {
        PanelColumn[] columns =
        [
            .. GanttTableSchema.Default.Columns.Select(
                (column, index) => new PanelColumn(column.Name, 20.0 + index)),
        ];
        PanelCellGrid grid = PanelCellGrid
            .TryCreate(
                columns,
                [RowHeight, RowHeight],
                HeaderHeight,
                [.. GanttTableSchema.Default.Columns.Where(column => column.IsRequired).Select(column => column.Name)])
            .Grid!;

        string[] cells = [.. Enumerable.Repeat("x", GanttTableSchema.Default.Columns.Count)];
        PanelRow[] rows =
        [
            new(GanttRowId.New(), cells),
            new(GanttRowId.New(), cells),
        ];

        PanelBuildRequest request = new(
            grid,
            rows,
            new RectD(PlotLeft, PlotTop, PlotWidth, PlotHeight),
            PlotTop - YearBandHeight,
            _theme);

        PanelBuildOutcome outcome = PanelBuilder.TryBuild(request);
        Assert.True(outcome.Succeeded, "Panel build refused: " + outcome.Refusal);
        return outcome.Result!;
    }
}
