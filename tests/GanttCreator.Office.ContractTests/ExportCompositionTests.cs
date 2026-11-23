using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R8.7 D4: the shared export-composition rule and the host-free export-aspect
/// record that feeds the Ribbon preview.
/// </summary>
/// <remarks>
/// The aspect and the selection are one authority for the preview, the Export PNG
/// command, Copy Editable and Send to PowerPoint, so these facts pin both the
/// value's guards and the composition rule. Every guard ships with its positive
/// test in this file, per the repository's validator rule.
/// </remarks>
public sealed class ExportCompositionTests
{
    /// <summary>
    /// The managed worksheet row height this fixture measures, taken from the
    /// code-owned catalogue rather than written as a literal — the same rule the
    /// scene-request factory's fixture follows, so a token change moves both.
    /// </summary>
    private static readonly double _managedRowHeight = GanttCatalogues.MetricDefault("GanttRowHeightPt");

    /// <summary>A valid span event for the composition fixtures.</summary>
    private static GanttEvent Event(DateOnly start, DateOnly finish) =>
        new(
            RowNumber: 1,
            Id: GanttRowId.New(),
            LaneId: GanttRowId.New(),
            StackIndex: 0,
            Type: GanttEntityType.AsPlannedActivity,
            Description: "Design",
            Start: start,
            Finish: finish,
            ParentId: null,
            StyleKey: "AsPlannedActivity",
            LabelPosition: null,
            FillColour: null,
            StrokeColour: null,
            Visible: true,
            SortOrder: 0);

    /// <summary>A measured grid the export composition accepts.</summary>
    private static PanelCellGrid Grid() =>
        PanelCellGrid
            .TryCreate(
                [new PanelColumn("Id", 120)],
                [_managedRowHeight],
                _managedRowHeight,
                ["Id"],
                58,
                0,
                leadingGutterPt: 0)
            .Grid!;

    /// <summary>A registry with the one style the fixture event references.</summary>
    private static GanttStyleRegistry Registry() =>
        new([
            new GanttStyleDefinition(
                "AsPlannedActivity",
                new HashSet<GanttLabelPosition> { GanttLabelPosition.Inside },
                EntityColourCapability.Fill | EntityColourCapability.Stroke,
                DefaultLabelPosition: GanttLabelPosition.Inside,
                FillColour: "#DDEBF7",
                StrokeColour: "#2E75B6",
                TextColour: "#000000",
                HatchPattern: GanttHatchPattern.None,
                StandardOutlinePt: 1,
                ActivityHeightPt: 12),
        ]);

    /// <summary>An extent that is not finite and strictly positive is refused at construction.</summary>
    [Theory]
    [InlineData(double.NaN, 10)]
    [InlineData(double.PositiveInfinity, 10)]
    [InlineData(0, 10)]
    [InlineData(-1, 10)]
    [InlineData(10, double.NaN)]
    [InlineData(10, double.PositiveInfinity)]
    [InlineData(10, 0)]
    [InlineData(10, -1)]
    public void An_extent_that_is_not_finite_and_positive_is_refused(double widthPt, double heightPt)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new ExportAspectPt(widthPt, heightPt));
    }

    /// <summary>The width guard names the width, so a caller cannot confuse the two axes.</summary>
    [Fact]
    public void The_width_guard_names_the_width()
    {
        ArgumentOutOfRangeException thrown =
            Assert.Throws<ArgumentOutOfRangeException>(() => new ExportAspectPt(0, 10));

        Assert.Equal("widthPt", thrown.ParamName);
    }

    /// <summary>The height guard names the height.</summary>
    [Fact]
    public void The_height_guard_names_the_height()
    {
        ArgumentOutOfRangeException thrown =
            Assert.Throws<ArgumentOutOfRangeException>(() => new ExportAspectPt(10, 0));

        Assert.Equal("heightPt", thrown.ParamName);
    }

    /// <summary>A valid extent builds an aspect that carries its values verbatim.</summary>
    [Fact]
    public void A_valid_extent_builds_an_aspect()
    {
        ExportAspectPt aspect = new(841.89, 479.5);

        Assert.Equal(841.89, aspect.WidthPt, 6);
        Assert.Equal(479.5, aspect.HeightPt, 6);
        Assert.True(ExportAspectPt.IsUsable(aspect.WidthPt, aspect.HeightPt));
        Assert.False(ExportAspectPt.IsUsable(0, 10));
        Assert.False(ExportAspectPt.IsUsable(10, double.NaN));
    }

    /// <summary>
    /// The selection is chart-only only on an explicit case-insensitive FALSE;
    /// everything else fails open to table-plus-chart (the catalogue default).
    /// </summary>
    [Theory]
    [InlineData(null, ExportSelection.TablePlusChart)]
    [InlineData("TRUE", ExportSelection.TablePlusChart)]
    [InlineData("true", ExportSelection.TablePlusChart)]
    [InlineData("garbage", ExportSelection.TablePlusChart)]
    [InlineData("FALSE", ExportSelection.ChartOnly)]
    [InlineData(" false ", ExportSelection.ChartOnly)]
    public void The_selection_is_chart_only_only_on_an_explicit_false(
        string? stored,
        ExportSelection expected)
    {
        Dictionary<string, string> settings = [];
        if (stored is not null)
        {
            settings[ExportComposition.ExportIncludeDataPanelSettingKey] = stored;
        }

        Assert.Equal(expected, ExportComposition.SelectExportSelection(settings));
    }

    /// <summary>Null settings are a caller defect and are refused.</summary>
    [Fact]
    public void Null_settings_are_refused()
    {
        Assert.Throws<ArgumentNullException>(
            () => ExportComposition.SelectExportSelection(null!));
    }

    /// <summary>
    /// The aspect equals the bounds the export composition independently resolves
    /// for the same inputs — the preview's input and the command's input are one
    /// function (R8.7 D4/D14).
    /// </summary>
    [Fact]
    public void The_aspect_equals_the_bounds_the_export_composition_resolves()
    {
        GanttEvent[] events = [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))];
        Dictionary<string, string> settings = [];
        GanttStyleRegistry registry = Registry();
        PanelCellGrid grid = Grid();
        Dictionary<string, double> metrics = [];

        ExportAspectPt? aspect = ExportComposition.TryResolveAspect(events, settings, registry, grid, metrics);
        Assert.NotNull(aspect);

        SceneBuildRequestOutcome request = EditableExportRequest.Create(events, settings, registry, grid, metrics);
        Assert.True(request.Succeeded, "export request refused: " + request.Message);

        ExportBoundsOutcome resolution = ExportBoundsResolver.TryResolve(
            request.Request,
            ExportPanelTheme.Resolve(metrics),
            ExportSelection.TablePlusChart);
        Assert.True(resolution.Succeeded, "bounds refused: " + resolution.Refusal);

        Assert.Equal(resolution.Bounds!.Bounds.Width, aspect.Value.WidthPt, 6);
        Assert.Equal(resolution.Bounds!.Bounds.Height, aspect.Value.HeightPt, 6);
    }

    /// <summary>
    /// An explicit chart-only selection resolves against the chart bounds with no
    /// panel theme, exactly as the command's pre-flight does.
    /// </summary>
    [Fact]
    public void A_chart_only_selection_resolves_against_the_chart_bounds()
    {
        GanttEvent[] events = [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))];
        Dictionary<string, string> settings =
            new() { [ExportComposition.ExportIncludeDataPanelSettingKey] = "FALSE" };
        GanttStyleRegistry registry = Registry();
        PanelCellGrid grid = Grid();
        Dictionary<string, double> metrics = [];

        ExportAspectPt? aspect = ExportComposition.TryResolveAspect(events, settings, registry, grid, metrics);
        Assert.NotNull(aspect);

        SceneBuildRequestOutcome request = EditableExportRequest.Create(events, settings, registry, grid, metrics);
        Assert.True(request.Succeeded, "export request refused: " + request.Message);

        ExportBoundsOutcome resolution = ExportBoundsResolver.TryResolve(
            request.Request,
            panelTheme: null,
            ExportSelection.ChartOnly);
        Assert.True(resolution.Succeeded, "bounds refused: " + resolution.Refusal);

        Assert.Equal(resolution.Bounds!.Bounds.Width, aspect.Value.WidthPt, 6);
        Assert.Equal(resolution.Bounds!.Bounds.Height, aspect.Value.HeightPt, 6);
    }

    /// <summary>
    /// A composition that cannot be built records no aspect: the empty table is a
    /// supported workbook state and the preview shows the not-available state.
    /// </summary>
    [Fact]
    public void A_composition_that_cannot_be_built_records_no_aspect()
    {
        ExportAspectPt? aspect = ExportComposition.TryResolveAspect(
            [],
            new Dictionary<string, string>(),
            Registry(),
            Grid(),
            new Dictionary<string, double>());

        Assert.Null(aspect);
    }

    /// <summary>Every input is required; a null is a caller defect, not a refusal.</summary>
    [Fact]
    public void Null_inputs_are_refused()
    {
        GanttEvent[] events = [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))];
        Dictionary<string, string> settings = [];
        Dictionary<string, double> metrics = [];

        Assert.Throws<ArgumentNullException>(
            () => ExportComposition.TryResolveAspect(null!, settings, Registry(), Grid(), metrics));
        Assert.Throws<ArgumentNullException>(
            () => ExportComposition.TryResolveAspect(events, null!, Registry(), Grid(), metrics));
        Assert.Throws<ArgumentNullException>(
            () => ExportComposition.TryResolveAspect(events, settings, null!, Grid(), metrics));
        Assert.Throws<ArgumentNullException>(
            () => ExportComposition.TryResolveAspect(events, settings, Registry(), null!, metrics));
        Assert.Throws<ArgumentNullException>(
            () => ExportComposition.TryResolveAspect(events, settings, Registry(), Grid(), null!));
    }
}
