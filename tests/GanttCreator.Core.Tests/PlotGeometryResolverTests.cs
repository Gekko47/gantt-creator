using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests;

/// <summary>
/// R4.7H: <c>PlotGeometryResolver</c> is the single authority for plot bounds, the
/// formula is fixed, the measured text panel is never resized, and an insufficient
/// remainder refuses with the shortfall.
/// </summary>
public sealed class PlotGeometryResolverTests
{
    private const double TolerancePt = 1e-9;

    private static PlotGeometryOutcome Resolve(
        SizePreset preset,
        double panel,
        double left = 0,
        double right = 0,
        double top = 60,
        double height = 140) =>
        PlotGeometryResolver.TryResolve(preset, panel, left, right, top, height);

    [Fact]
    public void Plot_width_is_the_preset_width_minus_the_panel_and_the_chrome()
    {
        // D2, exactly. Everything the preset does not consume becomes plot width.
        PlotGeometryOutcome outcome = Resolve(SizePresets.Presentation16x9, 400, left: 10, right: 6);

        Assert.True(outcome.Succeeded);
        Assert.Equal(1280.0 - 400 - 10 - 6, outcome.Geometry!.PlotBounds.Width, TolerancePt);
    }

    [Theory]
    [InlineData(200.0)]
    [InlineData(400.0)]
    [InlineData(600.0)]
    public void A_wider_text_panel_gives_a_proportionally_narrower_plot(double panel)
    {
        PlotGeometryOutcome outcome = Resolve(SizePresets.Presentation16x9, panel);

        Assert.True(outcome.Succeeded);
        Assert.Equal(1280.0 - panel, outcome.Geometry!.PlotBounds.Width, TolerancePt);
    }

    [Fact]
    public void The_measured_text_panel_width_is_carried_through_unchanged()
    {
        // D3. The user's column widths are the input, not the output: a preset
        // that resized them would silently rewrite the part of the chart they read.
        const double measured = 517.25;
        PlotGeometryOutcome outcome = Resolve(SizePresets.A4Landscape, measured);

        Assert.True(outcome.Succeeded);
        Assert.Equal(measured, outcome.Geometry!.TextPanelWidthPt, TolerancePt);
    }

    [Fact]
    public void The_plot_origin_sits_after_the_panel_and_the_left_chrome()
    {
        PlotGeometryOutcome outcome = Resolve(SizePresets.Presentation16x9, 400, left: 12);

        Assert.Equal(412.0, outcome.Geometry!.PlotBounds.X, TolerancePt);
    }

    [Fact]
    public void The_plot_never_crosses_the_presets_right_edge()
    {
        PlotGeometryOutcome outcome = Resolve(SizePresets.Presentation16x9, 900, left: 20, right: 20);

        Assert.True(outcome.Succeeded);
        Assert.True(outcome.Geometry!.PlotBounds.Right <= SizePresets.Presentation16x9.WidthPt);
    }

    [Fact]
    public void An_insufficient_remainder_refuses_and_names_the_shortfall()
    {
        // D4, as a POSITIVE test for the refusal path. The number is the point:
        // it tells the user how much to narrow a column.
        PlotGeometryOutcome outcome = Resolve(SizePresets.Presentation16x9, 1270);

        Assert.False(outcome.Succeeded);
        Assert.Equal(PlotGeometryRefusal.InsufficientPlotWidth, outcome.Refusal);
        Assert.Equal(
            PlotGeometryResolver.MinimumPlotWidthPt - 10.0,
            outcome.Geometry!.ShortfallPt,
            TolerancePt);
    }

    [Fact]
    public void An_insufficient_remainder_does_not_shrink_the_text_panel()
    {
        // The refusal must leave the measurement exactly as it was. Shrinking it
        // here would be D3's defect arriving through the refusal path.
        const double measured = 1270.0;
        PlotGeometryOutcome outcome = Resolve(SizePresets.Presentation16x9, measured);

        Assert.False(outcome.Succeeded);
        Assert.Equal(measured, outcome.Geometry!.TextPanelWidthPt, TolerancePt);
    }

    [Fact]
    public void A_panel_wider_than_the_preset_refuses_rather_than_a_negative_plot()
    {
        // RectD refuses a negative extent, so a naive implementation would throw or
        // silently clamp. The typed refusal is the third option.
        PlotGeometryOutcome outcome = Resolve(SizePresets.Presentation16x9, 5000);

        Assert.False(outcome.Succeeded);
        Assert.Equal(PlotGeometryRefusal.InsufficientPlotWidth, outcome.Refusal);
    }

    [Fact]
    public void The_same_panel_under_two_presets_differs_only_in_plot_width()
    {
        // The work item's acceptance test. Date range and scale are not inputs to
        // this resolver, so they cannot vary here; what must hold is that two
        // presets over the same panel differ only in the plot they hand on.
        PlotGeometryOutcome wide = Resolve(SizePresets.Presentation16x9, 400);
        PlotGeometryOutcome narrow = Resolve(SizePresets.A4Landscape, 400);

        Assert.True(wide.Succeeded);
        Assert.True(narrow.Succeeded);
        Assert.Equal(400.0, wide.Geometry!.TextPanelWidthPt, TolerancePt);
        Assert.Equal(400.0, narrow.Geometry!.TextPanelWidthPt, TolerancePt);
        Assert.True(wide.Geometry!.PlotBounds.Width > narrow.Geometry!.PlotBounds.Width);
    }

    [Fact]
    public void An_absent_preset_refuses_rather_than_defaulting_one()
    {
        // Defaulting would render a chart the user never chose. The caller decides
        // whether to fall back, and can report that it did.
        PlotGeometryOutcome outcome = PlotGeometryResolver.TryResolve(null, 400, 0, 0, 60, 140);

        Assert.False(outcome.Succeeded);
        Assert.Equal(PlotGeometryRefusal.NoPreset, outcome.Refusal);
        Assert.Null(outcome.Geometry);
    }

    [Fact]
    public void A_degenerate_preset_refuses_rather_than_producing_a_plot()
    {
        var broken = new SizePreset("Broken", "Broken", 0, 100, SizeOrientation.Portrait);

        PlotGeometryOutcome outcome = Resolve(broken, 10);

        Assert.False(outcome.Succeeded);
        Assert.Equal(PlotGeometryRefusal.InvalidPresetDimensions, outcome.Refusal);
    }

    [Theory]
    [InlineData(-1.0)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    public void An_invalid_measured_panel_width_refuses_rather_than_being_clamped(double panel)
    {
        // Clamping a negative measurement would launder the defect into a
        // plausible-looking plot, which is how a wrong panel width reaches a chart.
        PlotGeometryOutcome outcome = Resolve(SizePresets.Presentation16x9, panel);

        Assert.False(outcome.Succeeded);
        Assert.Equal(PlotGeometryRefusal.InvalidTextPanelWidth, outcome.Refusal);
    }

    [Fact]
    public void Negative_chrome_refuses_rather_than_widening_the_plot()
    {
        PlotGeometryOutcome outcome = Resolve(SizePresets.Presentation16x9, 400, left: -50);

        Assert.False(outcome.Succeeded);
        Assert.Equal(PlotGeometryRefusal.InvalidChrome, outcome.Refusal);
    }

    [Fact]
    public void A_non_positive_plot_height_refuses()
    {
        PlotGeometryOutcome outcome =
            PlotGeometryResolver.TryResolve(SizePresets.Presentation16x9, 400, 0, 0, 60, 0);

        Assert.False(outcome.Succeeded);
        Assert.Equal(PlotGeometryRefusal.InvalidOrigin, outcome.Refusal);
    }

    [Fact]
    public void The_resolved_geometry_records_the_preset_it_came_from()
    {
        PlotGeometryOutcome outcome = Resolve(SizePresets.Presentation4x3, 300);

        Assert.Same(SizePresets.Presentation4x3, outcome.Geometry!.Preset);
    }

    [Fact]
    public void A_fully_consumed_preset_refuses_and_reports_the_whole_minimum_as_short()
    {
        // The extreme edge: a panel exactly as wide as the preset leaves zero.
        PlotGeometryOutcome outcome = Resolve(SizePresets.Presentation16x9, 1280);

        Assert.False(outcome.Succeeded);
        Assert.Equal(PlotGeometryResolver.MinimumPlotWidthPt, outcome.Geometry!.ShortfallPt, TolerancePt);
    }
}
