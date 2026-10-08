using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R4.8A D5: the scene-request factory is the single authority for every scene
/// input, and the profile it emits is the one D4 requires.
/// </summary>
/// <remarks>
/// No Excel: the factory takes four values and returns a request, so every claim here
/// is about the translation itself rather than about a host.
/// </remarks>
public class ExcelSceneBuildRequestFactoryTests
{
    /// <summary>
    /// The managed worksheet row height this fixture measures, taken from the
    /// code-owned catalogue rather than written as a literal.
    /// </summary>
    /// <remarks>
    /// It was 15, which is shorter than the <c>GanttRowHeightPt</c> default of 18. A
    /// label box is now one row tall, so an 18pt box inside a 15pt plot cannot be
    /// contained by the chart bounds and every label was suppressed with
    /// <c>LabelSuppressedNoSpace</c> - a fixture that measured a row height the
    /// product never produces. Reading the token keeps the two in step: if the
    /// catalogue row height changes, this fixture follows it.
    /// </remarks>
    private static readonly double _managedRowHeight = GanttCatalogues.MetricDefault("GanttRowHeightPt");

    private static GanttEvent Event(DateOnly start, DateOnly finish, int rowNumber = 1) =>
        new(
            RowNumber: rowNumber,
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
            SortOrder: 0
        );

    /// <summary>
    /// A measured grid narrow enough to leave a readable A4-landscape plot.
    /// </summary>
    /// <param name="widthPt">The panel's total measured width.</param>
    /// <param name="originTopPt">The first body row's absolute top edge (ADR-0030 D3).</param>
    /// <param name="rowHeightsPt">The measured body row heights.</param>
    private static PanelCellGrid Grid(double widthPt = 120, double originTopPt = 58, double[]? rowHeightsPt = null) =>
        PanelCellGrid
            .TryCreate([new PanelColumn("Id", widthPt)], rowHeightsPt ?? [_managedRowHeight], _managedRowHeight, ["Id"], originTopPt)
            .Grid!;

    /// <summary>
    /// The LIVE request anchors its lanes to the measured worksheet rows (ADR-0034 D1).
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>This is the wiring test, and it exists because of an observed measurement.</b>
    /// Setting this flag to <see langword="false"/> in the factory disables row
    /// anchoring for every live refresh and restores the pre-ADR-0034 stacking that
    /// leaves bars one row away from their cells. Before this test, that mutation left
    /// the entire suite green: the golden snapshot is built on the <c>Raster</c>
    /// profile with anchoring off, and the Core resolver tests exercise the resolver
    /// directly rather than through this factory. So nothing asserted that the LIVE
    /// composition actually asks for anchoring.
    /// </para>
    /// <para>
    /// The assertion is on the request's own flag rather than on the resulting geometry,
    /// because the geometry is asserted where it belongs — in
    /// <c>LaneRowAnchorResolverTests</c> and
    /// <c>An_anchored_lane_coincides_with_its_measured_row</c>. This test's single claim
    /// is that the live path turns the feature on.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_live_request_anchors_its_lanes_to_the_measured_rows()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            new Dictionary<string, string>(),
            StyleRegistry(),
            Grid()
        );

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);

        // Anchoring needs the measured grid, so the flag and the grid travel together.
        // Asserting both keeps the test honest: a future edit that set the flag without
        // carrying the grid would be refused by SceneBuilder rather than silently
        // stacked, and this assertion says which of the two went missing.
        Assert.True(outcome.Request!.AnchorLanesToRows);
        Assert.NotNull(outcome.Request.Grid);
    }

    /// <summary>
    /// The anchored live request places each lane on its own measured row, so a row's
    /// bar moves with the MEASURED heights rather than with the fixed lane metric.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>Why this needs three rows and an uneven middle.</b> A single-row scene cannot
    /// tell anchoring from stacking: both put the first lane at offset 0, and the bar's
    /// height comes from the resolved style, not from the lane, so a one-row assertion
    /// passes under both. The discriminating case is a LATER row, whose top depends on
    /// the rows above it — and only when those heights differ from the lane metric.
    /// </para>
    /// <para>
    /// Rows of 18 and 30pt make the third lane's top 48 (anchored: 18 + 30) against 36
    /// (stacked: 18 + the 18pt lane metric). The 12pt gap is the whole feature: under
    /// stacking the third row's bar would sit a row above its cell. This test fails with
    /// <c>AnchorLanesToRows = false</c>, which is how its non-vacuity was established.
    /// </para>
    /// </remarks>
    [Fact]
    public void An_anchored_live_request_places_each_lane_on_its_own_measured_row()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());
        GanttEvent[] events =
        [
            Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19), rowNumber: 1),
            Event(new DateOnly(2024, 1, 22), new DateOnly(2024, 2, 2), rowNumber: 2),
            Event(new DateOnly(2024, 2, 5), new DateOnly(2024, 2, 16), rowNumber: 3),
        ];

        SceneBuildRequestOutcome outcome = factory.Create(
            events,
            new Dictionary<string, string>(),
            StyleRegistry(),
            Grid(originTopPt: 58, rowHeightsPt: [18, 30, 18])
        );

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);
        Assert.True(outcome.Request!.AnchorLanesToRows);

        SceneBuildOutcome built = SceneBuilder.TryBuild(outcome.Request);
        Assert.True(built.Succeeded, "the anchored live request was refused: " + built.Refusal);

        // The bars are ordered by lane, so the third scene bar is the third row's.
        List<SceneRect> bars =
        [
            .. built
                .Result!.Scene.Primitives.OfType<SceneRect>()
                .Where(primitive => primitive.PrimitiveId.Contains(":bar", StringComparison.Ordinal))
                .OrderBy(primitive => primitive.LaneOrder ?? 0),
        ];

        Assert.Equal(3, bars.Count);

        // Each bar sits inside its own measured row: plot top (58) + the lane's anchor
        // offset + the centring inset. ADR-0034 D2 centres the content block between the
        // paddings, so an 18pt row with a 12pt bar and 3pt paddings has no free space
        // and the bar starts at paddingTop, while the 30pt middle row has 12pt free and
        // puts 6pt above the block.
        //   lane 0: 58 + (0 + 3)                    = 61
        //   lane 1: 58 + (18 + 3 + 6)               = 85
        //   lane 2: 58 + (18 + 30 + 3)              = 109
        Assert.Equal([61, 85, 109], bars.Select(bar => bar.Bounds.Top));

        // The gap between consecutive bars is the SPACING between their rows, and it
        // follows the MEASURED heights rather than a constant. This is the contract in
        // one assertion:
        //   bar1 - bar0 = 24: row 1 is 30pt, minus the 6pt its centring puts above
        //   bar2 - bar1 = 24: row 2 is 18pt with no free space, plus row 1's 6pt inset
        //                    that row 2 inherits by starting below it
        // Stacking instead adds the 18pt lane metric every time, giving 61/79/97 and a
        // uniform 18pt gap — the pre-ADR-0034 misalignment this row removed, where the
        // third bar sits a full row above its cell. Established by mutation: setting
        // AnchorLanesToRows = false fails this test with gaps of 18 and 18.
        Assert.Equal([24, 24], [bars[2].Bounds.Top - bars[1].Bounds.Top, bars[1].Bounds.Top - bars[0].Bounds.Top]);
    }

    private static FakeTextMetrics Metrics() => new();

    /// <summary>
    /// The measured set is every VISIBLE schema column, and it is derived from the
    /// schema rather than restated in the factory.
    /// </summary>
    /// <remarks>
    /// This is the regression test for the live "The worksheet columns could not be
    /// measured" failure. <c>MeasuredColumns</c> was the literal
    /// <c>["Id", "Type", "Description"]</c> and <c>Id</c> is an <c>EngineHidden</c>
    /// column: step 4 of the refresh hides it, Excel reports a hidden column's
    /// <c>Range.Width</c> as <c>0</c>, and <c>PanelCellGrid.TryCreate</c> refuses a
    /// non-positive width. Every live refresh therefore refused at
    /// <c>MeasurementRefused</c>.
    /// </remarks>
    [Fact]
    public void Measured_columns_are_exactly_the_visible_schema_columns_in_schema_order()
    {
        Assert.Equal(
            GanttTableSchema.Default.Columns.Where(column => !column.IsHidden).Select(column => column.Name),
            ExcelSceneBuildRequestFactory.MeasuredColumns
        );
    }

    /// <summary>
    /// The positive assertion of the fix, and the one that would have caught it:
    /// <b>no measured column may be a hidden column.</b>
    /// </summary>
    /// <remarks>
    /// Asserting the exact list alone would still pass if the schema were changed to
    /// hide <c>Type</c>, so the invariant is asserted directly as well. This is the
    /// pairing the failure needed: the check that a hidden column is measured at all,
    /// which is the condition that produces a zero width and the refusal.
    /// </remarks>
    [Fact]
    public void Measured_columns_contain_no_hidden_column_because_a_hidden_one_measures_zero()
    {
        IReadOnlyList<string> measured = ExcelSceneBuildRequestFactory.MeasuredColumns;

        List<string> hidden = [.. GanttTableSchema.Default.Columns.Where(c => c.IsHidden).Select(c => c.Name)];

        Assert.DoesNotContain(measured, name => hidden.Contains(name));
    }

    /// <summary>
    /// The measured set is not empty, so the measurement is never refused for want of
    /// columns.
    /// </summary>
    /// <remarks>
    /// <c>Measure</c> refuses an empty column list outright, so a schema change that
    /// hid every column would otherwise surface as the same user-facing
    /// "columns could not be measured" message with a different cause.
    /// </remarks>
    [Fact]
    public void Measured_columns_are_not_empty_because_an_empty_set_is_refused_outright()
    {
        Assert.NotEmpty(ExcelSceneBuildRequestFactory.MeasuredColumns);
    }

    /// <summary>
    /// The plot's left edge starts exactly where the measured panel ends, with NO
    /// left chrome (ADR-0031 D2).
    /// </summary>
    /// <remarks>
    /// <para>
    /// Measuring only the label columns would leave the plot drawn on top of the
    /// still-visible <c>Start</c>, <c>Finish</c>, and <c>Duration</c> columns, so this
    /// asserts the placement rather than merely the list.
    /// </para>
    /// <para>
    /// <b>This assertion was inverted, deliberately.</b> It previously required
    /// <c>panel + ChartOuterPaddingPt</c>, which is the behaviour the owner asked to
    /// change: the plot sat one margin-width away from the table it belongs to. The
    /// margin is now exactly zero, and the test is the one place that will fail if a
    /// later change reintroduces a gap.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_plot_left_edge_starts_flush_where_the_measured_panel_ends()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            new Dictionary<string, string>(),
            StyleRegistry(),
            Grid(widthPt: 120)
        );

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);

        SceneBuildRequest request = Assert.IsType<SceneBuildRequest>(outcome.Request);

        // Both `Grid` and `PlotBounds` are nullable on the request and are unwrapped
        // with the `!.Value`/`!` form the rest of this file already uses. The factory
        // only emits a request whose grid and plot both resolved, so a null here is a
        // defect the equality assertion reports as a mismatch rather than a null crash.
        Assert.Equal(request.Grid!.TotalWidthPt, request.PlotBounds!.Value.X, precision: 6);

        // The right margin is NOT removed: the plot's right edge is the sheet's own
        // edge, and the chrome there is what keeps the last period label off it.
        Assert.Equal(GanttCatalogues.MetricDefault("ChartOuterPaddingPt"), request.ChartPadding.RightPt, precision: 6);
    }

    /// <summary>
    /// The plot's TOP is the first body row's measured top edge (ADR-0030 D1).
    /// </summary>
    /// <remarks>
    /// <b>This is the test for the reported defect.</b> The plot's top used to be a
    /// fixed token sum — <c>TitleBand + YearBand + PeriodBand</c> = 58pt — regardless
    /// of the worksheet, which put lane 0 roughly 43pt (2.4 rows) below its own row.
    /// The value here is deliberately neither 58 nor zero, so neither the old page
    /// origin nor a default can satisfy it.
    /// </remarks>
    /// <summary>
    /// Builds a request over one event and the supplied measured grid.
    /// </summary>
    /// <param name="grid">The measured grid the scene is composed against.</param>
    /// <returns>The factory outcome.</returns>
    private static SceneBuildRequestOutcome Create(PanelCellGrid grid)
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());
        return factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            new Dictionary<string, string>(),
            StyleRegistry(),
            grid
        );
    }

    /// <summary>
    /// Builds a request over one span-dated event, for the plot-range assertions.
    /// </summary>
    /// <param name="start">The event's start date.</param>
    /// <param name="finish">The event's finish date; defaults to <paramref name="start"/>.</param>
    /// <returns>The factory outcome.</returns>
    private static SceneBuildRequestOutcome CreateForRange(DateOnly start, DateOnly? finish = null) =>
        new ExcelSceneBuildRequestFactory(Metrics()).Create(
            [Event(start, finish ?? start)],
            new Dictionary<string, string>(),
            StyleRegistry(),
            Grid()
        );

    [Fact]
    public void The_plot_top_is_the_first_body_rows_measured_top()
    {
        SceneBuildRequestOutcome outcome = Create(Grid(originTopPt: 137.5, rowHeightsPt: [15.0, 15.0, 15.0]));

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);

        SceneBuildRequest request = Assert.IsType<SceneBuildRequest>(outcome.Request);
        Assert.Equal(137.5, request.PlotBounds!.Value.Y, precision: 6);
    }

    /// <summary>
    /// The plot's HEIGHT is the measured body total, so the chart ends with the table
    /// (ADR-0030 D2).
    /// </summary>
    /// <remarks>
    /// Previously the height was the remaining <em>page</em> height (~531pt on A4
    /// landscape), so a 3-row chart ran ~400pt past the last row. The sum here is
    /// 45pt, and it is asserted rather than any page figure.
    /// </remarks>
    [Fact]
    public void The_plot_height_is_the_measured_body_total_not_the_page_height()
    {
        SceneBuildRequestOutcome outcome = Create(Grid(rowHeightsPt: [15.0, 15.0, 15.0]));

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);

        SceneBuildRequest request = Assert.IsType<SceneBuildRequest>(outcome.Request);
        RectD plot = request.PlotBounds!.Value;

        Assert.Equal(45.0, plot.Height, precision: 6);
        Assert.Equal(45.0, request.Grid!.TotalRowHeightPt, precision: 6);
    }

    /// <summary>
    /// The plot's bottom edge is the last row's bottom edge.
    /// </summary>
    /// <remarks>
    /// Stated as a relationship rather than a number because it is the property the
    /// user sees: the chart must stop where the table stops. Before ADR-0030 the chart
    /// ran far past it, so nothing asserted this.
    /// </remarks>
    [Fact]
    public void The_plot_bottom_coincides_with_the_last_rows_bottom()
    {
        SceneBuildRequestOutcome outcome = Create(Grid(originTopPt: 64, rowHeightsPt: [18.0, 18.0, 24.0, 18.0]));

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);

        SceneBuildRequest request = Assert.IsType<SceneBuildRequest>(outcome.Request);
        RectD plot = request.PlotBounds!.Value;

        Assert.Equal(64.0 + 18.0 + 18.0 + 24.0 + 18.0, plot.Y + plot.Height, precision: 6);
    }

    /// <summary>
    /// A table taller than the preset's page still renders (ADR-0030 D7).
    /// </summary>
    /// <remarks>
    /// <b>The counterweight to the height test above.</b> With the page still bounding
    /// the vertical extent, a 40-row schedule on A4 landscape (29 rows tall) would be
    /// refused outright — the same defect as lane auto-growth, wearing a page budget
    /// instead of a row height. The width budget is untouched, so an over-wide panel
    /// still refuses.
    /// </remarks>
    [Fact]
    public void A_table_taller_than_the_page_is_not_refused()
    {
        double[] fortyRows = [.. Enumerable.Repeat(18.0, 40)];

        SceneBuildRequestOutcome outcome = Create(Grid(rowHeightsPt: fortyRows));

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);

        SceneBuildRequest request = Assert.IsType<SceneBuildRequest>(outcome.Request);
        Assert.Equal(40 * 18.0, request.PlotBounds!.Value.Height, precision: 6);
    }

    /// <summary>
    /// A panel too wide to leave a readable plot is still refused (R4.7H D4).
    /// </summary>
    /// <remarks>
    /// The counterweight that D7 did not loosen: dropping the page bound is about the
    /// <em>vertical</em> extent only. The width budget still refuses a panel that
    /// would leave an unreadable plot, so "the page no longer bounds it" cannot be
    /// read as "the page no longer constrains anything".
    /// </remarks>
    /// <summary>
    /// The plot extent is the whole month containing the padded earliest and latest
    /// dates (owner ruling, 2026-10-02).
    /// </summary>
    /// <remarks>
    /// Both cases were stated by the owner and they pin the ORDER of the two steps,
    /// because padding-then-snap and snap-then-pad disagree about the 10 Jan case:
    /// <list type="bullet">
    /// <item>10 Jan - 3 = 7 Jan, still inside January, so it snaps back to 1 Jan.</item>
    /// <item>3 Jan - 3 = 31 Dec, which is inside December, so it snaps to 1 Dec -
    /// a whole month further out than 1 Jan.</item>
    /// </list>
    /// Snapping first and then padding would instead give 27 Dec for the first case,
    /// which is exactly the day-padded edge this change exists to remove.
    /// </remarks>
    [Theory]
    [InlineData(2025, 1, 10, 2025, 1, 1)] // mid-month: the pad is absorbed by the snap
    [InlineData(2025, 1, 3, 2024, 12, 1)] // near the start: the pad escapes a whole month
    [InlineData(2025, 1, 1, 2024, 12, 1)] // exactly on the boundary
    public void The_plot_starts_on_the_first_day_of_the_month_of_the_padded_earliest_date(
        int year,
        int month,
        int day,
        int expectedYear,
        int expectedMonth,
        int expectedDay
    )
    {
        SceneBuildRequestOutcome outcome = CreateForRange(new DateOnly(year, month, day));

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);
        Assert.Equal(new DateOnly(expectedYear, expectedMonth, expectedDay), outcome.Request!.PlotStart);
    }

    /// <summary>
    /// The plot ends on the last day of the month containing the padded latest date,
    /// which is what stops a bar being cut off mid-month at the right edge.
    /// </summary>
    [Theory]
    [InlineData(2025, 8, 10, 2025, 8, 31)] // a 31-day month
    [InlineData(2025, 2, 10, 2025, 2, 28)] // a non-leap February
    [InlineData(2024, 2, 10, 2024, 2, 29)] // a leap February
    [InlineData(2025, 12, 10, 2025, 12, 31)] // the year boundary
    public void The_plot_ends_on_the_last_day_of_the_month_of_the_padded_latest_date(
        int year,
        int month,
        int day,
        int expectedYear,
        int expectedMonth,
        int expectedDay
    )
    {
        SceneBuildRequestOutcome outcome = CreateForRange(new DateOnly(year, month, day));

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);
        Assert.Equal(new DateOnly(expectedYear, expectedMonth, expectedDay), outcome.Request!.PlotFinish);
    }

    /// <summary>
    /// A range whose padding stays inside its own months is not widened past them.
    /// </summary>
    /// <remarks>
    /// The dates are chosen so the 3-day pad does not cross a month boundary on
    /// either side, which is the only way a range can be genuinely unchanged by the
    /// snap. A range starting exactly on the 1st CANNOT be: 1 Jan pads to 29 Dec and
    /// therefore escapes to 1 Dec, the same as the owner's 3 Jan case. That is a
    /// consequence of padding-before-snapping and is asserted separately rather than
    /// hidden here.
    /// </remarks>
    [Fact]
    public void A_range_whose_padding_stays_inside_its_months_is_unchanged_by_the_snap()
    {
        SceneBuildRequestOutcome outcome = CreateForRange(new DateOnly(2025, 1, 10), new DateOnly(2025, 6, 25));

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);
        Assert.Equal(new DateOnly(2025, 1, 1), outcome.Request!.PlotStart);
        Assert.Equal(new DateOnly(2025, 6, 30), outcome.Request.PlotFinish);
    }

    [Fact]
    public void An_over_wide_panel_is_still_refused()
    {
        // The default preset is A4 landscape (841.89pt wide); at 830 the panel plus
        // both chromes leaves under 1pt of plot, below MinimumPlotWidthPt (24).
        SceneBuildRequestOutcome outcome = Create(Grid(widthPt: 830));

        Assert.False(outcome.Succeeded);
    }

    [Fact]
    public void The_live_profile_is_emitted_and_no_data_panel_is_supplied()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            new Dictionary<string, string>(),
            StyleRegistry(),
            Grid()
        );

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);

        // D4. Chosen in code, and the panel is null because the worksheet's own cells
        // ARE the panel. A drawn replica would cover the user's rows.
        Assert.Equal(SceneCompositionProfile.LiveExcel, outcome.Request!.Profile);
        Assert.Null(outcome.Request.Panel);
        Assert.False(SceneCompositionProfiles.DrawsDataPanel(outcome.Request.Profile));
    }

    /// <summary>
    /// The request the factory produces is accepted by the real scene builder.
    /// </summary>
    /// <remarks>
    /// This is the test that makes the factory useful rather than merely plausible. A
    /// factory that returned a request missing Metrics or LaneMetrics would satisfy
    /// every field-level assertion above and still be unusable in production, because
    /// the scene builder would refuse it.
    /// </remarks>
    [Fact]
    public void The_request_it_produces_is_accepted_by_the_scene_builder()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            new Dictionary<string, string>(),
            StyleRegistry(),
            Grid()
        );

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);

        SceneBuildOutcome built = SceneBuilder.TryBuild(outcome.Request);

        Assert.True(built.Succeeded, "the produced request was refused: " + built.Refusal);
        Assert.NotEmpty(built.Result!.Scene.Primitives);
    }

    /// <summary>The plot bounds come from the resolver, not from a second derivation.</summary>
    [Fact]
    public void The_plot_bounds_are_the_resolver_output_for_the_preset()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());
        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            new Dictionary<string, string>(),
            StyleRegistry(),
            Grid(widthPt: 200)
        );

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);
        Assert.Equal(SizePresets.Default, outcome.Request!.Preset);

        // A wider panel must leave a NARROWER plot on the same page, which is only
        // true if the plot is the remainder after the panel rather than a constant.
        // The resolver's rule is PlotWidth = presetWidth - panelWidth - chrome, so
        // the 200pt panel's plot is the narrower of the two.
        SceneBuildRequestOutcome widePanel = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            new Dictionary<string, string>(),
            StyleRegistry(),
            Grid(widthPt: 100)
        );

        Assert.True(widePanel.Succeeded);
        Assert.True(
            outcome.Request!.PlotBounds!.Value.Width < widePanel.Request!.PlotBounds!.Value.Width,
            "the 200pt panel's plot ("
                + outcome.Request.PlotBounds.Value.Width
                + "pt) must be narrower than the 100pt panel's plot ("
                + widePanel.Request.PlotBounds.Value.Width
                + "pt)"
        );
    }

    /// <summary>An unknown preset key refuses rather than silently falling back to A4.</summary>
    [Fact]
    public void An_unknown_size_preset_refuses_and_names_the_alternatives()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            new Dictionary<string, string> { ["SizePreset"] = "A3Gigantic" },
            StyleRegistry(),
            Grid()
        );

        Assert.False(outcome.Succeeded);
        Assert.Equal(SceneBuildRequestRefusal.UnknownSizePreset, outcome.Refusal);

        // The message must be actionable: it names the key that was rejected and the
        // keys that would work.
        Assert.Contains("A3Gigantic", outcome.Message!, StringComparison.Ordinal);
        Assert.Contains("A4Portrait", outcome.Message!, StringComparison.Ordinal);
    }

    /// <summary>
    /// A known preset key is honoured <em>when the factory is handed one</em>.
    /// </summary>
    /// <remarks>
    /// The name used to read "a known preset key is honoured", which overstated it. This
    /// test injects the key into a hand-built dictionary, so it proves the factory
    /// parses a preset it is <em>given</em> — it does NOT prove the workbook can give
    /// it one. That separate reachability claim is made by
    /// <see cref="A_preset_selected_in_a_catalogue_shaped_map_reaches_the_factory"/>.
    /// Both keys are real catalogue settings as of schema version 7; before that they
    /// were absent, so every case here was unreachable in production while this test
    /// stayed green.
    /// </remarks>
    [Theory]
    [InlineData("A4Portrait")]
    [InlineData("A4Landscape")]
    [InlineData("Presentation16x9")]
    [InlineData("Presentation4x3")]
    public void A_known_preset_key_is_honoured(string key)
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            new Dictionary<string, string> { ["SizePreset"] = key },
            StyleRegistry(),
            Grid()
        );

        Assert.True(outcome.Succeeded, key + " refused: " + outcome.Message);
        Assert.Equal(key, outcome.Request!.Preset!.Key);
    }

    /// <summary>
    /// The <c>SizePreset</c> key exists in the catalogue, and a value set against the
    /// catalogue's own key set reaches the factory.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This is the reachability test the injected-dictionary version of
    /// <c>A_known_preset_key_is_honoured</c> could never provide. That test proved the
    /// factory parses a preset it is <em>handed</em>; this one proves the catalogue
    /// actually hands it one, which is what schema version 7 fixed.
    /// </para>
    /// <para>
    /// The settings map is built from <c>GanttCatalogues.Settings</c> and only the
    /// <c>SizePreset</c> value is overridden per case — the rest stay at their catalogue
    /// defaults, exactly as a live workbook would read them. Before version 7 the key
    /// was absent, so every one of these cases was unreachable in production while the
    /// injected-dictionary test stayed green.
    /// </para>
    /// </remarks>
    [Theory]
    [InlineData("A4Portrait")]
    [InlineData("A4Landscape")]
    [InlineData("Presentation16x9")]
    [InlineData("Presentation4x3")]
    public void A_preset_selected_in_a_catalogue_shaped_map_reaches_the_factory(string key)
    {
        Dictionary<string, string> fromCatalogue = GanttCatalogues.Settings.ToDictionary(
            static setting => setting.Key,
            static setting => setting.DefaultValue,
            StringComparer.Ordinal
        );

        // The catalogue must actually offer this key, or the rest of the test is
        // vacuous — this is the assertion that would have caught the original gap.
        Assert.Contains("SizePreset", fromCatalogue.Keys);

        fromCatalogue["SizePreset"] = key;

        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            fromCatalogue,
            StyleRegistry(),
            Grid()
        );

        Assert.True(outcome.Succeeded, key + " refused: " + outcome.Message);
        Assert.Equal(key, outcome.Request!.Preset!.Key);
    }

    /// <summary>
    /// The catalogue's own default preset is the one a freshly initialised workbook
    /// renders with.
    /// </summary>
    /// <remarks>
    /// Before schema 7 there was no such row, so every live chart was A4-portrait
    /// regardless of anything the user did. This pins the default now that a default
    /// exists, so retuning it is a deliberate change.
    /// </remarks>
    [Fact]
    public void The_catalogue_default_preset_reaches_the_factory()
    {
        Dictionary<string, string> fromCatalogue = GanttCatalogues.Settings.ToDictionary(
            static setting => setting.Key,
            static setting => setting.DefaultValue,
            StringComparer.Ordinal
        );

        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            fromCatalogue,
            StyleRegistry(),
            Grid()
        );

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);
        Assert.Equal(fromCatalogue["SizePreset"], outcome.Request!.Preset!.Key);
    }

    /// <summary>A text panel too wide for the page refuses with the shortfall named.</summary>
    [Fact]
    public void A_text_panel_too_wide_for_the_preset_refuses_with_the_shortfall()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            new Dictionary<string, string>(),
            StyleRegistry(),
            Grid(widthPt: 5000)
        );

        Assert.False(outcome.Succeeded);
        Assert.Equal(SceneBuildRequestRefusal.PlotGeometryRefused, outcome.Refusal);
        Assert.Contains("too wide", outcome.Message!, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// The plot range covers whole months around the events rather than a stored
    /// range that could survive the data changing.
    /// </summary>
    /// <remarks>
    /// A whole March now renders from 1 Feb to 30 Apr: the 2-day pad carries 1 Mar
    /// back into February and 31 Mar forward into April, and each is then snapped to
    /// its own month's boundary. This asserted 28 Feb and 2 Apr, which were the raw
    /// padded days - the exact day-granular edges the month snap removes, so a bar
    /// could begin part-way through a month column.
    /// </remarks>
    [Fact]
    public void The_plot_range_is_derived_from_the_events_and_snapped_to_month_boundaries()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 3, 1), new DateOnly(2024, 3, 31))],
            new Dictionary<string, string> { ["RangePaddingDays"] = "2" },
            StyleRegistry(),
            Grid()
        );

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);
        Assert.Equal(new DateOnly(2024, 2, 1), outcome.Request!.PlotStart);
        Assert.Equal(new DateOnly(2024, 4, 30), outcome.Request.PlotFinish);
    }

    /// <summary>
    /// The stored padding setting still widens the range when it crosses a month.
    /// </summary>
    /// <remarks>
    /// This is what keeps <c>RangePaddingDays</c> meaningful now that the snap
    /// dominates. With the default 3-day pad the start would be 1 March either way;
    /// a 31-day pad carries 15 March back to 13 February and forward to 15 April, so
    /// the range grows to whole February and April and the setting is observable
    /// rather than dead.
    /// </remarks>
    [Fact]
    public void The_stored_padding_setting_still_widens_the_range_across_a_month()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 3, 15), new DateOnly(2024, 3, 15))],
            new Dictionary<string, string> { ["RangePaddingDays"] = "31" },
            StyleRegistry(),
            Grid()
        );

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);
        Assert.Equal(new DateOnly(2024, 2, 1), outcome.Request!.PlotStart);
        Assert.Equal(new DateOnly(2024, 4, 30), outcome.Request.PlotFinish);
    }

    /// <summary>
    /// Rows with no start date refuse rather than producing a chart over a range
    /// invented from the epoch.
    /// </summary>
    [Fact]
    public void No_event_with_a_date_refuses_rather_than_inventing_a_range()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create([], new Dictionary<string, string>(), StyleRegistry(), Grid());

        Assert.False(outcome.Succeeded);
        Assert.Equal(SceneBuildRequestRefusal.NoPlotRange, outcome.Refusal);
    }

    /// <summary>A single-day span is widened rather than refused as unreadable.</summary>
    [Fact]
    public void A_one_day_span_is_widened_to_two_days_rather_than_refused()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 3, 1), new DateOnly(2024, 3, 1))],
            new Dictionary<string, string> { ["RangePaddingDays"] = "0" },
            StyleRegistry(),
            Grid()
        );

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);
        Assert.True(outcome.Request!.PlotFinish > outcome.Request.PlotStart);
    }

    /// <summary>
    /// Explicit modes use the supplied dates verbatim: the request's plot range
    /// is the month-snapped explicit extent, ignoring the events entirely.
    /// Dates are the chart's dd/MM/yyyy stored form (exact plot dates, never
    /// the activity auto-parser equivalents).
    /// </summary>
    [Fact]
    public void Explicit_modes_resolve_the_request_range_from_the_supplied_dates()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 3, 1), new DateOnly(2024, 3, 31))],
            new Dictionary<string, string>
            {
                ["PlotStartMode"] = "Explicit",
                ["PlotFinishMode"] = "Explicit",
                ["PlotStartDate"] = "10/02/2024",
                ["PlotFinishDate"] = "10/04/2024",
            },
            StyleRegistry(),
            Grid()
        );

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);
        // Explicit dates pass through verbatim (no month-snap)
        Assert.Equal(new DateOnly(2024, 2, 10), outcome.Request!.PlotStart);
        Assert.Equal(new DateOnly(2024, 4, 10), outcome.Request.PlotFinish);
    }

    /// <summary>
    /// Mixed modes resolve each end from its own source: the explicit start is
    /// used verbatim while the finish still derives from the events.
    /// </summary>
    [Fact]
    public void A_mixed_mode_request_resolves_each_end_from_its_own_source()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 3, 1), new DateOnly(2024, 3, 31))],
            new Dictionary<string, string>
            {
                ["PlotStartMode"] = "Explicit",
                ["PlotFinishMode"] = "DataRange",
                ["PlotStartDate"] = "15/01/2024",
            },
            StyleRegistry(),
            Grid()
        );

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);
        // Explicit start passes through verbatim; DataRange finish is month-snapped
        Assert.Equal(new DateOnly(2024, 1, 15), outcome.Request!.PlotStart);
        Assert.Equal(new DateOnly(2024, 4, 30), outcome.Request.PlotFinish);
    }

    /// <summary>
    /// An unknown mode refuses as an invalid setting with the accepted values
    /// named, rather than falling back to the data range silently.
    /// </summary>
    [Fact]
    public void An_unknown_mode_refuses_as_an_invalid_setting()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 3, 1), new DateOnly(2024, 3, 31))],
            new Dictionary<string, string> { ["PlotStartMode"] = "Automatic" },
            StyleRegistry(),
            Grid()
        );

        Assert.False(outcome.Succeeded);
        Assert.Equal(SceneBuildRequestRefusal.InvalidSetting, outcome.Refusal);
        Assert.Contains("DataRange", outcome.Message!, StringComparison.Ordinal);
        Assert.Contains("Explicit", outcome.Message!, StringComparison.Ordinal);
    }

    /// <summary>
    /// Explicit mode with no date refuses as an invalid setting naming the
    /// date key, rather than rendering the data range the user overrode.
    /// </summary>
    [Fact]
    public void Explicit_mode_with_no_date_refuses_as_an_invalid_setting()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 3, 1), new DateOnly(2024, 3, 31))],
            new Dictionary<string, string> { ["PlotStartMode"] = "Explicit" },
            StyleRegistry(),
            Grid()
        );

        Assert.False(outcome.Succeeded);
        Assert.Equal(SceneBuildRequestRefusal.InvalidSetting, outcome.Refusal);
        Assert.Contains("PlotStartDate", outcome.Message!, StringComparison.Ordinal);
    }

    /// <summary>
    /// Explicit start after explicit finish refuses as an invalid setting.
    /// Uses dd/MM/yyyy values and asserts the StartAfterFinish path message.
    /// </summary>
    [Fact]
    public void An_explicit_start_after_the_explicit_finish_refuses()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 3, 1), new DateOnly(2024, 3, 31))],
            new Dictionary<string, string>
            {
                ["PlotStartMode"] = "Explicit",
                ["PlotFinishMode"] = "Explicit",
                ["PlotStartDate"] = "01/05/2024",
                ["PlotFinishDate"] = "01/04/2024",
            },
            StyleRegistry(),
            Grid()
        );

        Assert.False(outcome.Succeeded);
        Assert.Equal(SceneBuildRequestRefusal.InvalidSetting, outcome.Refusal);
        Assert.Contains("after the plot finish", outcome.Message!, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>A malformed numeric setting falls back rather than refusing the refresh.</summary>
    /// <summary>
    /// A metric-looking key in the settings map cannot corrupt a metric, because
    /// metrics are resolved from the catalogue and never from the settings map.
    /// </summary>
    /// <remarks>
    /// This test previously asserted that a malformed <c>GridLinePt</c> setting fell
    /// back to 0.5. That contract is withdrawn: the key was never read in the first
    /// place, so the test was passing over dead code and gave false assurance that the
    /// settings path was live. The behaviour now asserted is the real one — a garbage
    /// value under a metric name is ignored, and the catalogue default survives.
    /// </remarks>
    [Fact]
    public void A_malformed_metric_setting_is_ignored_rather_than_applied()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            new Dictionary<string, string> { ["GridLinePt"] = "not-a-number" },
            StyleRegistry(),
            Grid()
        );

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);
        Assert.Equal(GanttCatalogues.MetricDefault("GridLinePt"), outcome.Request!.GridLinePt);
    }

    /// <summary>
    /// Every metric the factory emits equals the code-owned catalogue default.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>This is the test that would have caught the dead-metric-token defect.</b> The
    /// factory previously read its metrics with <c>ReadDouble(settings, tokenName,
    /// fallback)</c>, but the settings map carries only the approved setting keys and
    /// no metric name is one of them, so every lookup missed and the divergent literal
    /// fallback was used instead — six of eight disagreed with the catalogue. Asserting
    /// the request against <see cref="GanttCatalogues.Metrics"/> states the contract
    /// that matters: the rendered chart uses the tokens the workbook publishes.
    /// </para>
    /// <para>
    /// It is written as a whole-catalogue sweep rather than as eight hand-picked
    /// assertions, so a token added to the request later is covered without anyone
    /// remembering to extend this list.
    /// </para>
    /// </remarks>
    [Fact]
    public void Every_emitted_metric_equals_the_catalogue_default()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            new Dictionary<string, string>(),
            StyleRegistry(),
            Grid()
        );

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);
        SceneBuildRequest request = outcome.Request!;

        (string Token, double Actual)[] emitted =
        [
            ("TitleBandHeightPt", request.TitleBandHeightPt),
            ("YearBandHeightPt", request.YearBandHeightPt),
            ("PeriodBandHeightPt", request.PeriodBandHeightPt),
            ("MinimumHeaderLabelWidthPt", request.MinimumHeaderLabelWidthPt),
            ("GridLinePt", request.GridLinePt),
            ("MajorBoundaryPt", request.MajorBoundaryPt),
            ("MilestoneSizePt", request.MilestoneSizePt),
            ("DelineatorLinePt", request.DelineatorLinePt),
            ("LabelGapPt", request.LabelGapPt),
            // The label box is one worksheet row tall (owner ruling), so the member is
            // fed by the row-height token rather than the retired LabelHeightPt.
            ("GanttRowHeightPt", request.RowHeightPt),
            ("LanePaddingTopPt", request.LaneMetrics!.LanePaddingTopPt),
            ("LanePaddingBottomPt", request.LaneMetrics.LanePaddingBottomPt),
            ("StackGapPt", request.LaneMetrics.StackGapPt),
            ("SplitterHeightPt", request.LaneMetrics.SplitterHeightPt),
            ("SpacerHeightPt", request.LaneMetrics.SpacerHeightPt),
        ];

        foreach ((string token, double actual) in emitted)
        {
            double expected = GanttCatalogues.MetricDefault(token);
            Assert.True(
                Math.Abs(expected - actual) < 1e-9,
                $"Metric '{token}' was emitted as {actual} but the catalogue default is {expected}."
            );
        }

        // ADR-0031 D1: the padding is per-side, so it cannot ride in the sweep above.
        // Its right side is the chrome token and is still a catalogue value; its top
        // and bottom are MEASURED padding rows and its left is deliberately zero, so
        // asserting a single number against a single token would assert something
        // that is no longer true of any side.
        Assert.Equal(GanttCatalogues.MetricDefault("ChartOuterPaddingPt"), request.ChartPadding.RightPt);
        Assert.Equal(0, request.ChartPadding.LeftPt);
    }

    /// <summary>
    /// A metric name is never read from the settings map, so editing a settings cell
    /// cannot move a metric the catalogue owns.
    /// </summary>
    /// <remarks>
    /// The defect was a lookup against the wrong table: the key was always absent, so
    /// the branch that consumed it was dead. This asserts the absence directly — the
    /// setting is supplied, and the emitted value is still the catalogue default. Under
    /// the old code this returned 0.5 from the fallback and would still have passed for
    /// <c>GridLinePt</c>, so the token chosen here is one whose old fallback differed
    /// from its catalogue default (<c>TitleBandHeightPt</c>: 14 against 24).
    /// </remarks>
    [Fact]
    public void A_metric_setting_is_not_read_from_the_settings_map()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        SceneBuildRequestOutcome outcome = factory.Create(
            [Event(new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19))],
            new Dictionary<string, string> { ["TitleBandHeightPt"] = "999", ["ChartPadding"] = "999" },
            StyleRegistry(),
            Grid()
        );

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);
        Assert.Equal(GanttCatalogues.MetricDefault("TitleBandHeightPt"), outcome.Request!.TitleBandHeightPt);

        // ADR-0031 D1: this used to assert one ChartOuterPaddingPt figure on every
        // side. The live profile now pads per side -- measured rows top and bottom,
        // nothing on the left -- so the token survives only on the right, and the left
        // is asserted as the deliberate zero the owner asked for.
        Assert.Equal(GanttCatalogues.MetricDefault("ChartOuterPaddingPt"), outcome.Request.ChartPadding.RightPt);
        Assert.Equal(0, outcome.Request.ChartPadding.LeftPt);
    }

    /// <summary>
    /// The catalogue accessor itself refuses an unknown token name rather than
    /// inventing a value.
    /// </summary>
    /// <remarks>
    /// The positive test for the validator: a bad token name must produce the error
    /// path. Without it, a typo in one of the factory's token constants would resolve
    /// to nothing and the metric would silently disappear from the chart.
    /// </remarks>
    [Fact]
    public void An_unknown_metric_token_is_refused_by_the_catalogue()
    {
        Assert.Throws<ArgumentException>(() => GanttCatalogues.MetricDefault("NoSuchTokenPt"));
        Assert.False(GanttCatalogues.IsMetricToken("NoSuchTokenPt"));
        Assert.False(GanttCatalogues.IsMetricToken(null));
        Assert.True(GanttCatalogues.IsMetricToken("TitleBandHeightPt"));
    }

    /// <summary>
    /// Every setting key the factory reads is a key the configuration reader can
    /// actually return.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>This is the test that would have caught the key-name mismatches.</b> The
    /// factory read <c>"DateFormat"</c> while the catalogue key is
    /// <c>DateDisplayFormat</c>, and read <c>"DelineatorStackGapPt"</c> while the
    /// metric token is <c>StackGapPt</c>. Both lookups always missed. Their tests
    /// passed because each injected the key straight into a hand-built dictionary,
    /// so the suite proved the factory honours a key it was handed — never that the
    /// workbook can hand it one.
    /// </para>
    /// <para>
    /// Asserting against <see cref="GanttCatalogues.Settings"/> closes that gap: a key
    /// that is not in the approved set is now a failing test rather than a silently
    /// dead branch. The <c>SizePreset</c> and <c>RangePaddingDays</c> keys were the
    /// two that were missing for a long time; they are now real catalogue settings,
    /// pinned by <see cref="The_size_preset_keys_are_now_catalogue_settings"/> rather
    /// than recorded as a known gap here.
    /// </para>
    /// </remarks>
    [Fact]
    public void Every_setting_key_the_factory_reads_exists_in_the_catalogue()
    {
        HashSet<string> approved = new(GanttCatalogues.Settings.Select(static setting => setting.Key), StringComparer.Ordinal);

        // The keys the factory reads as *settings*. Each was read by literal string in
        // the pre-fix code; DateFormat is listed because it is the name the factory
        // used and the catalogue does NOT carry it, so this assertion is what turns
        // the rename into a test failure rather than a silent behaviour change.
        string[] read =
        [
            "ChartTitle",
            "AlternateBanding",
            "ShowMinorGrid",
            "ShowMajorGrid",
            "TimeScale",
            "PeriodLabelFormat",
            "DateDisplayFormat",
        ];

        foreach (string key in read)
        {
            Assert.True(approved.Contains(key), $"The factory reads setting '{key}', which the catalogue does not define.");
        }

        // The retired spelling must be gone, or the guard above would still pass on a
        // rename that left the old name in place as a second, dead read.
        Assert.DoesNotContain("DateFormat", approved);
    }

    /// <summary>
    /// The two settings R4.7H and the plot-range padding need are now real catalogue
    /// keys, so the factory's reads of them can actually be satisfied.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This test previously asserted the <em>opposite</em>: that <c>SizePreset</c> and
    /// <c>RangePaddingDays</c> were absent from the catalogue, which recorded a known
    /// gap so it could not be closed silently. The owner's ruling permitted the schema
    /// bump that adding a key requires, so the gap is closed and the assertion is
    /// inverted — the keys must now be present, or the factory reads them from a map
    /// that can never supply them and the presets fall back silently again.
    /// </para>
    /// <para>
    /// Both defaults are asserted, not just membership: a present key carrying a value
    /// the factory cannot parse would pass a membership check and still be dead.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_size_preset_keys_are_now_catalogue_settings()
    {
        HashSet<string> approved = new(GanttCatalogues.Settings.Select(static setting => setting.Key), StringComparer.Ordinal);

        Assert.Contains("SizePreset", approved);
        Assert.Contains("RangePaddingDays", approved);

        // The stored defaults must be values the factory accepts: a preset key that
        // `SizePresets.ByKey` rejects would be rejected at render time on a freshly
        // initialised workbook.
        GanttSettingDefinition preset = GanttCatalogues.Settings.First(s => s.Key == "SizePreset");
        Assert.NotNull(SizePresets.ByKey(preset.DefaultValue));

        GanttSettingDefinition padding = GanttCatalogues.Settings.First(s => s.Key == "RangePaddingDays");
        Assert.True(
            int.TryParse(padding.DefaultValue, CultureInfo.InvariantCulture, out int days) && days >= 0,
            $"RangePaddingDays default '{padding.DefaultValue}' is not a non-negative integer."
        );
    }

    /// <summary>
    /// A live request carries a date-label style, so §23 start/finish dates are
    /// actually emitted (ADR-0032 D2).
    /// </summary>
    /// <remarks>
    /// <b>The regression this pins.</b> The factory never assigned
    /// <c>LabelStyle</c>, and <c>SceneBuilder</c> skips the ENTIRE date-label pass
    /// when it is null - with a bare <c>return</c>, not a warning. So every live
    /// activity bar rendered with no start or finish date and nothing in the scene
    /// recorded why. A delineator label kept appearing because it is built on a
    /// different path with its own metrics, which is precisely the reported symptom:
    /// "only delineator labels generate". Asserting the property is non-null is the
    /// narrowest statement of the fix; the test below asserts the consequence.
    /// </remarks>
    [Fact]
    public void The_live_request_carries_the_date_label_style()
    {
        SceneBuildRequestOutcome outcome = Create(Grid());

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);
        Assert.NotNull(outcome.Request!.LabelStyle);
    }

    /// <summary>
    /// The label style carries the font family and size the scene MEASURED the label
    /// text with, so the host renders the font the box was sized for.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>Found live, and my first diagnosis of it was wrong.</b> The reported symptom
    /// was that label text clipped and that the box did not hug the text. I first
    /// concluded the width formula was at fault and that the box could not be too
    /// narrow because <c>LabelPlanner</c> sizes it with
    /// <c>Math.Min(measured.WidthPt, width)</c>. That reasoning only holds if the
    /// MEASURED font is the RENDERED font, and it was not: <c>AptosTextMetrics</c>
    /// measured at 8pt while this style carried no <c>fontSizePt</c>, so
    /// <c>ExcelShapeWriter</c>'s <c>if (request.FontSizePt is { } size)</c> never ran,
    /// the host rendered its own default, and the box sized from the 8pt measurement
    /// could not contain the text actually drawn.
    /// </para>
    /// <para>
    /// The assertion is that both values equal the catalogue token, not that they are
    /// 8pt. A test pinning the literal would keep passing if the token and the
    /// measuring seam drifted apart again, which is the defect this one exists to stop.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_live_label_style_carries_the_measured_font_family_and_size()
    {
        SceneBuildRequestOutcome outcome = Create(Grid());

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);
        SceneStyle? labelStyle = outcome.Request!.LabelStyle;

        Assert.NotNull(labelStyle);
        Assert.Equal(GanttCatalogues.LabelFontFamily, labelStyle!.FontFamily);
        Assert.Equal(GanttCatalogues.LabelBodyFontSizePt, labelStyle.FontSizePt);
    }

    /// <summary>
    /// The measuring seam and the emitting style resolve the SAME size, which is the
    /// invariant the clipping defect actually broke.
    /// </summary>
    /// <remarks>
    /// The counterweight to the property assertion above, and the test that would have
    /// caught the defect at its source rather than at the style. Either value alone
    /// could be correct while the other drifted: a style carrying 11pt with an 8pt
    /// measuring seam still under-measures, and an 11pt seam with a null style still
    /// renders at the host default. Only the equality states the contract the renderer
    /// depends on — the box is sized from the same font the host draws.
    /// </remarks>
    [Fact]
    public void The_measuring_seam_and_the_label_style_agree_on_the_font_size()
    {
        var metrics = new AptosTextMetrics();

        SceneBuildRequestOutcome outcome = Create(Grid());

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);

        double configuredFromSeam = metrics.FontSizePt;

        Assert.Equal(GanttCatalogues.LabelBodyFontSizePt, configuredFromSeam, precision: 10);
        Assert.Equal(outcome.Request!.LabelStyle!.FontSizePt!.Value, configuredFromSeam, precision: 10);
    }

    /// <summary>
    /// An unknown typography token is refused rather than defaulted, and a malformed
    /// point size is refused rather than becoming a zero-width box.
    /// </summary>
    /// <remarks>
    /// The positive cases for the two new validators in
    /// <see cref="GanttCatalogues.TypographyDefault"/> and
    /// <see cref="GanttCatalogues.TypographyDefaultText"/>, in the same commit as the
    /// validators. A defaulted value here would be a second authority that disagrees
    /// with the catalogue, which is the failure mode the whole change removes.
    /// </remarks>
    [Fact]
    public void An_unknown_or_malformed_typography_token_is_refused_rather_than_defaulted()
    {
        Assert.Throws<ArgumentNullException>(() => GanttCatalogues.TypographyDefault(null!));
        Assert.Throws<ArgumentException>(() => GanttCatalogues.TypographyDefault("NoSuchTokenPt"));

        // The catalogue's own values parse, which is what makes the two refusals above
        // refusals rather than a parser that rejects everything.
        Assert.Equal(8d, GanttCatalogues.LabelBodyFontSizePt);
        Assert.Equal("Aptos", GanttCatalogues.LabelFontFamily);
    }

    /// <summary>
    /// A built live scene emits the §23 start/finish date labels and reports no
    /// <c>DateLabelRefused</c> warning.
    /// </summary>
    /// <remarks>
    /// The counterweight to the property assertion above. <c>LabelStyle</c> could be
    /// non-null and still not produce labels - the builder also needs a measurable
    /// text metric, a resolvable position, and a style that survives the Inside/Outside
    /// split - so this asserts the OBSERVABLE outcome rather than the input. It is
    /// what would have caught the original defect, and it fails if a future change
    /// nulls the style again somewhere downstream of the factory.
    /// </remarks>
    [Fact]
    public void A_built_live_scene_emits_date_labels_for_an_activity()
    {
        SceneBuildRequestOutcome outcome = Create(Grid());

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);
        SceneBuildOutcome built = SceneBuilder.TryBuild(outcome.Request!);

        Assert.True(built.Succeeded, "scene refused: " + built.Refusal);
        Assert.DoesNotContain(built.Result!.Scene.Warnings, w => w.Code == "DateLabelRefused");

        // Count only THIS ROW's date labels, not every text primitive in the scene.
        // The month and year band labels are SceneText too, so a scene that emitted
        // ZERO date labels would still satisfy ">= 2". A mutation that removed the
        // LabelStyle assignment proved exactly that: the property test failed and
        // this one passed. Chart-owned band labels are excluded by ownership kind.
        int dateLabels = built.Result.Scene.Primitives.OfType<SceneText>().Count(text => text.OwnerId.Kind == SceneOwnerKind.Row);

        Assert.True(
            dateLabels >= 2,
            $"Expected a start and a finish date label, but the scene emitted {dateLabels} row-owned text primitive(s)."
        );
    }

    /// <summary>A null collaborator is refused rather than dereferenced.</summary>
    [Fact]
    public void Null_inputs_are_refused()
    {
        ExcelSceneBuildRequestFactory factory = new(Metrics());

        Assert.Throws<ArgumentNullException>(() => factory.Create(null!, new Dictionary<string, string>(), StyleRegistry(), Grid()));
    }

    /// <summary>A registry with the one style the fixture events reference.</summary>
    private static GanttStyleRegistry StyleRegistry() =>
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
                ActivityHeightPt: 12
            ),
        ]);
}
