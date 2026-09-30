using System.Globalization;
using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

public sealed class PanelBuilderTests
{
    private const double _plotLeft = 200.0;
    private const double _plotY = 66.0;
    private const double _yearBandHeight = 20.0;
    private const double _headerHeight = 18.0;
    private const double _rowHeight = 12.0;

    private static readonly string[] _cultures = ["en-GB", "en-US"];

    private static readonly PanelTheme _theme = new(
        new SceneStyle("DataPanelFill"),
        new SceneStyle("DefaultText"),
        new SceneStyle("HeaderFill"),
        new SceneStyle("DefaultText", fontSizePt: 9.0, bold: true),
        new SceneStyle("GridLine", strokeColour: ColourHex.Parse("#BFBFBF"), outlineWidthPt: 0.5));

    [Fact]
    public void The_header_carries_the_schema_display_names_in_order_and_nothing_else()
    {
        // The header set equals the schema's, so the emitted headers cannot drift
        // from the workbook contract in either direction. This assertion was
        // formerly a *negative* "no Duration header" pin, written when Duration did
        // not exist; R4.7C added it to the schema, so the pin had to invert with
        // the contract rather than be deleted (ADR-0029 D1).
        string[] expected = [.. GanttTableSchema.Default.Columns.Select(column => column.Name)];

        SceneText[] headers = [.. Build().Result!.Primitives.OfType<SceneText>()
            .Where(text => text.PrimitiveId.Contains("header-text:", StringComparison.Ordinal))];

        Assert.Equal(expected, headers.Select(text => text.Text).ToArray());
        // R4.7C: Duration is a visible schema column, so its header must appear.
        Assert.Contains(headers, text => text.Text == "Duration");
        // Sort/filter arrows and filter funnels are not reproduced (section 4).
        Assert.All(headers, text => Assert.DoesNotContain(text.Text, IsSortOrFilterGlyph));
        Assert.Equal(GanttTableSchema.Default.Columns.Count, headers.Length);
    }

    [Fact]
    public void Description_is_left_aligned_and_every_other_header_is_centred()
    {
        SceneText[] headers = [.. Build().Result!.Primitives.OfType<SceneText>()
            .Where(text => text.PrimitiveId.Contains("header-text:", StringComparison.Ordinal))];

        SceneText description = Assert.Single(headers, text => text.Text == "Description");
        Assert.Equal(GanttTextAlignment.Left, description.Alignment);
        Assert.All(
            headers.Where(text => text.Text != "Description"),
            text => Assert.Equal(GanttTextAlignment.Centre, text.Alignment));
    }

    [Fact]
    public void The_header_bottom_lands_exactly_on_the_period_header_bottom()
    {
        // Section 4: the header height aligns exactly with the period-header
        // bottom. R3.5 places the period band at PlotBounds.Y - YearBandHeightPt,
        // so that is the value the caller must pass; this asserts the panel lands
        // on it rather than near it.
        PanelBuildResult result = Build().Result!;
        var periodHeaderBottom = _plotY - _yearBandHeight;

        Assert.Equal(periodHeaderBottom, result.HeaderRowBounds.Bottom);
        Assert.Equal(periodHeaderBottom - _headerHeight, result.HeaderRowBounds.Y);
        Assert.All(
            result.Primitives.OfType<SceneRect>().Where(rect => rect.PrimitiveId.Contains("header-cell:", StringComparison.Ordinal)),
            rect => Assert.Equal(periodHeaderBottom, rect.Bounds.Bottom));
    }

    [Fact]
    public void The_panel_right_edge_touches_the_plot_left_edge_with_no_gap_or_overlap()
    {
        PanelBuildResult result = Build().Result!;

        Assert.Equal(_plotLeft, result.PanelBounds.Right);
        Assert.Equal(_plotLeft, result.HeaderRowBounds.Right);
        Assert.All(
            result.Primitives.OfType<SceneRect>().Where(rect => rect.PrimitiveId.Contains("header-cell:", StringComparison.Ordinal)),
            rect => Assert.True(
                rect.Bounds.Right <= _plotLeft,
                $"header cell {rect.PrimitiveId} overlapped the plot by {rect.Bounds.Right - _plotLeft}pt."));
    }

    [Fact]
    public void Every_cell_lands_on_the_exact_measured_grid()
    {
        // The measured grid is the only geometry source (D2), so a cell's bounds
        // must be the cumulative column widths, not a re-measured approximation.
        PanelCellGrid grid = Grid();
        PanelBuildResult result = Build().Result!;

        var expectedLeft = _plotLeft - grid.TotalWidthPt;
        for (var index = 0; index < grid.Columns.Count; index++)
        {
            PanelColumn column = grid.Columns[index];
            SceneRect header = result.Primitives.OfType<SceneRect>()
                .Single(rect => rect.PrimitiveId.EndsWith($"header-cell:{column.Name}", StringComparison.Ordinal));
            Assert.Equal(new RectD(expectedLeft, _plotY - _yearBandHeight - _headerHeight, column.WidthPt, _headerHeight), header.Bounds);
            expectedLeft += column.WidthPt;
        }

        Assert.Equal(_plotLeft, expectedLeft);
    }

    [Fact]
    public void A_body_row_lands_on_the_measured_row_height_below_the_header()
    {
        PanelBuildResult result = Build().Result!;
        SceneRect[] firstColumn =
        [
            .. result.Primitives.OfType<SceneRect>()
                .Where(rect => rect.PrimitiveId.EndsWith(":panel-cell:Id", StringComparison.Ordinal)),
        ];

        // Row one starts on the header's bottom edge; row two is exactly one
        // measured row height below it, so the panel tracks the lanes.
        Assert.Equal(_plotY - _yearBandHeight, firstColumn[0].Bounds.Y);
        Assert.Equal(_rowHeight, firstColumn[0].Bounds.Height);
        Assert.Equal(firstColumn[0].Bounds.Bottom, firstColumn[1].Bounds.Y);
    }
    [Fact]
    public void A_mixed_height_body_lays_out_on_the_exact_cumulative_positions()
    {
        // The regression this change exists for. With 12/24/15pt rows the cell tops
        // must be bodyTop, bodyTop+12, bodyTop+36, and the panel bottom bodyTop+51.
        // A single sample height would either refuse this table or place every row
        // at the wrong edge, and a tolerance here would accept exactly that, so the
        // positions are asserted exactly.
        PanelRow[] rows = [Row(1, "first"), Row(2, "second"), Row(3, "third")];

        PanelBuildOutcome outcome = BuildWithHeights(rows, 12.0, 24.0, 15.0);

        Assert.True(outcome.Succeeded, "Panel build refused: " + outcome.Refusal);
        double bodyTop = _plotY - _yearBandHeight;
        double[] expectedTops = [bodyTop, bodyTop + 12.0, bodyTop + 36.0];
        double[] expectedHeights = [12.0, 24.0, 15.0];

        for (var index = 0; index < rows.Length; index++)
        {
            GanttRowId id = rows[index].RowId;
            SceneRect cell = outcome.Result!.Primitives.OfType<SceneRect>()
                .Single(rect => rect.PrimitiveId == ScenePrimitive.CreateId(SceneOwnerId.ForRow(id), "panel-cell:Id"));
            Assert.Equal(new RectD(cell.Bounds.X, expectedTops[index], cell.Bounds.Width, expectedHeights[index]), cell.Bounds);
        }

        // The panel bottom is the exact sum of the measured heights plus the header.
        Assert.Equal(bodyTop + 51.0, outcome.Result!.PanelBounds.Bottom);
    }

    [Fact]
    public void Every_row_border_lands_on_its_own_rows_cumulative_edge()
    {
        // Borders use the same running sum as the cells, so a border can never frame
        // a different edge than the cell it is supposed to frame.
        PanelRow[] rows = [Row(1, "first"), Row(2, "second"), Row(3, "third")];

        PanelBuildOutcome outcome = BuildWithHeights(rows, 12.0, 24.0, 15.0);

        Assert.True(outcome.Succeeded, "Panel build refused: " + outcome.Refusal);
        double bodyTop = _plotY - _yearBandHeight;
        SceneLine[] horizontals = [.. outcome.Result!.Primitives.OfType<SceneLine>()
            .Where(line => line.PrimitiveId.Contains("panel-border-h:", StringComparison.Ordinal))];
        // Panel top, header/body boundary, then one bottom edge per measured row:
        // 28 (top), 46 (boundary), 58, 82, 97 for 12/24/15pt rows.
        Assert.Equal(
            [bodyTop - _headerHeight, bodyTop, bodyTop + 12.0, bodyTop + 36.0, bodyTop + 51.0],
            [.. horizontals.Select(line => line.From.Y)]);
    }

    [Fact]
    public void A_row_count_that_disagrees_with_the_measured_heights_is_refused()
    {
        // Positional alignment is the whole contract, so a source row that failed
        // validation - leaving more rows than heights, or fewer - must refuse rather
        // than shift every cell below the divergence into the wrong row.
        PanelRow[] threeRows = [Row(1, "first"), Row(2, "second"), Row(3, "third")];

        Assert.Equal(
            PanelBuildRefusal.RowCountMismatch,
            PanelBuilder.TryBuild(Request(threeRows, GridWith([12.0, 24.0]))).Refusal);
        Assert.Equal(
            PanelBuildRefusal.RowCountMismatch,
            PanelBuilder.TryBuild(Request(threeRows, GridWith([12.0, 24.0, 15.0, 9.0]))).Refusal);
    }

    [Fact]
    public void Adjacent_cells_share_one_border_and_none_is_doubled()
    {
        // Shared borders, not per-cell borders: with N columns there are N+1
        // vertical edges. Horizontally the edges are the panel top, the
        // header/body boundary, and each row's bottom, so N rows give N+2. A
        // per-cell implementation would emit N*(N+1) lines and a visibly
        // doubled grid.
        PanelBuildResult result = Build().Result!;
        SceneLine[] verticals = [.. result.Primitives.OfType<SceneLine>()
            .Where(line => line.PrimitiveId.Contains("panel-border-v:", StringComparison.Ordinal))];
        SceneLine[] horizontals = [.. result.Primitives.OfType<SceneLine>()
            .Where(line => line.PrimitiveId.Contains("panel-border-h:", StringComparison.Ordinal))];

        var columns = Grid().Columns.Count;
        var rows = TwoRows().Length;

        Assert.Equal(columns + 1, verticals.Length);
        Assert.Equal(rows + 2, horizontals.Length);
        Assert.Equal(verticals.Length, verticals.Select(line => line.PrimitiveId).Distinct(StringComparer.Ordinal).Count());
        Assert.Equal(horizontals.Length, horizontals.Select(line => line.PrimitiveId).Distinct(StringComparer.Ordinal).Count());
        Assert.Equal(
            verticals.Select(line => (line.From.X, line.To.X)).Distinct(),
            verticals.Select(line => (line.From.X, line.To.X)));
    }

    [Fact]
    public void The_header_and_body_are_separated_by_their_shared_border()
    {
        // Section 3 requires a border at every cell edge, and the header's
        // bottom edge is the first row's top edge. Without this border the
        // header and the body grid run together and the panel has no top edge
        // at bodyTop at all -- the old code drew only the panel top and the row
        // bottoms, so the boundary was missing entirely.
        PanelBuildResult result = Build().Result!;
        SceneLine boundary = Assert.Single(
            result.Primitives.OfType<SceneLine>(),
            line => line.PrimitiveId.EndsWith("panel-border-h:1", StringComparison.Ordinal));

        Assert.Equal(result.HeaderRowBounds.Bottom, boundary.From.Y);
        Assert.Equal(boundary.From.Y, boundary.To.Y);
    }

    [Fact]
    public void A_blank_cell_emits_its_background_but_no_text()
    {
        // A blank cell is legal data (R2.5 U2); the background must still render
        // so the grid has no hole, but an empty text primitive is noise.
        PanelBuildResult result = Build(blanks: ["ParentId", "Description"]).Result!;

        SceneRect[] parentCells =
        [
            .. result.Primitives.OfType<SceneRect>()
                .Where(rect => rect.PrimitiveId.EndsWith(":panel-cell:ParentId", StringComparison.Ordinal)),
        ];
        SceneText[] parentTexts =
        [
            .. result.Primitives.OfType<SceneText>()
                .Where(text => text.PrimitiveId.EndsWith(":panel-text:ParentId", StringComparison.Ordinal)),
        ];

        Assert.Equal(2, parentCells.Length);
        Assert.Empty(parentTexts);
    }

    [Fact]
    public void A_zero_row_panel_is_now_unrepresentable_rather_than_a_header_with_no_body()
    {
        // This test used to assert that a row-less panel still closed itself neatly.
        // It no longer can be built: the measured grid requires at least one body row
        // height, and the panel requires one row per measured height. That is the
        // better outcome - a header with no body reads as a data table that lost its
        // data, which is silent data loss, and `tblGanttData` always has a body. The
        // guard is pinned here at the grid, which is where the emptiness is refused.
        Assert.Equal(
            PanelCellGridRefusal.NoRows,
            PanelCellGrid.TryCreate([new PanelColumn("Id", 40)], [], 18.0, ["Id"]).Refusal);
    }

    [Fact]
    public void Date_cells_carry_the_approved_invariant_format()
    {
        // ADR-0016: dd/mm/yyyy, culture-invariant. The builder emits the text it
        // is given, so this pins the composed result rather than a re-format.
        CultureInfo original = CultureInfo.CurrentCulture;
        try
        {
            string[] observed = [.. _cultures.Select(culture =>
            {
                CultureInfo.CurrentCulture = new CultureInfo(culture);
                SceneText start = Build().Result!.Primitives.OfType<SceneText>()
                    .First(text => text.PrimitiveId.EndsWith(":panel-text:Start", StringComparison.Ordinal));
                return start.Text;
            })];

            Assert.All(observed, text => Assert.Equal("05/01/2024", text));
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }

    [Fact]
    public void Body_text_follows_the_measured_column_alignment()
    {
        PanelBuildResult result = Build().Result!;
        SceneText[] typeCells =
        [
            .. result.Primitives.OfType<SceneText>()
                .Where(text => text.PrimitiveId.EndsWith(":panel-text:Type", StringComparison.Ordinal)),
        ];

        Assert.NotEmpty(typeCells);
        Assert.All(typeCells, text => Assert.Equal(GanttTextAlignment.Centre, text.Alignment));

        // The rest keep the default Left, so the grid's measured alignment is
        // genuinely driving the body text rather than a hard-coded Centre.
        SceneText[] idCells =
        [
            .. result.Primitives.OfType<SceneText>()
                .Where(text => text.PrimitiveId.EndsWith(":panel-text:Id", StringComparison.Ordinal)),
        ];
        Assert.NotEmpty(idCells);
        Assert.All(idCells, text => Assert.Equal(GanttTextAlignment.Left, text.Alignment));
    }

    [Fact]
    public void Building_the_same_request_three_times_produces_an_identical_sequence()
    {
        // Determinism: nothing may depend on hash or collection iteration order.
        PanelBuildRequest request = Request(TwoRows());
        string[][] runs =
        [
            .. Enumerable.Range(0, 3).Select(_ => PanelBuilder.TryBuild(request).Result!.Primitives.Select(p => p.PrimitiveId).ToArray()),
        ];

        Assert.Equal(runs[0], runs[1]);
        Assert.Equal(runs[0], runs[2]);
    }
    [Fact]
    public void A_null_request_is_refused() =>
        Assert.Equal(PanelBuildRefusal.NullRequest, PanelBuilder.TryBuild(null).Refusal);

    [Fact]
    public void A_null_grid_is_refused() =>
        Assert.Equal(
            PanelBuildRefusal.NullGrid,
            PanelBuilder.TryBuild(Request(TwoRows()) with { Grid = null! }).Refusal);

    [Fact]
    public void A_null_row_list_is_refused() =>
        Assert.Equal(
            PanelBuildRefusal.NullRows,
            PanelBuilder.TryBuild(Request(TwoRows()) with { Rows = null! }).Refusal);

    [Fact]
    public void A_null_theme_is_refused() =>
        Assert.Equal(
            PanelBuildRefusal.NullTheme,
            PanelBuilder.TryBuild(Request(TwoRows()) with { Theme = null! }).Refusal);

    // RectD validates its own extents, so negative and non-finite widths cannot be
    // constructed; a zero-extent rectangle is the only reachable bad input.
    [Theory]
    [InlineData(0.0, 134.0)]
    [InlineData(400.0, 0.0)]
    public void Invalid_plot_bounds_are_refused(double width, double height) =>
        Assert.Equal(
            PanelBuildRefusal.InvalidPlotBounds,
            PanelBuilder.TryBuild(Request(TwoRows()) with { PlotBounds = new RectD(_plotLeft, _plotY, width, height) }).Refusal);

    [Fact]
    public void A_default_plot_bounds_is_refused() =>
        Assert.Equal(
            PanelBuildRefusal.InvalidPlotBounds,
            PanelBuilder.TryBuild(Request(TwoRows()) with { PlotBounds = default }).Refusal);

    [Fact]
    public void A_non_finite_header_bottom_is_refused() =>
        Assert.Equal(
            PanelBuildRefusal.NonFiniteHeaderBottom,
            PanelBuilder.TryBuild(Request(TwoRows()) with { HeaderBottomPt = double.NaN }).Refusal);

    [Fact]
    public void A_null_row_is_refused() =>
        Assert.Equal(
            PanelBuildRefusal.NullRow,
            PanelBuilder.TryBuild(Request([null!])).Refusal);

    [Fact]
    public void A_row_whose_cell_count_differs_from_the_grid_is_refused() =>
        Assert.Equal(
            PanelBuildRefusal.CellCountMismatch,
            PanelBuilder.TryBuild(Request([new PanelRow(GanttRowId.New(), ["only one"])])).Refusal);

    [Fact]
    public void Two_rows_sharing_an_identity_are_refused()
    {
        // A repeated row id would emit two primitives with the same ID, which a
        // renderer cannot disambiguate.
        GanttRowId shared = GanttRowId.New();
        Assert.Equal(
            PanelBuildRefusal.DuplicateRow,
            PanelBuilder.TryBuild(Request([new PanelRow(shared, Cells()), new PanelRow(shared, Cells())])).Refusal);
    }

    /// <summary>
    /// Recognises the arrow and funnel glyphs an Excel Table would draw for
    /// sort and filter. Section 4 does not reproduce them.
    /// </summary>
    private static bool IsSortOrFilterGlyph(char candidate) =>
        candidate is '\u25B2' or '\u25B3' or '\u25BC' or '\u25BD' or '\u25B6' or '\u25A0' or '\u25AC';

    private static PanelBuildOutcome Build(
        IReadOnlyList<PanelRow>? rows = null,
        IReadOnlyList<string>? blanks = null) =>
        PanelBuilder.TryBuild(Request(rows ?? TwoRows(blanks)));

    private static PanelBuildOutcome BuildWithHeights(
        IReadOnlyList<PanelRow> rows,
        params double[] rowHeightsPt) =>
        PanelBuilder.TryBuild(Request(rows, GridWith(rowHeightsPt)));

    private static PanelBuildRequest Request(IReadOnlyList<PanelRow> rows, PanelCellGrid? grid = null) =>
        new(
            grid ?? GridFor(rows.Count),
            rows,
            new RectD(_plotLeft, _plotY, 400.0, 134.0),
            _plotY - _yearBandHeight,
            _theme);

    /// <summary>Builds a uniform grid carrying exactly one height per supplied row.</summary>
    private static PanelCellGrid GridFor(int rowCount) =>
        GridWith([.. Enumerable.Repeat(_rowHeight, rowCount)]);

    private static PanelCellGrid Grid(params double[] rowHeightsPt) =>
        GridWith([.. (rowHeightsPt.Length == 0 ? [_rowHeight] : rowHeightsPt)]);

    private static PanelCellGrid GridWith(IReadOnlyList<double> rowHeightsPt)
    {
        // Widths vary so a cumulative-sum error cannot pass, and one column is
        // given a non-default alignment so the body-alignment pin has something
        // to distinguish rather than reading the enum's default back.
        PanelColumn[] columns =
        [
            .. GanttTableSchema.Default.Columns.Select((column, index) =>
                new PanelColumn(
                    column.Name,
                    10.0 + index,
                    string.Equals(column.Name, "Type", StringComparison.Ordinal)
                        ? GanttTextAlignment.Centre
                        : GanttTextAlignment.Left)),
        ];
        return PanelCellGrid
            .TryCreate(columns, rowHeightsPt, _headerHeight, [.. GanttTableSchema.Default.Columns.Where(c => c.IsRequired).Select(c => c.Name)])
            .Grid!;
    }

    private static string?[] Cells(IReadOnlyList<string>? blankColumns = null)
    {
        var cells = new string?[GanttTableSchema.Default.Columns.Count];
        Array.Fill(cells, "x");
        cells[SchemaIndex("Start")] = GanttDateFormatting.FormatDdMMyyyy(new DateOnly(2024, 1, 5));
        cells[SchemaIndex("Finish")] = GanttDateFormatting.FormatDdMMyyyy(new DateOnly(2024, 1, 9));

        // Blanking is keyed by column name, not index, so a schema change cannot
        // silently redirect a fixture at a different column.
        foreach (var column in blankColumns ?? [])
        {
            cells[SchemaIndex(column)] = "  ";
        }

        return cells;
    }

    private static int SchemaIndex(string name)
    {
        for (var index = 0; index < GanttTableSchema.Default.Columns.Count; index++)
        {
            if (string.Equals(GanttTableSchema.Default.Columns[index].Name, name, StringComparison.Ordinal))
            {
                return index;
            }
        }

        throw new InvalidOperationException($"'{name}' is not a schema column.");
    }

    private static PanelRow[] TwoRows(IReadOnlyList<string>? blanks = null) =>
    [
        new PanelRow(GanttRowId.New(), Cells(blanks)),
        new PanelRow(GanttRowId.New(), Cells(blanks)),
    ];

    /// <summary>One body row with distinct identity and the standard cell fixture.</summary>
    private static PanelRow Row(int index, string description) =>
        new(GanttRowId.New(), [.. Cells().Select((cell, columnIndex) => columnIndex == SchemaIndex("Description") ? description : cell)]);
}
