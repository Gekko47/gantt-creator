namespace GanttCreator.Core.Scene;

/// <summary>One data-panel body row: its row identity and its cell texts.</summary>
/// <param name="RowId">The row's stable identity, which owns every primitive it emits.</param>
/// <param name="Cells">
/// The cell texts in <see cref="PanelBuildRequest.Grid"/> column order. A
/// <see langword="null"/> or blank entry emits the background rectangle but no
/// text, because a blank cell is legal data (R2.5 U2) and an empty text
/// primitive would be noise a renderer has to special-case.
/// </param>
public sealed record PanelRow(GanttRowId RowId, IReadOnlyList<string?> Cells);

/// <summary>The resolved export style set for the data panel and its header.</summary>
/// <param name="BodyFill">The cell background fill (section 3 <c>DataPanelFill</c>).</param>
/// <param name="BodyText">The cell text style (section 3 <c>DefaultText</c>).</param>
/// <param name="HeaderFill">The header band fill (section 4 <c>HeaderFill</c>).</param>
/// <param name="HeaderText">The header text style (section 4 <c>HeaderFontSizePt</c>, bold).</param>
/// <param name="Border">The shared cell border style (section 3 resolved border token).</param>
public sealed record PanelTheme(
    SceneStyle BodyFill,
    SceneStyle BodyText,
    SceneStyle HeaderFill,
    SceneStyle HeaderText,
    SceneStyle Border);

/// <summary>The typed data-panel build request.</summary>
/// <param name="Grid">The measured cell grid. The builder never measures (R3.11 D2).</param>
/// <param name="Rows">The body rows to reproduce, in display order.</param>
/// <param name="PlotBounds">
/// The plot bounds. Section 3 requires the panel's right edge to touch this
/// rectangle's left edge with zero gap and zero overlap, so the panel is
/// positioned from that value rather than from its own width.
/// </param>
/// <param name="HeaderBottomPt">
/// The bottom edge of the period header, in points. Section 4 requires the
/// data-panel header's bottom to align exactly with it. The builder cannot
/// derive this: R3.5 places the period band at <c>PlotBounds.Y - YearBandHeightPt</c>,
/// and the band height is a Phase 4 concern. The caller passes the R3.5 value
/// and the header band is placed upward from it, which makes the alignment
/// structural rather than a coincidence.
/// </param>
/// <param name="HeaderHeightPt">The header band height in points.</param>
/// <param name="Theme">The resolved export styles.</param>
public sealed record PanelBuildRequest(
    PanelCellGrid Grid,
    IReadOnlyList<PanelRow> Rows,
    RectD PlotBounds,
    double HeaderBottomPt,
    double HeaderHeightPt,
    PanelTheme Theme);

/// <summary>The reason a data-panel build was refused.</summary>
public enum PanelBuildRefusal
{
    /// <summary>The request was null.</summary>
    NullRequest = 0,

    /// <summary>The measured grid was null.</summary>
    NullGrid = 1,

    /// <summary>The row list was null.</summary>
    NullRows = 2,

    /// <summary>The theme was null.</summary>
    NullTheme = 3,

    /// <summary>The plot bounds were not finite, or had a non-positive size.</summary>
    InvalidPlotBounds = 4,

    /// <summary>The header bottom was not finite.</summary>
    NonFiniteHeaderBottom = 5,

    /// <summary>The header height was not finite, or was not positive.</summary>
    NonPositiveHeaderHeight = 6,

    /// <summary>A row was null.</summary>
    NullRow = 7,

    /// <summary>A row's cell count did not match the grid's column count.</summary>
    CellCountMismatch = 8,

    /// <summary>Two rows shared a row identity, which would collide primitive IDs.</summary>
    DuplicateRow = 9,
}

/// <summary>The resolved data-panel primitives and their extents.</summary>
/// <param name="Primitives">The emitted primitives in deterministic order.</param>
/// <param name="PanelBounds">
/// The full panel including its header, whose right edge equals the left edge
/// of <see cref="PanelBuildRequest.PlotBounds"/>.
/// </param>
/// <param name="HeaderRowBounds">The header band's exact bounds.</param>
public sealed record PanelBuildResult(
    IReadOnlyList<ScenePrimitive> Primitives,
    RectD PanelBounds,
    RectD HeaderRowBounds);

/// <summary>The typed result of building the data panel.</summary>
/// <param name="Result">The resolved panel, or <see langword="null"/>.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/>.</param>
public sealed record PanelBuildOutcome(PanelBuildResult? Result, PanelBuildRefusal? Refusal)
{
    /// <summary>Gets whether the panel was built.</summary>
    public bool Succeeded => Result is not null;
}

/// <summary>
/// Builds the section 3 data panel and its section 4 header as export-ready scene
/// primitives.
/// </summary>
/// <remarks>
/// <para>
/// The builder measures nothing (R3.11 D2). <see cref="PanelCellGrid"/> arrives
/// already measured by a Phase 4/5 adapter, and this type only projects those
/// exact point widths into primitive bounds.
/// </para>
/// <para>
/// Z-order follows the guide's table: the header band and the shared cell
/// borders sit at <see cref="ZLayer.Frame"/> (80, "table borders, period
/// headers"), cell backgrounds sit at <see cref="ZLayer.Background"/> (0), and
/// cell text sits at <see cref="ZLayer.Label"/> (70) so it is above its own
/// background and below the borders that frame it. The header fill and its text
/// share layer 80, matching how R3.5 already emits a period band and its label.
/// </para>
/// </remarks>
public static class PanelBuilder
{
    /// <summary>The one schema column section 4 aligns left; every other header centres.</summary>
    private const string _descriptionColumn = "Description";

    /// <summary>Attempts to build the data panel and its header.</summary>
    /// <param name="request">The typed panel build request.</param>
    /// <returns>A typed result or refusal.</returns>
    public static PanelBuildOutcome TryBuild(PanelBuildRequest? request)
    {
        if (request is null)
        {
            return Refused(PanelBuildRefusal.NullRequest);
        }

        if (request.Grid is null)
        {
            return Refused(PanelBuildRefusal.NullGrid);
        }

        if (request.Rows is null)
        {
            return Refused(PanelBuildRefusal.NullRows);
        }

        if (request.Theme is null)
        {
            return Refused(PanelBuildRefusal.NullTheme);
        }

        RectD plot = request.PlotBounds;

        // RectD validates its own extents in its constructor, so a negative or
        // non-finite width can never reach here. Only a zero-extent rectangle
        // (for example `default`) is constructible, and a panel beside a
        // zero-width plot has no edge to touch, so that case is refused.
        if (plot.Width <= 0 || plot.Height <= 0)
        {
            return Refused(PanelBuildRefusal.InvalidPlotBounds);
        }

        if (!double.IsFinite(request.HeaderBottomPt))
        {
            return Refused(PanelBuildRefusal.NonFiniteHeaderBottom);
        }

        if (!double.IsFinite(request.HeaderHeightPt) || request.HeaderHeightPt <= 0)
        {
            return Refused(PanelBuildRefusal.NonPositiveHeaderHeight);
        }

        PanelCellGrid grid = request.Grid;
        HashSet<GanttRowId> seen = [];
        foreach (PanelRow row in request.Rows)
        {
            if (row?.RowId is null)
            {
                return Refused(PanelBuildRefusal.NullRow);
            }

            if (row.Cells is null || row.Cells.Count != grid.Columns.Count)
            {
                return Refused(PanelBuildRefusal.CellCountMismatch);
            }

            if (!seen.Add(row.RowId))
            {
                return Refused(PanelBuildRefusal.DuplicateRow);
            }
        }

        // Section 3: the panel's right edge touches the plot's left edge. Deriving
        // the panel origin from PlotBounds.Left rather than from the panel's own
        // width makes a gap or an overlap unrepresentable.
        var panelLeft = plot.Left - grid.TotalWidthPt;
        var headerTop = request.HeaderBottomPt - request.HeaderHeightPt;
        var bodyTop = request.HeaderBottomPt;
        var panelBottom = bodyTop + (request.Rows.Count * grid.RowHeightPt);

        List<ScenePrimitive> primitives = [];
        AddHeader(primitives, request, grid, panelLeft, headerTop);
        AddBody(primitives, request, grid, panelLeft, bodyTop);
        AddBorders(primitives, request, grid, panelLeft, headerTop, bodyTop, panelBottom);

        return new PanelBuildOutcome(
            new PanelBuildResult(
                primitives,
                new RectD(panelLeft, headerTop, grid.TotalWidthPt, panelBottom - headerTop),
                new RectD(panelLeft, headerTop, grid.TotalWidthPt, request.HeaderHeightPt)),
            null);
    }
    private static void AddHeader(
        List<ScenePrimitive> primitives,
        PanelBuildRequest request,
        PanelCellGrid grid,
        double panelLeft,
        double headerTop)
    {
        var x = panelLeft;
        foreach (PanelColumn column in grid.Columns)
        {
            var bounds = new RectD(x, headerTop, column.WidthPt, request.HeaderHeightPt);
            primitives.Add(
                new SceneRect(
                    ScenePrimitive.CreateId(SceneOwnerId.Chart, $"header-cell:{column.Name}"),
                    SceneOwnerId.Chart,
                    ZLayer.Frame,
                    bounds,
                    request.Theme.HeaderFill));
            primitives.Add(
                new SceneText(
                    ScenePrimitive.CreateId(SceneOwnerId.Chart, $"header-text:{column.Name}"),
                    SceneOwnerId.Chart,
                    ZLayer.Frame,
                    column.Name,
                    bounds,
                    request.Theme.HeaderText,
                    HeaderAlignmentFor(column.Name)));
            x += column.WidthPt;
        }
    }

    private static void AddBody(
        List<ScenePrimitive> primitives,
        PanelBuildRequest request,
        PanelCellGrid grid,
        double panelLeft,
        double bodyTop)
    {
        for (var rowIndex = 0; rowIndex < request.Rows.Count; rowIndex++)
        {
            PanelRow row = request.Rows[rowIndex];
            var owner = SceneOwnerId.ForRow(row.RowId);
            var top = bodyTop + (rowIndex * grid.RowHeightPt);
            var x = panelLeft;

            for (var columnIndex = 0; columnIndex < grid.Columns.Count; columnIndex++)
            {
                PanelColumn column = grid.Columns[columnIndex];
                var bounds = new RectD(x, top, column.WidthPt, grid.RowHeightPt);
                primitives.Add(
                    new SceneRect(
                        ScenePrimitive.CreateId(owner, $"panel-cell:{column.Name}"),
                        owner,
                        ZLayer.Background,
                        bounds,
                        request.Theme.BodyFill));

                var text = row.Cells[columnIndex];
                if (!string.IsNullOrWhiteSpace(text))
                {
                    primitives.Add(
                        new SceneText(
                            ScenePrimitive.CreateId(owner, $"panel-text:{column.Name}"),
                            owner,
                            ZLayer.Label,
                            text,
                            bounds,
                            request.Theme.BodyText,
                            column.Alignment));
                }

                x += column.WidthPt;
            }
        }
    }
    /// <summary>
    /// Emits each shared cell edge exactly once, so adjacent cells never double
    /// a border (section 3) and the grid is not one line thicker than the tokens
    /// imply.
    /// </summary>
    private static void AddBorders(
        List<ScenePrimitive> primitives,
        PanelBuildRequest request,
        PanelCellGrid grid,
        double panelLeft,
        double headerTop,
        double bodyTop,
        double panelBottom)
    {
        // One vertical line per column boundary, including both outer edges, each
        // spanning the whole panel so no cell edge is drawn twice.
        var right = panelLeft;
        var x = panelLeft;
        for (var index = 0; index <= grid.Columns.Count; index++)
        {
            primitives.Add(
                new SceneLine(
                    ScenePrimitive.CreateId(SceneOwnerId.Chart, $"panel-border-v:{index}"),
                    SceneOwnerId.Chart,
                    ZLayer.Frame,
                    new PointD(x, headerTop),
                    new PointD(x, panelBottom),
                    request.Theme.Border));

            if (index < grid.Columns.Count)
            {
                x += grid.Columns[index].WidthPt;
                right = x;
            }
        }

        // The panel's top edge, then the header/body boundary, then each row's
        // bottom edge. The boundary is a distinct shared edge: section 3 requires
        // a border at every cell edge, and the header's bottom is the first row's
        // top, so it belongs to the header band and the body alike. With no rows
        // it is also the panel's bottom edge, which keeps an empty panel closed.
        // Identifiers are shifted up by one so each stays unique.
        AddHorizontalBorder(primitives, request, "panel-border-h:0", panelLeft, right, headerTop);
        AddHorizontalBorder(primitives, request, "panel-border-h:1", panelLeft, right, bodyTop);
        for (var index = 0; index < request.Rows.Count; index++)
        {
            var y = bodyTop + ((index + 1) * grid.RowHeightPt);
            AddHorizontalBorder(primitives, request, $"panel-border-h:{index + 2}", panelLeft, right, y);
        }
    }

    private static void AddHorizontalBorder(
        List<ScenePrimitive> primitives,
        PanelBuildRequest request,
        string role,
        double left,
        double right,
        double y) =>
        primitives.Add(
            new SceneLine(
                ScenePrimitive.CreateId(SceneOwnerId.Chart, role),
                SceneOwnerId.Chart,
                ZLayer.Frame,
                new PointD(left, y),
                new PointD(right, y),
                request.Theme.Border));

    /// <summary>
    /// Section 4: headers are centred "unless the schema defines left alignment
    /// for Description". Confirmed by the product owner on 2026-09-26, so the rule
    /// is taken from the schema column name rather than from the measured body
    /// alignment, which governs cell text and not header text.
    /// </summary>
    private static GanttTextAlignment HeaderAlignmentFor(string columnName) =>
        string.Equals(columnName, _descriptionColumn, StringComparison.Ordinal)
            ? GanttTextAlignment.Left
            : GanttTextAlignment.Centre;

    private static PanelBuildOutcome Refused(PanelBuildRefusal refusal) => new(null, refusal);
}
