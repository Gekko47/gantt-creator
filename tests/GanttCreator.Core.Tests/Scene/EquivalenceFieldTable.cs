using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// The R3.16 transcription of the entity guide's "Per-entity field contract" table:
/// every one of its thirteen rows, keyed to the guide's own row names.
/// </summary>
/// <remarks>
/// <para>
/// The field probes live in <see cref="EquivalenceFieldChecks"/> so each row below reads
/// the way the guide reads. <see cref="All"/> is the set the drift guard compares
/// against the parsed guide table.
/// </para>
/// <para>
/// Eleven rows are exercised against the real built scene. The panel row is exercised
/// against a purpose-built panel-bearing scene, because the canonical reference build
/// omits the optional panel theme and therefore contains no panel primitive at all.
/// The validation-indicator and legend rows are <em>exclusion</em> rows: the model
/// deliberately carries no such primitive, and each names the token that must appear in
/// no identifier, so its absence is asserted rather than merely untested.
/// </para>
/// </remarks>
internal static class EquivalenceFields
{
    private const string _fieldTable = "Entity guide 'Per-entity field contract'";

    /// <summary>The data-panel row. The guide states the live worksheet draws no shapes.</summary>
    public static EquivalenceRow DataPanel { get; } =
        new(
            "data panel",
            "Data panel/header",
            $"{_fieldTable}, row 'Data panel/header' (line 717); geometry and z-layers from §3 'Data panel' and §4 'Table header'",
            [
                new("header-cell", ScenePrimitive.CreateId(SceneOwnerId.Chart, "header-cell:Id")),
                new("header-text", ScenePrimitive.CreateId(SceneOwnerId.Chart, "header-text:Id")),
                new("border", ScenePrimitive.CreateId(SceneOwnerId.Chart, "panel-border-v:0")),
                new("body-cell", ScenePrimitive.CreateId(SceneOwnerId.ForRow(SceneBuilderTests.PanelProbeRow), "panel-cell:Id")),
                new("body-text", ScenePrimitive.CreateId(SceneOwnerId.ForRow(SceneBuilderTests.PanelProbeRow), "panel-text:Id")),
            ],
            [
                EquivalenceFieldChecks.Kind("header-cell", "rect", $"{EquivalenceFieldChecks.Shared} - point-based geometry"),
                EquivalenceFieldChecks.Extents("header-cell", "Entity guide §3 'Data panel' - the panel reproduces the measured cell grid"),
                EquivalenceFieldChecks.AtLayer("header-cell", ZLayer.Frame, "Entity guide §4 'Table header' and the Z-order table: the header band and shared cell borders sit at Frame (80)"),
                EquivalenceFieldChecks.ChartOwned("header-cell", $"{EquivalenceFieldChecks.Identity}: chart furniture is never a row-reconciliation deletion candidate"),
                EquivalenceFieldChecks.Kind("header-text", "text", $"{EquivalenceFieldChecks.Shared} - point-based geometry"),
                EquivalenceFieldChecks.TextContent("header-text", "Entity guide §4: the header cell carries the schema display name"),
                EquivalenceFieldChecks.TextExtents("header-text", "Entity guide §4 'Table header' - the header text sits in the measured header cell"),
                EquivalenceFieldChecks.AlignmentDefined("header-text", "Entity guide §4: the header is aligned per the schema column (ADR-0018)"),
                EquivalenceFieldChecks.AtLayer("header-text", ZLayer.Frame, "Entity guide §4: the header fill and its text share layer 80"),
                EquivalenceFieldChecks.Kind("border", "line", $"{EquivalenceFieldChecks.Shared} - point-based geometry"),
                EquivalenceFieldChecks.AtLayer("border", ZLayer.Frame, "Entity guide §3 and the Z-order table: the shared cell borders sit at Frame (80)"),
                EquivalenceFieldChecks.LineEndpoints("border", "Entity guide §3 'Data panel': each shared cell edge is emitted once and spans the panel"),
                EquivalenceFieldChecks.ChartOwned("border", $"{EquivalenceFieldChecks.Identity}: the border is chart furniture, owned by the panel and not by any row"),
                EquivalenceFieldChecks.Kind("body-cell", "rect", $"{EquivalenceFieldChecks.Shared} - point-based geometry"),
                // The z-layer R3.16 corrected: a body cell fill is Background, not Frame.
                // Transcribed from the corrected guide cell, and it is the field that
                // would have failed against the uncorrected one.
                EquivalenceFieldChecks.AtLayer("body-cell", ZLayer.Background, "Entity guide §3 and the Z-order table: cell backgrounds sit at Background (0)"),
                EquivalenceFieldChecks.RowOwned("body-cell", "Entity guide §3: a body cell is owned by the row it restates, and dies with that row"),
                EquivalenceFieldChecks.Kind("body-text", "text", $"{EquivalenceFieldChecks.Shared} - point-based geometry"),
                EquivalenceFieldChecks.AtLayer("body-text", ZLayer.Label, "Entity guide §3 and the Z-order table: cell text sits at Label (70), above its own background and below the borders framing it"),
                EquivalenceFieldChecks.TextContent("body-text", "Entity guide §3: the cell text is the worksheet value, not a re-derivation"),
                EquivalenceFieldChecks.TextExtents("body-text", "Entity guide §3 'Data panel' - the cell text sits in the measured cell"),
                EquivalenceFieldChecks.RowOwned("body-text", "Entity guide §3: a body cell is owned by the row it restates"),
            ]);

    /// <summary>The title row. A hidden title emits no primitives at all.</summary>
    public static EquivalenceRow Title { get; } =
        new(
            "chart title",
            "Title",
            $"{_fieldTable}, row 'Title' (line 718); geometry from §1 'Chart title'",
            [
                new("band", ScenePrimitive.CreateId(SceneOwnerId.Chart, "title-band")),
                new("text", ScenePrimitive.CreateId(SceneOwnerId.Chart, "title-text")),
            ],
            [
                EquivalenceFieldChecks.Kind("band", "rect", $"{EquivalenceFieldChecks.Shared} - point-based geometry"),
                EquivalenceFieldChecks.Extents("band", "Entity guide §1: the title band spans the content width"),
                EquivalenceFieldChecks.AtLayer("band", ZLayer.Title, "Entity guide §1 and the Z-order table: the chart title sits at Title (90), above every other layer"),
                EquivalenceFieldChecks.ChartOwned("band", $"{EquivalenceFieldChecks.Identity}: the title is chart furniture"),
                EquivalenceFieldChecks.Kind("text", "text", $"{EquivalenceFieldChecks.Shared} - point-based geometry"),
                EquivalenceFieldChecks.TextContent("text", "Entity guide §1: the title text is the caller-supplied chart title"),
                EquivalenceFieldChecks.TextExtents("text", "Entity guide §1: the title text is placed within the title band"),
                EquivalenceFieldChecks.CentredAlignment("text", "Entity guide §1: the chart title is centred in its band"),
                EquivalenceFieldChecks.AtLayer("text", ZLayer.Title, "Entity guide §1 and the Z-order table: the chart title sits at Title (90)"),
                EquivalenceFieldChecks.ChartOwned("text", $"{EquivalenceFieldChecks.Identity}: the title is chart furniture"),
            ]);

    /// <summary>The year-header row. The year label names its parent by appending ":label".</summary>
    public static EquivalenceRow YearHeader { get; } =
        new(
            "year header",
            "Year header",
            $"{_fieldTable}, row 'Year header' (line 719); geometry from §5 'Year header'",
            [
                new("band", ScenePrimitive.CreateId(SceneOwnerId.Chart, "year:2026")),
                new("label", ScenePrimitive.CreateId(SceneOwnerId.Chart, "year:2026:label")),
            ],
            [
                EquivalenceFieldChecks.Kind("band", "rect", $"{EquivalenceFieldChecks.Shared} - point-based geometry"),
                EquivalenceFieldChecks.Extents("band", "Entity guide §5: the year band spans its year's share of the time scale"),
                EquivalenceFieldChecks.AtLayer("band", ZLayer.Frame, "Entity guide §5 and the Z-order table: the year header sits at Frame (80)"),
                EquivalenceFieldChecks.ChartOwned("band", $"{EquivalenceFieldChecks.Identity}: the year band is chart furniture"),
                EquivalenceFieldChecks.Kind("label", "text", $"{EquivalenceFieldChecks.Shared} - point-based geometry"),
                EquivalenceFieldChecks.TextContent("label", "Entity guide §5: the year label is the year in the chart's display format"),
                EquivalenceFieldChecks.TextExtents("label", "Entity guide §5: the year label is measured inside the year band"),
                EquivalenceFieldChecks.CentredAlignment("label", "Entity guide §5: the year label is centred in its band"),
                EquivalenceFieldChecks.AtLayer("label", ZLayer.Frame, "Entity guide §5 and the Z-order table: the year label sits at Frame (80)"),
                EquivalenceFieldChecks.NoOrderKeys("label", $"{_fieldTable}, row 'Year header' (line 719): no lane/stack keys"),
                // R3.17's convention, asserted as a field: the label names its parent by
                // appending ":label". This is the exact text a renderer reconciles.
                EquivalenceFieldChecks.LabelNamesParentRectangle("label", $"{_fieldTable}, row 'Year header' (line 719), and R3.17 D2: a year label names its parent by appending ':label'"),
                EquivalenceFieldChecks.ChartOwned("label", $"{EquivalenceFieldChecks.Identity}: the year label is chart furniture"),
            ]);

    /// <summary>The period-header row, keyed to the first period in the reference scene.</summary>
    public static EquivalenceRow PeriodHeader { get; } =
        new(
            "period header",
            "Period header",
            $"{_fieldTable}, row 'Period header' (line 720); geometry from §5 'Period header'",
            [
                new("band", ScenePrimitive.CreateId(SceneOwnerId.Chart, "period:2026-01-05")),
                new("label", ScenePrimitive.CreateId(SceneOwnerId.Chart, "period:2026-01-05:label")),
            ],
            [
                EquivalenceFieldChecks.Kind("band", "rect", $"{EquivalenceFieldChecks.Shared} - point-based geometry"),
                EquivalenceFieldChecks.Extents("band", "Entity guide §5: the period band spans its period's share of the time scale"),
                EquivalenceFieldChecks.AtLayer("band", ZLayer.Frame, "Entity guide §5 and the Z-order table: the period header sits at Frame (80)"),
                EquivalenceFieldChecks.ChartOwned("band", $"{EquivalenceFieldChecks.Identity}: the period band is chart furniture"),
                EquivalenceFieldChecks.Kind("label", "text", $"{EquivalenceFieldChecks.Shared} - point-based geometry"),
                EquivalenceFieldChecks.TextContent("label", "Entity guide §5: the period label is the period in the chart's display format"),
                EquivalenceFieldChecks.TextExtents("label", "Entity guide §5: the period label is measured inside the period band"),
                EquivalenceFieldChecks.CentredAlignment("label", "Entity guide §5: the period label is centred in its band"),
                EquivalenceFieldChecks.AtLayer("label", ZLayer.Frame, "Entity guide §5 and the Z-order table: the period label sits at Frame (80)"),
                EquivalenceFieldChecks.NoOrderKeys("label", $"{_fieldTable}, row 'Period header' (line 720): no lane/stack keys"),
                EquivalenceFieldChecks.LabelNamesParentRectangle("label", $"{_fieldTable}, row 'Period header' (line 720): the period label appends ':label' to its rectangle's id"),
                EquivalenceFieldChecks.ChartOwned("label", $"{EquivalenceFieldChecks.Identity}: the period label is chart furniture"),
            ]);

    /// <summary>The band, grid, and frame row. The background carries the derived chart bounds.</summary>
    public static EquivalenceRow BandsGridFrame { get; } =
        new(
            "bands, grid and frame",
            "Bands/grid/frame",
            $"{_fieldTable}, row 'Bands/grid/frame' (line 721); geometry from §2 'Chart bounds' and §6-§8",
            [
                new("background", ScenePrimitive.CreateId(SceneOwnerId.Chart, "background")),
                new("band", ScenePrimitive.CreateId(SceneOwnerId.Chart, "band:1")),
                new("grid", ScenePrimitive.CreateId(SceneOwnerId.Chart, "grid:820")),
                new("frame", ScenePrimitive.CreateId(SceneOwnerId.Chart, "frame:top")),
            ],
            [
                EquivalenceFieldChecks.Kind("background", "rect", $"{EquivalenceFieldChecks.Shared} - point-based geometry"),
                EquivalenceFieldChecks.Extents("background", "Entity guide §2: the chart background covers the chart bounds"),
                // R3.15's contract, asserted from the catalogue side: the declared
                // ChartBounds and the chart:background primitive are one rectangle.
                EquivalenceFieldChecks.EqualsSceneChartBounds("background", "Entity guide §2: 'ChartBounds is the union of the title, data panel, time headers, and plot plus ChartOuterPaddingPt' - the frame derives it, and the scene must declare that same value (R3.15 D1)"),
                EquivalenceFieldChecks.AtLayer("background", ZLayer.Background, "Entity guide §2 and the Z-order table: the chart background sits at Background (0)"),
                EquivalenceFieldChecks.ChartOwned("background", $"{EquivalenceFieldChecks.Identity}: the chart background is chart furniture"),
                EquivalenceFieldChecks.Kind("band", "rect", $"{EquivalenceFieldChecks.Shared} - point-based geometry"),
                EquivalenceFieldChecks.Extents("band", "Entity guide §6: an alternating time band spans its time slice"),
                EquivalenceFieldChecks.AtLayer("band", ZLayer.AlternateBand, "Entity guide §6 and the Z-order table: alternating time bands sit at AlternateBand (10)"),
                EquivalenceFieldChecks.Kind("grid", "line", $"{EquivalenceFieldChecks.Shared} - point-based geometry"),
                EquivalenceFieldChecks.AtLayer("grid", ZLayer.Grid, "Entity guide §7 and the Z-order table: grid lines sit at Grid (20)"),
                EquivalenceFieldChecks.LineEndpoints("grid", "Entity guide §7: a grid line spans the plot height at its time-slice boundary"),
                EquivalenceFieldChecks.Kind("frame", "line", $"{EquivalenceFieldChecks.Shared} - point-based geometry"),
                EquivalenceFieldChecks.AtLayer("frame", ZLayer.Frame, "Entity guide §8 and the Z-order table: the plot frame sits at Frame (80)"),
                EquivalenceFieldChecks.LineEndpoints("frame", "Entity guide §8: a frame edge spans its side of the plot"),
                EquivalenceFieldChecks.ChartOwned("frame", $"{EquivalenceFieldChecks.Identity}: the plot frame is chart furniture"),
            ]);

    /// <summary>
    /// The activity row. The guide's "Activity/delay" row requires a rectangle shape in
    /// the live worksheet, a rectangle shape in the editable composition and
    /// PowerPoint, and a raster rectangle in the PNG; all three are the same
    /// <see cref="SceneRect"/>, so these are the fields each renderer reads to draw it.
    /// </summary>
    public static EquivalenceRow ActivityBody { get; } =
        new(
            "span bar",
            "Activity/delay",
            $"{_fieldTable}, row 'Activity/delay' (line 722); geometry and style from §12 'General span activity' and §13 'As-planned activity'",
            [new("bar", ScenePrimitive.CreateId(SceneOwnerId.ForRow(SceneBuilderTests.BarProbeRow), "bar"))],
            [
                EquivalenceFieldChecks.Kind("bar", "rect", $"{EquivalenceFieldChecks.Shared} - point-based geometry"),
                EquivalenceFieldChecks.Extents("bar", "Entity guide §12 'Geometry': the bar is a resolved point rectangle, already plot-clipped"),
                EquivalenceFieldChecks.FillResolved("bar", "Entity guide §12 'Style': fill and outline are explicit; no renderer defaults"),
                EquivalenceFieldChecks.StrokeResolved("bar", "Entity guide §12 'Style': fill and outline are explicit; no renderer defaults"),
                EquivalenceFieldChecks.OutlineWidthResolved("bar", "Entity guide §12 'Style': the standard outline is resolved before rendering"),
                EquivalenceFieldChecks.NoHatch("bar", "Entity guide 'Type catalogue': an activity is Fill + outline, never Hatch + outline"),
                EquivalenceFieldChecks.AtLayer("bar", ZLayer.ActivityBody, $"{EquivalenceFieldChecks.ZOrder} - activity bodies sit at layer 40"),
                EquivalenceFieldChecks.OrderKeys("bar", "Entity guide 'Z-order contract' (line 272): after subtype priority, order by lane, stack, SortOrder, and stable ID"),
                EquivalenceFieldChecks.RoleDerivedId("bar", "bar", "Entity guide §12 with the R3.6 role-derived ':bar' identifier"),
                EquivalenceFieldChecks.RowOwned("bar", "Entity guide 'Shared entity contract' (line 43): a stable EntityId unrelated to worksheet row number"),
            ]);

    /// <summary>
    /// The procurement row. Its hatch field is asserted as <em>readable</em> only: the
    /// reference style set resolves HatchPattern to None for every style, so no
    /// committed primitive carries a hatch and a hatch-value assertion would test a
    /// value nothing produces.
    /// </summary>
    public static EquivalenceRow Procurement { get; } =
        new(
            "procurement bar",
            "Procurement",
            $"{_fieldTable}, row 'Procurement' (line 723); style from §16 'Procurement'. The host representation is unknown until the R4.6/R8.3 proof and the hatch value is unexercised, so only the scene fields are asserted",
            [new("bar", ScenePrimitive.CreateId(SceneOwnerId.ForRow(SceneBuilderTests.ProcurementProbeRow), "bar"))],
            [
                EquivalenceFieldChecks.Kind("bar", "rect", $"{EquivalenceFieldChecks.Shared} - point-based geometry"),
                EquivalenceFieldChecks.Extents("bar", "Entity guide §16: a procurement bar is a resolved span rectangle, already plot-clipped"),
                EquivalenceFieldChecks.FillResolved("bar", "Entity guide §16 'Style': a procurement bar resolves its fill and outline like any other span"),
                EquivalenceFieldChecks.StrokeResolved("bar", "Entity guide §16 'Style': a procurement bar resolves its stroke"),
                EquivalenceFieldChecks.OutlineWidthResolved("bar", "Entity guide §16 'Style': the resolved outline is drawn as resolved"),
                EquivalenceFieldChecks.HatchFieldReadable("bar", $"{_fieldTable}, row 'Procurement' (line 723): the style carries a HatchPattern field for a renderer to read. Unexercised: the reference style set resolves it to None, so the value itself is not asserted here"),
                EquivalenceFieldChecks.AtLayer("bar", ZLayer.ActivityBody, $"{EquivalenceFieldChecks.ZOrder} - a procurement bar is an activity body, layer 40"),
                EquivalenceFieldChecks.OrderKeys("bar", "Entity guide 'Z-order contract' (line 272): after subtype priority, order by lane, stack, SortOrder, and stable ID"),
                EquivalenceFieldChecks.RoleDerivedId("bar", "bar", "Entity guide §16 with the role-derived ':bar' identifier"),
                EquivalenceFieldChecks.RowOwned("bar", "Entity guide §16: a procurement bar is owned by its own row"),
            ]);

    /// <summary>
    /// The critical-interval row. The scene primitive is a thin <see cref="SceneRect"/>
    /// but the host object is a <em>line</em> in all three renderers, so the fields
    /// state the overlay's height and the guide's translation column carries the rest -
    /// that translation is the reason the column exists.
    /// </summary>
    public static EquivalenceRow CriticalInterval { get; } =
        new(
            "critical interval overlay",
            "Critical interval",
            $"{_fieldTable}, row 'Critical interval' (line 724); geometry from §11 'Critical interval'. The guide states the scene rect is not the host object: a renderer draws a line along its top edge",
            [new("overlay", ScenePrimitive.CreateId(SceneOwnerId.ForRow(SceneBuilderTests.CriticalProbeRow), "critical"))],
            [
                EquivalenceFieldChecks.Kind("overlay", "rect", $"{EquivalenceFieldChecks.Shared} - point-based geometry. §11 models the overlay as a thin rectangle in the scene; the *host* object is a line"),
                EquivalenceFieldChecks.CriticalLineHeight("overlay", "Entity guide §11 and the 'Critical interval' row: the overlay is half the predetermined ActivityHeightPt tall (owner ruling 2026-09-30), so a renderer draws the rect it is handed rather than a configured-thickness line"),
                EquivalenceFieldChecks.AtLayer("overlay", ZLayer.CriticalOverlay, $"{EquivalenceFieldChecks.ZOrder} - critical interval overlays sit at layer 50"),
                EquivalenceFieldChecks.StrokeResolved("overlay", "Entity guide §11 'Style': the overlay resolves the standard outline, which is the line thickness a renderer draws"),
                EquivalenceFieldChecks.OutlineWidthResolved("overlay", "Entity guide §11 'Style': the resolved outline is the resolved line width"),
                EquivalenceFieldChecks.OrderKeys("overlay", "Entity guide 'Z-order contract' (line 272): after subtype priority, order by lane, stack, SortOrder, and stable ID"),
                EquivalenceFieldChecks.RoleDerivedId("overlay", "critical", "Entity guide §11 with the R3.8 role-derived ':critical' identifier"),
                EquivalenceFieldChecks.RowOwned("overlay", "Entity guide §11: a critical overlay is owned by its own row"),
            ]);

    /// <summary>
    /// The milestone row. The guide's "Milestone" row requires a four-point freeform
    /// polygon in the live worksheet, a freeform polygon in the editable composition
    /// and PowerPoint, and a raster polygon in the PNG.
    /// </summary>
    public static EquivalenceRow MilestoneDiamond { get; } =
        new(
            "milestone diamond",
            "Milestone",
            $"{_fieldTable}, row 'Milestone' (line 725); geometry from §20 'General milestone diamond' and style from §21",
            [new("marker", ScenePrimitive.CreateId(SceneOwnerId.ForRow(SceneBuilderTests.MilestoneProbeRow), "marker"))],
            [
                EquivalenceFieldChecks.Kind("marker", "polygon", $"{EquivalenceFieldChecks.Shared} - point-based geometry"),
                EquivalenceFieldChecks.DiamondDrawOrder("marker", "Entity guide §20 'Geometry': build a four-point polygon - not a rotated square"),
                EquivalenceFieldChecks.TipToTipExtent("marker", "Entity guide §20 'Geometry': size = MilestoneSizePt makes the tip-to-tip bounds exact and consistent across Excel, PowerPoint, and PNG"),
                EquivalenceFieldChecks.FillResolved("marker", "Entity guide §21: a planned milestone fills with PlannedFill"),
                EquivalenceFieldChecks.StrokeResolved("marker", "Entity guide §21: a planned milestone outlines with PlannedOutline"),
                EquivalenceFieldChecks.OutlineWidthResolved("marker", "Entity guide §21: the resolved outline is drawn as resolved"),
                EquivalenceFieldChecks.NoHatch("marker", "Entity guide 'Type catalogue': a milestone is Fill + outline, never Hatch + outline"),
                EquivalenceFieldChecks.AtLayer("marker", ZLayer.Milestone, $"{EquivalenceFieldChecks.ZOrder} - milestone diamonds sit at layer 60"),
                EquivalenceFieldChecks.OrderKeys("marker", "Entity guide 'Z-order contract' (line 272): after subtype priority, order by lane, stack, SortOrder, and stable ID"),
                EquivalenceFieldChecks.RoleDerivedId("marker", "marker", "Entity guide §20 with the R3.7 role-derived ':marker' identifier"),
                EquivalenceFieldChecks.RowOwned("marker", "Entity guide 'Shared entity contract' (line 43): a stable EntityId unrelated to worksheet row number"),
            ]);

    /// <summary>
    /// The description/date-label row. The label side is already resolved into
    /// <c>TextBounds</c>, so a renderer must not re-measure or re-select a side.
    /// </summary>
    public static EquivalenceRow DescriptionAndDateLabels { get; } =
        new(
            "description and date labels",
            "Description/date labels",
            $"{_fieldTable}, row 'Description/date labels' (line 726); placement from §22 'Activity or milestone description label' and §23 'Date labels'",
            [
                new("description", ScenePrimitive.CreateId(SceneOwnerId.ForRow(SceneBuilderTests.BarProbeRow), "label")),
                new("date-start", ScenePrimitive.CreateId(SceneOwnerId.ForRow(SceneBuilderTests.MilestoneProbeRow), "date-start")),
            ],
            [
                EquivalenceFieldChecks.Kind("description", "text", $"{EquivalenceFieldChecks.Shared} - point-based geometry"),
                EquivalenceFieldChecks.TextContent("description", "Entity guide §22: the description is user-facing text"),
                EquivalenceFieldChecks.TextExtents("description", "Entity guide §22: the label side is already resolved into the text bounds; a renderer must not re-measure or re-select a side"),
                EquivalenceFieldChecks.AlignmentDefined("description", "Entity guide §22 with ADR-0018: text alignment is a resolved style value, not a label position"),
                EquivalenceFieldChecks.AtLayer("description", ZLayer.Label, $"{EquivalenceFieldChecks.ZOrder} - activity, milestone, and date labels sit at layer 70"),
                EquivalenceFieldChecks.OrderKeys("description", "Entity guide 'Z-order contract' (line 272): after subtype priority, order by lane, stack, SortOrder, and stable ID"),
                EquivalenceFieldChecks.RoleDerivedId("description", "label", "Entity guide §22: label IDs are derived from the parent event ID and label role"),
                EquivalenceFieldChecks.RowOwned("description", "Entity guide §22: a description label is owned by the event it describes"),
                EquivalenceFieldChecks.Kind("date-start", "text", $"{EquivalenceFieldChecks.Shared} - point-based geometry"),
                EquivalenceFieldChecks.TextContent("date-start", "Entity guide §23: the date label is the event date in the approved display format"),
                EquivalenceFieldChecks.TextExtents("date-start", "Entity guide §23: the date label is already measured; a clipped event still shows its true date"),
                EquivalenceFieldChecks.AtLayer("date-start", ZLayer.Label, $"{EquivalenceFieldChecks.ZOrder} - date labels sit at layer 70 with the other row labels"),
                EquivalenceFieldChecks.RoleDerivedId("date-start", "date-start", "Entity guide §23 with the R3.11 role-derived ':date-start' identifier"),
                EquivalenceFieldChecks.RowOwned("date-start", "Entity guide §23: a date label is owned by the event it dates"),
            ]);

    /// <summary>
    /// The delineator row. A same-date pair emits <em>one shared line</em> but
    /// <em>one label per row</em>, and the line's owner may be a membership-sensitive
    /// <c>Rows</c> set (ADR-0017), so its ownership field is deliberately not
    /// <c>RowOwned</c>.
    /// </summary>
    public static EquivalenceRow Delineator { get; } =
        new(
            "delineator",
            "Delineator",
            $"{_fieldTable}, row 'Delineator' (line 727); geometry from §24 'Delineators'",
            [
                new("line", SceneBuilderTests.SharedDelineatorPrimitiveId),
                new("label", ScenePrimitive.CreateId(SceneOwnerId.ForRow(SceneBuilderTests.DelineatorProbeRow), "delineator-label")),
            ],
            [
                EquivalenceFieldChecks.Kind("line", "line", $"{EquivalenceFieldChecks.Shared} - point-based geometry"),
                EquivalenceFieldChecks.FullPlotHeight("line", "Entity guide §24: a delineator is a full-plot-height line, spanning PlotBounds top to bottom"),
                EquivalenceFieldChecks.AtLayer("line", ZLayer.Delineator, $"{EquivalenceFieldChecks.ZOrder} - delineator lines sit at layer 25"),
                EquivalenceFieldChecks.StrokeResolved("line", "Entity guide §24 'Style': the line resolves its stroke"),
                EquivalenceFieldChecks.OutlineWidthResolved("line", "Entity guide §24 'Style': DelineatorLinePt is the resolved line width"),
                EquivalenceFieldChecks.SharedOrRowOwner("line", "Entity guide §24 and ADR-0017: a same-date pair emits one shared line under a membership-sensitive Rows owner; a membership change is remove-then-recreate, never in-place"),
                EquivalenceFieldChecks.Kind("label", "text", $"{EquivalenceFieldChecks.Shared} - point-based geometry"),
                EquivalenceFieldChecks.TextContent("label", "Entity guide §24: the delineator label is the resolved line style name"),
                EquivalenceFieldChecks.TextExtents("label", "Entity guide §24: the delineator label is stacked below the plot by StackGapPt and already measured"),
                EquivalenceFieldChecks.AtLayer("label", ZLayer.DelineatorLabel, $"{EquivalenceFieldChecks.ZOrder} - delineator labels sit at layer 75, above row labels"),
                EquivalenceFieldChecks.NoOrderKeys("label", $"{_fieldTable}, row 'Delineator' (line 727): no lane/stack keys"),
                EquivalenceFieldChecks.RowOwned("label", "Entity guide §24: a delineator label is owned by the row that contributed the line, one label per row"),
            ]);

    /// <summary>The validation indicator carries no scene primitive at all.</summary>
    public static EquivalenceRow ValidationIndicator { get; } =
        new(
            "validation indicator",
            "Validation indicator",
            $"{_fieldTable}, row 'Validation indicator' (line 728); §26 'Validation indicators'",
            [],
            [],
            "validation");

    /// <summary>The optional legend has no scene primitive today.</summary>
    public static EquivalenceRow Legend { get; } =
        new(
            "legend",
            "Legend (§25, optional)",
            $"{_fieldTable}, row 'Legend (§25, optional)' (line 729); §25 'Legend' is optional, hidden by default, and product-owner gated",
            [],
            [],
            "legend");

    /// <summary>
    /// Every row of the field-contract table, in the guide's own order. The drift guard
    /// compares this row-name set against the parsed guide table.
    /// </summary>
    public static IReadOnlyList<EquivalenceRow> All { get; } =
    [
        DataPanel,
        Title,
        YearHeader,
        PeriodHeader,
        BandsGridFrame,
        ActivityBody,
        Procurement,
        CriticalInterval,
        MilestoneDiamond,
        DescriptionAndDateLabels,
        Delineator,
        ValidationIndicator,
        Legend,
    ];
}
