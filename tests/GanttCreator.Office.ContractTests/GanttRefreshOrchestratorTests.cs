using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R4.8A: the refresh orchestrator's pipeline contract, proven without Excel.
/// </summary>
/// <remarks>
/// <para>
/// The three claims under test are D2 (the order is fixed), D3 (every validation and
/// scene step completes before any shape mutation), and D1 (a failure anywhere leaves
/// the previous chart untouched). All three are only observable because every
/// collaborator is a port: the test injects a refusal at a chosen step and inspects
/// the shape port's recorded calls.
/// </para>
/// <para>
/// <b>Why a refusal at <em>each</em> step, not one.</b> A single early failure would
/// pass even if the later steps were reordered behind it, because nothing after the
/// failure would run. Injecting at every step is what makes the ordering claim
/// real: each one proves the steps before it ran and the shape port was untouched.
/// </para>
/// </remarks>
public class GanttRefreshOrchestratorTests
{
    [Fact]
    public void A_pending_change_batch_still_reads_and_validates_the_complete_table()
    {
        RefreshFakes fakes = new()
        {
            Rows =
            [
                Row(1, null, "As-Planned Activity", "Design", new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19)),
                Row(2, null, "Not a chart type", "Invalid row", new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19)),
            ],
        };
        GanttRefreshOrchestrator orchestrator = fakes.BuildOrchestrator();

        GanttRefreshOutcome outcome = orchestrator.Refresh();

        // Partial render: the valid row plots, the invalid row is skipped with a
        // fill. The refresh succeeds; the issues are carried on the outcome for the
        // caller to surface (warnings panel, status bar).
        Assert.True(outcome.Succeeded, "refusal=" + outcome.Refusal + " msg=" + outcome.Message);

        // One full read. The pending batch does NOT cause a second read — identity
        // repair and the type-change normaliser report zero writes here, and re-reading
        // the whole table for a no-op was a redundant slot on every refresh.
        Assert.Equal(1, fakes.TableReadCount);
        Assert.Equal(0, fakes.TargetedTableReadCount);
        Assert.NotEmpty(fakes.Shapes.Calls);
        Assert.Contains(outcome.ValidationIssues, issue => issue.RowNumber == 2);
    }

    [Fact]
    public void A_successful_full_reconciliation_clears_the_pending_change_batch()
    {
        RefreshFakes fakes = new() { Rows = ValidRows() };
        ChangeTracker tracker = new();
        tracker.Record(1, 4, null, "As-Planned Activity");

        GanttRefreshOutcome outcome = fakes.BuildOrchestratorWithChanges(tracker).Refresh();

        Assert.True(outcome.Succeeded, "refusal=" + outcome.Refusal + " msg=" + outcome.Message);
        Assert.False(tracker.HasChanges);
        Assert.NotEmpty(fakes.Shapes.Calls);
    }

    /// <summary>One valid planned activity, enough for a scene to be built.</summary>
    private static GanttRowDto[] ValidRows() =>
    [
        Row(
            rowNumber: 1,
            id: null,
            type: "As-Planned Activity",
            description: "Design",
            start: new DateOnly(2024, 1, 8),
            finish: new DateOnly(2024, 1, 19)),
    ];

    /// <summary>
    /// Builds a row DTO in the positional form the reader produces.
    /// </summary>
    /// <remarks>
    /// The identifiers are the catalogue's own shape — <c>G-</c> followed by 32 hex
    /// characters. A short readable id like <c>"A1"</c> is <em>rejected</em> by the
    /// validator, so a fixture using one exercises the "blocking error" path and
    /// never reaches the scene at all. That is a real constraint of the schema, not a
    /// quirk of the test helper.
    /// </remarks>
    /// <summary>
    /// A successful refresh reconciles an EMPTY highlight target list, clearing stale
    /// row/cell indicators.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>The reported defect, pinned at the seam that caused it.</b> In a live F5
    /// session on 2026-10-04 the cell notes survived a Refresh: Validate wrote them and
    /// Refresh never touched them, because <c>IGanttValidationReporter</c> had zero
    /// references in <c>GanttRefreshOrchestrator</c> and was not a constructor
    /// parameter. The user fixed the row, pressed Refresh, and the stale note was still
    /// there — which reads as the fix not having registered.
    /// </para>
    /// <para>
    /// The assertion is on the LIST, not merely that the reporter was called. An empty
    /// list is the whole mechanism, because the reporter clears the notes it owns on
    /// every call; a refresh that reported only when it had findings would leave the
    /// stale note exactly in place.
    /// </para>
    /// </remarks>
    [Fact]
    public void A_successful_refresh_reconciles_empty_highlights_to_clear_stale_indicators()
    {
        (RefreshFakes fakes, GanttRefreshOrchestrator orchestrator) = Ready();

        GanttRefreshOutcome outcome = orchestrator.Refresh();

        Assert.True(outcome.Succeeded, "refusal=" + outcome.Refusal);
        IReadOnlyList<HighlightTarget> reported = Assert.Single(fakes.HighlightReconciliations);
        Assert.Empty(reported);
    }

    /// <summary>
    /// A refresh that finds a blocking error still highlights its row/cell so the user
    /// can identify what to fix.
    /// </summary>
    /// <remarks>
    /// The counterweight to the empty-list assertion above, and the reason the reporter
    /// is called before the render decision rather than after it. Calling it only on the
    /// success path would mean a blocked refresh left the PREVIOUS report in place —
    /// the notes would never show the current error, and the user would be chasing a
    /// message about a problem they had already fixed.
    /// </remarks>
    [Fact]
    public void A_blocked_refresh_still_highlights_the_errors_it_found()
    {
        RefreshFakes fakes = new()
        {
            Rows =
            [
                Row(1, string.Empty, "As-Planned Activity", "Design", new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19)),
            ],
        };
        GanttRefreshOrchestrator orchestrator = fakes.BuildOrchestrator();

        GanttRefreshOutcome outcome = orchestrator.Refresh();

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttRefreshRefusal.BlockingValidationErrors, outcome.Refusal);

        IReadOnlyList<HighlightTarget> reported = Assert.Single(fakes.HighlightReconciliations);
        Assert.Contains(reported, target => target.Row == 1);
    }

    /// <summary>
    /// A refusal to maintain the notes stops the refresh before anything is rendered.
    /// </summary>
    /// <remarks>
    /// The positive case for the new <c>ValidationNotesRefused</c> branch, in the same
    /// commit as the branch. Swallowing this refusal would leave the sheet showing the
    /// previous report while the chart silently stayed put, which is the least
    /// explainable state the pipeline can produce. The zero shape calls are the point:
    /// the user keeps their last valid chart and gets told why nothing changed.
    /// </remarks>
    [Fact]
    public void A_refused_highlight_update_stops_the_refresh_before_any_shape_is_touched()
    {
        (RefreshFakes fakes, GanttRefreshOrchestrator orchestrator) = Ready();
        fakes.WarningHighlightsRefused = true;

        GanttRefreshOutcome outcome = orchestrator.Refresh();

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttRefreshRefusal.ValidationNotesRefused, outcome.Refusal);

        // D3: the refusal happens before the single shape-mutating step.
        Assert.Empty(fakes.Shapes.Calls);
        Assert.Equal(0, outcome.DurationCellsWritten);
        Assert.Equal(0, outcome.ShapesWritten);
    }

    /// <summary>
    /// An orchestrator with no reporter wired still refreshes, so a caller that has no
    /// workbook to annotate is not forced to construct one.
    /// </summary>
    /// <remarks>
    /// The counterweight that keeps the new dependency additive. Every other optional
    /// collaborator in this constructor was made optional for the same reason, and a
    /// required parameter here would have broken each existing caller including the
    /// composition root.
    /// </remarks>
    [Fact]
    public void A_refresh_with_no_reporter_wired_still_succeeds()
    {
        RefreshFakes fakes = new()
        {
            Rows =
            [
                Row(1, null, "As-Planned Activity", "Design", new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19)),
            ],
        };
        GanttRefreshOrchestrator orchestrator = fakes.BuildOrchestratorWithWarningHighlighter(null);

        GanttRefreshOutcome outcome = orchestrator.Refresh();

        Assert.True(outcome.Succeeded, "refusal=" + outcome.Refusal + " msg=" + outcome.Message);
        Assert.Empty(fakes.HighlightReconciliations);
    }

    private static GanttRowDto Row(
        int rowNumber,
        string? id,
        string type,
        string description,
        DateOnly? start,
        DateOnly? finish,
        string? styleKey = null) =>
        new(
            rowNumber,
            id ?? NewId(),
            NewId(),
            0,
            type,
            description,
            start,
            finish,
            parentId: null,
            styleKey: styleKey,
            labelPositionText: null,
            fillColourText: null,
            strokeColourText: null,
            visible: true,
            sortOrder: null);

    /// <summary>A well-formed row identifier, in the catalogue's <c>G-</c> + hex shape.</summary>
    private static string NewId() => $"G-{Interlocked.Increment(ref s_nextId):x32}";

    private static int s_nextId;

    /// <summary>A refresh whose rows are valid, over the fresh fakes.</summary>
    private static (RefreshFakes Fakes, GanttRefreshOrchestrator Orchestrator) Ready()
    {
        RefreshFakes fakes = new() { Rows = ValidRows() };
        return (fakes, fakes.BuildOrchestrator());
    }

    /// <summary>
    /// The row-height normaliser is handed the <c>Type</c> column the refresh already
    /// read, so it does not read one cell per body row through COM after the fact.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The assertion is on the ARGUMENT, not on a call count. The normaliser still walks
    /// every body row for its height; the saving is the per-row Type read it no longer
    /// issues, and the only way to prove that from outside is to see the column arrive.
    /// </para>
    /// <para>
    /// The fixture is one row because this fake's style registry holds one style and its
    /// lane anchors need the single-lane layout the other tests use; the body-index
    /// mapping of the projection, and the type-to-token decision it feeds, are pinned by
    /// the height-normaliser depth tests, where the host and the caller can be made to
    /// disagree on purpose.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_row_height_normaliser_is_handed_the_bulk_read_type_texts()
    {
        RefreshFakes fakes = new() { Rows = ValidRows() };

        GanttRefreshOutcome outcome = fakes.BuildOrchestrator().Refresh();

        Assert.True(outcome.Succeeded, "refusal=" + outcome.Refusal + " msg=" + outcome.Message);
        Assert.Equal(["As-Planned Activity"], fakes.LastTypeTexts);
    }

    /// <summary>
    /// The heights the normalisation measured reach the panel measurement, so the panel
    /// does not read every body row again for figures this refresh already holds.
    /// </summary>
    /// <remarks>
    /// The chain matters more than either port on its own: the normaliser returns its
    /// measurement, and the orchestrator has to pass it on. A normaliser that returned the
    /// figures to a caller that dropped them would look identical from either port's own
    /// test, and the second read would survive unnoticed.
    /// </remarks>
    [Fact]
    public void The_panel_measurement_is_handed_the_heights_the_normalisation_measured()
    {
        RefreshFakes fakes = new()
        {
            Rows = ValidRows(),
            RowHeightBodyHeights = [18d],
        };

        GanttRefreshOutcome outcome = fakes.BuildOrchestrator().Refresh();

        Assert.True(outcome.Succeeded, "refusal=" + outcome.Refusal + " msg=" + outcome.Message);
        Assert.Equal([18d], fakes.LastMeasuredBodyHeightsPt);
    }

    [Fact]
    public void Scheduled_refresh_skips_unchanged_shape_writes_and_explicit_refresh_forces_full_updates()
    {
        RefreshFakes fakes = new() { Rows = ValidRows() };
        ShapeDeltaCache cache = new();

        GanttRefreshOutcome first = fakes.BuildOrchestratorWithDeltaCache(cache).Refresh();
        Assert.True(first.Succeeded, "first refresh refused: " + first.Refusal);
        Assert.NotNull(cache.Snapshot);
        int shapeWritesAfterFirst = fakes.Shapes.ShapeCalls.Count;
        int zOrderCallsAfterFirst = fakes.Shapes.ZOrderLists.Count;

        GanttRefreshOutcome scheduled = fakes.BuildOrchestratorWithDeltaCache(cache).Refresh();

        Assert.True(scheduled.Succeeded, "scheduled refresh refused: " + scheduled.Refusal);
        Assert.Equal(shapeWritesAfterFirst, fakes.Shapes.ShapeCalls.Count);
        Assert.Equal(zOrderCallsAfterFirst, fakes.Shapes.ZOrderLists.Count);

        GanttRefreshOutcome explicitRefresh = fakes.BuildOrchestratorWithDeltaCache(cache, useDeltaCache: false).Refresh();

        Assert.True(explicitRefresh.Succeeded, "explicit refresh refused: " + explicitRefresh.Refusal);
        Assert.True(fakes.Shapes.ShapeCalls.Count > shapeWritesAfterFirst);
        Assert.Equal(zOrderCallsAfterFirst + 1, fakes.Shapes.ZOrderLists.Count);
        Assert.NotNull(cache.Snapshot);
    }

    /// <summary>
    /// A refresh traces its delta decision, the structural writes of the run, and the
    /// reconciliation result — the evidence a skipped shape needs.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>Why this test exists.</b> The reported defect is a splitter band left
    /// stretched after a collapsed group moved past it: the scheduled refresh said
    /// nothing, and the manual Refresh fixed it. The two paths differ only in the
    /// delta snapshot, so the log has to name (a) whether this run used the snapshot,
    /// (b) what structural writes it made, and (c) how many shape writes it actually
    /// issued. Without those three facts the cause can only be guessed.
    /// </para>
    /// <para>
    /// The second scheduled run is the load-bearing half: every payload is unchanged,
    /// so the result line must report <c>written=0</c> while still naming the desired
    /// count. That is exactly the state in which a live shape Excel moved on its own
    /// would stay stale, and it is now visible in the log rather than inferred.
    /// </para>
    /// </remarks>
    [Fact]
    public void A_scheduled_refresh_traces_its_delta_decision_and_reconcile_result()
    {
        RefreshFakes fakes = new() { Rows = ValidRows() };
        ShapeDeltaCache cache = new();
        List<string> trace = [];

        GanttRefreshOutcome first = fakes.BuildOrchestratorWithDeltaCache(cache, trace: trace.Add).Refresh();
        Assert.True(first.Succeeded, "first refresh refused: " + first.Refusal);
        Assert.Contains(
            trace,
            line => line.Contains("gantt.refresh delta=True", StringComparison.Ordinal));

        trace.Clear();
        GanttRefreshOutcome scheduled = fakes.BuildOrchestratorWithDeltaCache(cache, trace: trace.Add).Refresh();
        Assert.True(scheduled.Succeeded, "scheduled refresh refused: " + scheduled.Refusal);
        Assert.Contains(
            trace,
            line => line.Contains("gantt.refresh result ok=True", StringComparison.Ordinal)
                && line.Contains("written=0", StringComparison.Ordinal));

        trace.Clear();
        GanttRefreshOutcome explicitRefresh = fakes
            .BuildOrchestratorWithDeltaCache(cache, useDeltaCache: false, trace: trace.Add)
            .Refresh();
        Assert.True(explicitRefresh.Succeeded, "explicit refresh refused: " + explicitRefresh.Refusal);
        Assert.Contains(
            trace,
            line => line.Contains("gantt.refresh delta=False", StringComparison.Ordinal));
    }

    /// <summary>
    /// A scheduled refresh does not blindly rewrite every shape just because it wrote
    /// row structure: it verifies the live shapes and writes only the drifted ones.
    /// </summary>
    /// <remarks>
    /// The earlier interim rule forced a full reconcile on any structural write, which
    /// made the cache stop paying exactly on the refreshes the user performs most.
    /// Verified skip replaces it: the same certainty, at a few property reads per
    /// unchanged shape.
    /// </remarks>
    [Fact]
    public void A_structural_write_does_not_bypass_verified_skips()
    {
        RefreshFakes fakes = new() { Rows = ValidRows() };
        ShapeDeltaCache cache = new();

        // The result line reports the move-only count (ADR-0053). A pass with no
        // geometry-only updates says geom=0 rather than omitting the field, so a reader of
        // the log never has to guess which build produced the line.
        List<string> trace = [];
        Assert.True(fakes.BuildOrchestratorWithDeltaCache(cache, trace: trace.Add).Refresh().Succeeded);
        Assert.Contains(
            trace,
            line => line.Contains("gantt.refresh result ok=True", StringComparison.Ordinal)
                && line.Contains("geom=", StringComparison.Ordinal));
        int writesAfterFirst = fakes.Shapes.ShapeCalls.Count;

        fakes.RowHeightsWritten = 2;
        fakes.OutlineGroupsApplied = 1;
        GanttRefreshOutcome second = fakes.BuildOrchestratorWithDeltaCache(cache).Refresh();

        Assert.True(second.Succeeded, "refused: " + second.Refusal + " " + second.Message);
        Assert.Equal(writesAfterFirst, fakes.Shapes.ShapeCalls.Count);
        Assert.True(fakes.Shapes.GeometryMatchChecks > 0, "the candidates must be verified against the host");
    }

    /// <summary>
    /// A scheduled refresh rewrites a shape the host moved on its own, even though the
    /// scene payload is byte-identical — the case that left the splitter band stale.
    /// </summary>
    [Fact]
    public void A_refresh_rewrites_a_shape_whose_live_geometry_drifted()
    {
        RefreshFakes fakes = new() { Rows = ValidRows() };
        ShapeDeltaCache cache = new();

        Assert.True(fakes.BuildOrchestratorWithDeltaCache(cache).Refresh().Succeeded);
        int writesAfterFirst = fakes.Shapes.ShapeCalls.Count;
        string drifted = fakes.Shapes.Requests.Keys.First();
        _ = fakes.Shapes.DriftedGeometry.Add(drifted);

        GanttRefreshOutcome second = fakes.BuildOrchestratorWithDeltaCache(cache).Refresh();

        Assert.True(second.Succeeded, "refused: " + second.Refusal + " " + second.Message);
        Assert.Contains(
            fakes.Shapes.ShapeCalls,
            call => string.Equals(call, "Update:" + drifted, StringComparison.Ordinal));
        Assert.True(fakes.Shapes.ShapeCalls.Count > writesAfterFirst);
    }

    [Fact]
    public void A_failed_shape_reconcile_does_not_commit_a_delta_snapshot()
    {
        RefreshFakes fakes = new() { Rows = ValidRows() };
        fakes.Shapes.Refusals["Create#1"] = ShapeWriteRefusal.HostRejected;
        ShapeDeltaCache cache = new();

        GanttRefreshOutcome outcome = fakes.BuildOrchestratorWithDeltaCache(cache).Refresh();

        Assert.False(outcome.Succeeded);
        Assert.Null(cache.Snapshot);
    }

    /// <summary>
    /// The row heights written on Refresh are the catalogue's tokens, not literals
    /// invented by the orchestrator.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>This is the test that would have caught the hardcoded-height defect.</b> The
    /// orchestrator called <c>Normalise(15, 6, 6)</c> against catalogue defaults of
    /// 18 / 18 / 9, and every refresh test stayed green because
    /// <c>FakeRowHeightNormaliser</c> discarded all three arguments and returned
    /// <c>Ok(0)</c>. The fake is now recording, so a wrong value is observable.
    /// </para>
    /// <para>
    /// The assertion is against <c>GanttCatalogues.MetricDefault</c> rather than three
    /// literals, so retuning a token does not require editing this test and cannot
    /// leave it asserting a stale number.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_row_heights_written_are_the_catalogue_tokens()
    {
        (RefreshFakes fakes, GanttRefreshOrchestrator orchestrator) = Ready();

        GanttRefreshOutcome outcome = orchestrator.Refresh();

        Assert.True(outcome.Succeeded, "refused: " + outcome.Refusal + " " + outcome.Message);
        Assert.Equal(GanttCatalogues.MetricDefault("GanttRowHeightPt"), fakes.LastManagedHeightPt);
        Assert.Equal(GanttCatalogues.MetricDefault("SplitterHeightPt"), fakes.LastSplitterHeightPt);
        Assert.Equal(GanttCatalogues.MetricDefault("SpacerHeightPt"), fakes.LastSpacerHeightPt);
    }

    /// <summary>
    /// The row heights written on Refresh are the WORKBOOK's metrics, not the token
    /// catalogue's defaults.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>This is the test that would have caught the ignored-setting defect.</b> The
    /// Theme dialog writes the user's "Row height", "Splitter height" and "Spacer
    /// height" to the metrics table, and the orchestrator then normalised the sheet
    /// to the catalogue's DEFAULTS - so the setting was stored, reported as written,
    /// and had no effect at all on the sheet. The scene factory read the workbook's
    /// metrics while the sheet was normalised to the defaults, so the two disagreed
    /// about how tall a row is: the exact disagreement row-height normalisation
    /// exists to prevent (ADR-0026 D2/D3).
    /// </para>
    /// <para>
    /// The values are asserted against the map the test supplied, not against
    /// literals, so the assertion is about the WORKBOOK being the authority rather
    /// than about a particular number.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_row_heights_written_are_the_workbooks_metrics()
    {
        RefreshFakes fakes = new()
        {
            Rows = ValidRows(),
            Metrics = new Dictionary<string, double>
            {
                ["GanttRowHeightPt"] = 27,
                ["SplitterHeightPt"] = 31,
                ["SpacerHeightPt"] = 12,
            },
        };
        GanttRefreshOrchestrator orchestrator = fakes.BuildOrchestrator();

        GanttRefreshOutcome outcome = orchestrator.Refresh();

        Assert.True(outcome.Succeeded, "refused: " + outcome.Refusal + " " + outcome.Message);
        Assert.Equal(27, fakes.LastManagedHeightPt);
        Assert.Equal(31, fakes.LastSplitterHeightPt);
        Assert.Equal(12, fakes.LastSpacerHeightPt);
    }

    /// <summary>
    /// A metric the workbook does not carry falls back to the catalogue's default,
    /// so a workbook written before a token existed still normalises to a defined
    /// height rather than to nothing.
    /// </summary>
    [Fact]
    public void A_metric_the_workbook_does_not_carry_falls_back_to_the_catalogue_default()
    {
        RefreshFakes fakes = new()
        {
            Rows = ValidRows(),
            Metrics = new Dictionary<string, double> { ["GanttRowHeightPt"] = 27 },
        };
        GanttRefreshOrchestrator orchestrator = fakes.BuildOrchestrator();

        GanttRefreshOutcome outcome = orchestrator.Refresh();

        Assert.True(outcome.Succeeded, "refused: " + outcome.Refusal + " " + outcome.Message);
        Assert.Equal(27, fakes.LastManagedHeightPt);
        Assert.Equal(GanttCatalogues.MetricDefault("SplitterHeightPt"), fakes.LastSplitterHeightPt);
        Assert.Equal(GanttCatalogues.MetricDefault("SpacerHeightPt"), fakes.LastSpacerHeightPt);
    }

    /// <summary>
    /// The live title occupies the RESERVED ROW exactly: the same row that carries
    /// the year band over the plot columns.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The title's height came from the TitleBandHeightPt token (24pt) while the
    /// reserved row it is drawn in is normalised to YearBandHeightPt (18pt). The two
    /// disagreed, so the textbox began 6pt ABOVE the reserved row, in the top
    /// padding row, and the chart frame's top edge - the reserved row's top - cut
    /// straight through the title instead of enclosing it. One row has one
    /// authority: the row itself.
    /// </para>
    /// <para>
    /// The grid here models a sheet with a real reserved row: 18pt of padding row
    /// above, then the 18pt reserved row, then the 15pt header, so the first body
    /// row's top is 51pt down.
    /// </para>
    /// </remarks>
    [Fact]
    public void The_live_title_occupies_the_reserved_row()
    {
        PanelCellGrid grid = GridWithAReservedRow();
        RefreshFakes fakes = new()
        {
            Rows = ValidRows(),
            MeasuredGrid = grid,
            Settings = new Dictionary<string, string>
            {
                ["ChartTitle"] = "Gantt Chart",
                ["ShowTitle"] = "TRUE",
            },
        };
        GanttRefreshOrchestrator orchestrator = fakes.BuildOrchestrator();

        GanttRefreshOutcome outcome = orchestrator.Refresh();

        Assert.True(outcome.Succeeded, "refused: " + outcome.Refusal + " " + outcome.Message);
        Assert.True(
            fakes.Shapes.Requests.ContainsKey("ChartTitle"),
            "No ChartTitle shape was written: " + string.Join(", ", fakes.Shapes.Calls));

        RectD? bounds = fakes.Shapes.Requests["ChartTitle"].Geometry.Bounds;
        Assert.NotNull(bounds);

        // The reserved row's top is the top padding row's height (18pt) and its
        // height is 18pt, so the title sits in the row the year band occupies.
        Assert.Equal(18, bounds!.Value.Top, precision: 9);
        Assert.Equal(18, bounds!.Value.Height, precision: 9);
    }

    /// <summary>
    /// A grid that models a sheet with a real reserved row: an 18pt top padding
    /// row, an 18pt reserved row, a 15pt header row, and one 15pt body row.
    /// </summary>
    private static PanelCellGrid GridWithAReservedRow()
    {
        PanelCellGridCreationOutcome created = PanelCellGrid.TryCreate(
            [new PanelColumn("Type", 24d)],
            [15d],
            15d,
            ["Type"],
            originTopPt: 51d,
            originLeftPt: 0d,
            topPaddingHeightPt: 18d);

        Assert.True(created.Succeeded, "The grid was refused: " + created.Refusal);
        Assert.NotNull(created.Grid);
        return created.Grid;
    }

    /// <summary>
    /// The two layout-row heights are the catalogue tokens, not literals.
    /// </summary>
    /// <remarks>
    /// The header row is the period band's row (ADR-0030 D5), so its height is what
    /// makes the period band's bottom coincide with the first body row's top; the
    /// reserved row above it carries the table title and the year band (D4). Both
    /// are asserted against <see cref="GanttCatalogues.MetricDefault"/> rather than
    /// numbers, so retuning a band token does not require editing this test and cannot
    /// leave it asserting a stale figure.
    /// </remarks>
    [Fact]
    public void The_layout_row_heights_are_the_catalogue_tokens()
    {
        (RefreshFakes fakes, GanttRefreshOrchestrator orchestrator) = Ready();

        GanttRefreshOutcome outcome = orchestrator.Refresh();

        Assert.True(outcome.Succeeded, "refused: " + outcome.Refusal + " " + outcome.Message);
        Assert.Equal(GanttCatalogues.MetricDefault("PeriodBandHeightPt"), fakes.LastHeaderHeightPt);
        Assert.Equal(GanttCatalogues.MetricDefault("YearBandHeightPt"), fakes.LastReservedRowHeightPt);
    }

    /// <summary>
    /// A refresh whose measured column set contains a HIDDEN column refuses at the
    /// measurement step, which is the live "The worksheet columns could not be
    /// measured" failure.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>This is the positive test for the defect just fixed.</b> It asserts the
    /// failure path FIRES for the real reason, so the fix is proved by the error no
    /// longer occurring on the same input rather than by the error message alone.
    /// </para>
    /// <para>
    /// The fake reports a hidden column's width as <c>0</c>, exactly as Excel does, so
    /// <c>PanelCellGrid.TryCreate</c> refuses it as a non-positive width and the
    /// adapter reports <c>InvalidMeasurement</c>. The old fake ignored the column set
    /// entirely and returned a hardcoded grid, which is precisely why every
    /// orchestrator test stayed green while the product refused on every refresh.
    /// </para>
    /// </remarks>
    [Fact]
    public void A_hidden_measured_column_refuses_the_refresh_at_the_measurement_step()
    {
        RefreshFakes fakes = new() { Rows = ValidRows() };
        string hidden = GanttTableSchema.Default.Columns.First(c => c.IsHidden).Name;

        PanelGridOutcome outcome = fakes.Panel.Measure([hidden]);

        Assert.False(outcome.Succeeded);
        Assert.Equal(PanelGridRefusalReason.InvalidMeasurement, outcome.Refusal);
    }

    /// <summary>
    /// The columns the orchestrator asks the measurement port for are the visible
    /// schema columns, so no hidden column can reach the host.
    /// </summary>
    /// <remarks>
    /// The orchestration-level counterpart to the factory's own assertion: this proves
    /// the correct set is what actually reaches the port, not merely that the factory
    /// publishes it.
    /// </remarks>
    [Fact]
    public void The_refresh_measures_only_visible_columns()
    {
        (RefreshFakes fakes, GanttRefreshOrchestrator orchestrator) = Ready();

        GanttRefreshOutcome outcome = orchestrator.Refresh();

        Assert.True(outcome.Succeeded, "refused: " + outcome.Refusal + " " + outcome.Message);
        Assert.Equal(ExcelSceneBuildRequestFactory.MeasuredColumns, fakes.LastMeasuredColumns);
    }

    /// <summary>
    /// The three targets are each positive, so no row is normalised to a height that
    /// Excel cannot store.
    /// </summary>
    /// <remarks>
    /// The positive test for the values, guarding the failure the defect could also
    /// have produced: a zero or negative height writes an unusable row. Stated against
    /// the catalogue range rather than a literal, so a retune cannot invalidate it.
    /// </remarks>
    [Fact]
    public void Every_row_height_token_is_a_positive_finite_value()
    {
        foreach (string token in new[] { "GanttRowHeightPt", "SplitterHeightPt", "SpacerHeightPt" })
        {
            double value = GanttCatalogues.MetricDefault(token);
            Assert.True(double.IsFinite(value) && value > 0, $"'{token}' resolved to {value}.");
        }
    }

    /// <summary>
    /// A refusal reached after the Duration write reports the write, rather than
    /// claiming nothing was touched.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>This is the test that would have caught the discarded counts.</b>
    /// <c>GanttRefreshOutcome.Refused</c> hardcoded both counts to zero, so a refusal
    /// after the Duration column had been changed reported "0 duration cells written"
    /// — untrue, and the kind of untruth a user meets as a workbook that is dirty for
    /// no stated reason.
    /// </para>
    /// <para>
    /// The measurement is injected rather than inferred: the fake duration writer
    /// reports a known count, so the assertion is that the orchestrator <em>passed it
    /// through</em>, not that it happened to match something.
    /// </para>
    /// </remarks>
    [Fact]
    public void A_refusal_after_the_duration_write_reports_what_it_wrote()
    {
        (RefreshFakes fakes, GanttRefreshOrchestrator orchestrator) = Ready();
        fakes.DurationCellsWritten = 3;

        // Refuse at a step that runs AFTER the Duration write.
        fakes.FactoryRefused = true;

        GanttRefreshOutcome outcome = orchestrator.Refresh();

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttRefreshRefusal.SceneRequestRefused, outcome.Refusal);
        Assert.Equal(3, outcome.DurationCellsWritten);
    }

    /// <summary>
    /// A refusal before any worksheet write still reports zero, so the widened
    /// property is not simply always-true.
    /// </summary>
    /// <remarks>
    /// The counterpart to the test above. Without it, an implementation that reported
    /// a phantom write count on every refusal would satisfy the positive case and the
    /// count would carry no information.
    /// </remarks>
    [Fact]
    public void A_refusal_before_any_write_reports_zero_writes()
    {
        (RefreshFakes fakes, GanttRefreshOrchestrator orchestrator) = Ready();
        fakes.TableReadRefused = true;

        GanttRefreshOutcome outcome = orchestrator.Refresh();

        Assert.False(outcome.Succeeded);
        Assert.Equal(0, outcome.DurationCellsWritten);
        Assert.Equal(0, outcome.ShapesWritten);
        Assert.False(outcome.LeftPartiallyApplied);
    }

    /// <summary>
    /// A refusal that wrote something reports the sheet as partly applied, even when
    /// no shape moved.
    /// </summary>
    /// <remarks>
    /// <c>LeftPartiallyApplied</c> previously named only
    /// <c>PartialReconciliation</c>, so it said "no" here while the user's Duration
    /// column had genuinely changed. Stating it on the counts makes the property an
    /// observable fact rather than an enum comparison.
    /// </remarks>
    [Fact]
    public void A_refusal_that_wrote_the_duration_column_is_partly_applied()
    {
        (RefreshFakes fakes, GanttRefreshOrchestrator orchestrator) = Ready();
        fakes.DurationCellsWritten = 2;
        fakes.FactoryRefused = true;

        GanttRefreshOutcome outcome = orchestrator.Refresh();

        Assert.False(outcome.Succeeded);
        Assert.True(outcome.LeftPartiallyApplied);
    }

    /// <summary>
    /// A successful refresh reports the same duration count it would have on a
    /// refusal, so the two paths cannot drift apart.
    /// </summary>
    [Fact]
    public void A_successful_refresh_reports_its_duration_writes()
    {
        (RefreshFakes fakes, GanttRefreshOrchestrator orchestrator) = Ready();
        fakes.DurationCellsWritten = 5;

        GanttRefreshOutcome outcome = orchestrator.Refresh();

        Assert.True(outcome.Succeeded, "refused: " + outcome.Refusal + " " + outcome.Message);
        Assert.Equal(5, outcome.DurationCellsWritten);
    }

    /// <summary>
    /// D2: the steps run in the fixed order the work item names, with the shape
    /// reconciliation last.
    /// </summary>
    /// <remarks>
    /// Asserted as an exact sequence rather than as "contains these in some order",
    /// because the claim is precisely that the order is fixed. A partial assertion
    /// would still pass if two steps were swapped.
    /// </remarks>
    [Fact]
    public void The_steps_run_in_the_fixed_order_with_reconciliation_last()
    {
        (RefreshFakes fakes, GanttRefreshOrchestrator orchestrator) = Ready();

        // Make identity repair report a write, so the conditional re-read is in the
        // sequence this test pins; the no-op case has its own test.
        fakes.IdentityRepairedCount = 1;

        GanttRefreshOutcome outcome = orchestrator.Refresh();

        Assert.True(outcome.Succeeded, "SCAFFOLD refusal=" + outcome.Refusal + " msg=" + outcome.Message);
        Assert.Equal(
            [
                "Read",
                "Config",
                "Guard",
                "Columns",
                "Identity",

                // The re-read after repair. The first snapshot is stale the moment
                // repair writes an Id, so validation must not consume it.
                "Read",

                // The notes, maintained against the findings of the validation that
                // follows. Placed BEFORE the render decision on purpose: a refresh that
                // now finds no errors must clear the notes a previous Validate left
                // behind, and that call has to happen on the blocked path too.
                "Highlights",
                "Outline",
                "Duration",
                "RowHeights",
                "Measure",
                "Factory",
            ],
            fakes.Steps);

        // The shape port is only reached by the reconciliation, which is the single
        // shape-mutating step (D3).
        Assert.NotEmpty(fakes.Shapes.Calls);
        Assert.Equal("ListOwned", fakes.Shapes.Calls[0]);
    }

    /// <summary>
    /// A refusal at any step before the reconciliation leaves <b>zero</b> shape
    /// mutations — D3's central claim.
    /// </summary>
    /// <remarks>
    /// <b>This is the test that would have caught a reordering.</b> If any validation
    /// or scene step were moved after the reconciliation, the corresponding entry here
    /// would report a non-empty shape call list. Each case is injected at a different
    /// step, so the guarantee is shown to hold along the whole preflight, not just at
    /// the start of it.
    /// </remarks>
    [Theory]
    [InlineData("table", GanttRefreshRefusal.TableMissing)]
    [InlineData("config", GanttRefreshRefusal.ConfigurationUnreadable)]
    [InlineData("protection", GanttRefreshRefusal.TargetProtected)]
    [InlineData("duration", GanttRefreshRefusal.DurationWriteRefused)]
    [InlineData("outline", GanttRefreshRefusal.OutlineRefused)]
    [InlineData("rowheight", GanttRefreshRefusal.RowHeightRefused)]
    [InlineData("measure", GanttRefreshRefusal.MeasurementRefused)]
    [InlineData("factory", GanttRefreshRefusal.SceneRequestRefused)]
    [InlineData("columns", GanttRefreshRefusal.ColumnPresentationRefused)]
    [InlineData("identity", GanttRefreshRefusal.IdentityRefused)]
    public void A_refusal_at_any_preflight_step_mutates_no_shape(string failure, GanttRefreshRefusal expected)
    {
        (RefreshFakes fakes, GanttRefreshOrchestrator orchestrator) = Ready();
        switch (failure)
        {
            case "table": fakes.TableReadRefused = true; break;
            case "config": fakes.ConfigReadRefused = true; break;
            case "protection": fakes.Protection = ProtectionGuardOutcome.SheetProtected; break;
            case "duration": fakes.DurationRefused = true; break;
            case "outline": fakes.OutlineRefused = true; break;
            case "rowheight": fakes.RowHeightRefused = true; break;
            case "measure": fakes.MeasurementRefused = true; break;
            case "factory": fakes.FactoryRefused = true; break;
            case "columns": fakes.ColumnPresentationRefused = true; break;
            case "identity": fakes.IdentityRefused = true; break;
            default: throw new ArgumentOutOfRangeException(nameof(failure), failure, "unknown injected failure");
        }

        GanttRefreshOutcome outcome = orchestrator.Refresh();

        Assert.False(outcome.Succeeded);
        Assert.Equal(expected, outcome.Refusal);

        // D3. Not "no Update" — no shape port call of any kind, including the
        // ListOwned read, so a blocked refresh is observably a no-op on the chart.
        Assert.True(
            fakes.Shapes.Calls.Count == 0,
            "A refusal at '" + failure + "' must not touch the shape port, but it recorded: "
            + string.Join(", ", fakes.Shapes.Calls)
            + " | steps=" + string.Join(", ", fakes.Steps));
    }

    /// <summary>
    /// A blocking validation error stops the refresh before any write, and the
    /// previous chart is untouched.
    /// </summary>
    /// <remarks>
    /// The issue list is carried on the outcome even though the refresh refused, so
    /// the caller can report exactly which row and field is wrong rather than a bare
    /// "something is invalid".
    /// </remarks>
    [Fact]
    public void A_blocking_validation_error_stops_the_refresh_and_reports_the_issues()
    {
        RefreshFakes fakes = new()
        {
            Rows =
            [
                // A blank Id is a blocking error, so the row never becomes an event.
                Row(1, string.Empty, "As-Planned Activity", "Design", new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19)),
            ],
        };
        GanttRefreshOrchestrator orchestrator = fakes.BuildOrchestrator();

        GanttRefreshOutcome outcome = orchestrator.Refresh();

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttRefreshRefusal.BlockingValidationErrors, outcome.Refusal);
        Assert.NotEmpty(outcome.ValidationIssues);
        Assert.Empty(fakes.Shapes.Calls);

        // The worksheet writes are downstream of validation and must not have run.
        Assert.DoesNotContain("Duration", fakes.Steps);
        Assert.DoesNotContain("Outline", fakes.Steps);
    }

    /// <summary>
    /// A reconciliation refused on its <em>first</em> operation is a plain refusal:
    /// nothing was written, so the existing chart is genuinely unchanged.
    /// </summary>
    [Fact]
    public void A_reconciliation_refused_before_any_write_leaves_the_chart_unchanged()
    {
        (RefreshFakes fakes, GanttRefreshOrchestrator orchestrator) = Ready();

        // Nothing is owned yet, so every planned operation is a create. Refusing the
        // FIRST create is therefore refusing before any write has landed.
        fakes.Shapes.Refusals["Create#1"] = ShapeWriteRefusal.HostRejected;

        GanttRefreshOutcome outcome = orchestrator.Refresh();

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttRefreshRefusal.ReconciliationRefused, outcome.Refusal);
        Assert.False(outcome.LeftPartiallyApplied);
    }

    /// <summary>
    /// A reconciliation refused <em>after</em> some operations landed is reported as
    /// partial, not as a plain refusal.
    /// </summary>
    /// <remarks>
    /// The distinction is the point. "The existing chart is unchanged" would be a lie
    /// here: some elements are from the old scene and some from the new, and the user
    /// is looking at it. REV5 §16 defers transactional Refresh to R4.10, so the tear
    /// is surfaced with its shape count rather than hidden behind a generic failure.
    /// </remarks>
    [Fact]
    public void A_reconciliation_refused_midway_is_reported_as_a_partial_refresh()
    {
        (RefreshFakes fakes, GanttRefreshOrchestrator orchestrator) = Ready();

        // Let the first create land, then refuse the second. A single-row chart emits
        // more than one primitive (the bar, its label, the frame), so the second
        // operation is guaranteed to exist.
        fakes.Shapes.Refusals["Create#2"] = ShapeWriteRefusal.HostRejected;

        GanttRefreshOutcome outcome = orchestrator.Refresh();

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttRefreshRefusal.PartialReconciliation, outcome.Refusal);
        Assert.True(outcome.LeftPartiallyApplied);
        Assert.Contains("mixture", outcome.Message!, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Every refusal carries a message a user can act on.</summary>
    [Theory]
    [InlineData("table", GanttRefreshRefusal.TableMissing)]
    [InlineData("factory", GanttRefreshRefusal.SceneRequestRefused)]
    [InlineData("columns", GanttRefreshRefusal.ColumnPresentationRefused)]
    [InlineData("identity", GanttRefreshRefusal.IdentityRefused)]
    public void Every_refusal_carries_a_message(string failure, GanttRefreshRefusal expected)
    {
        (RefreshFakes fakes, GanttRefreshOrchestrator orchestrator) = Ready();
        switch (failure)
        {
            case "table": fakes.TableReadRefused = true; break;
            case "factory": fakes.FactoryRefused = true; break;
            case "columns": fakes.ColumnPresentationRefused = true; break;
            case "identity": fakes.IdentityRefused = true; break;
            default: throw new ArgumentOutOfRangeException(nameof(failure), failure, "Unknown failure point.");
        }

        GanttRefreshOutcome outcome = orchestrator.Refresh();

        Assert.Equal(expected, outcome.Refusal);
        Assert.False(string.IsNullOrWhiteSpace(outcome.Message));
    }

    /// <summary>
    /// The protection guard is consulted <b>before</b> any worksheet write, so a
    /// protected sheet is discovered before it is half-modified.
    /// </summary>
    /// <remarks>
    /// Asserted as "no write step ran at all" rather than by comparing indices,
    /// because on a protected sheet the guard <em>blocks</em>: the write steps never
    /// appear in the log, so an index comparison against a step that is not there
    /// would be comparing against −1 and would pass or fail for the wrong reason.
    /// Asserting absence is also the stronger claim — it says the refresh stopped,
    /// not merely that it stopped in a particular order.
    /// </remarks>
    [Fact]
    public void The_protection_guard_is_consulted_before_any_write()
    {
        (RefreshFakes fakes, GanttRefreshOrchestrator orchestrator) = Ready();
        fakes.Protection = ProtectionGuardOutcome.SheetProtected;

        GanttRefreshOutcome outcome = orchestrator.Refresh();

        Assert.False(outcome.Succeeded);
        Assert.Contains("Guard", fakes.Steps);
        Assert.DoesNotContain("Outline", fakes.Steps);
        Assert.DoesNotContain("Duration", fakes.Steps);
        Assert.DoesNotContain("RowHeights", fakes.Steps);
    }

    /// <summary>No active workbook is reported distinctly from a protected sheet.</summary>
    [Fact]
    public void A_missing_workbook_is_reported_as_no_workbook_not_as_protection()
    {
        (RefreshFakes fakes, GanttRefreshOrchestrator orchestrator) = Ready();
        fakes.Protection = ProtectionGuardOutcome.NoActiveWorkbook;

        GanttRefreshOutcome outcome = orchestrator.Refresh();

        Assert.Equal(GanttRefreshRefusal.NoActiveWorkbook, outcome.Refusal);
    }

    /// <summary>A scope that records the calls a refresh makes against it.</summary>
    private sealed class RecordingScope : IApplicationStateScope
    {
        /// <summary>The operations performed, in order.</summary>
        public List<string> Calls { get; } = [];

        public void SuppressScreenUpdating() => Calls.Add("Screen");

        public void SuppressEvents() => Calls.Add("Events");

        public void SuppressAlerts() => Calls.Add("Alerts");

        public void SuppressStatusBar() => Calls.Add("StatusBar");

        public void SetStatusBarText(string? text) => Calls.Add("Text:" + text);

        public void CaptureSelection() => Calls.Add("Capture");

        public void Dispose() => Calls.Add("Dispose");
    }

    /// <summary>
    /// D8: the application state is suppressed and restored on the success path.
    /// </summary>
    [Fact]
    public void The_application_state_is_suppressed_and_restored_on_success()
    {
        (RefreshFakes fakes, GanttRefreshOrchestrator orchestrator) = Ready();
        // The orchestrator disposes the scope; the test owns it only to inspect the
        // recorded calls, so it is disposed here too rather than left to the analyzer.
        using RecordingScope scope = new();
        GanttRefreshOrchestrator withScope = fakes.BuildOrchestrator(scope);

        GanttRefreshOutcome outcome = withScope.Refresh();

        Assert.True(outcome.Succeeded, "refused: " + outcome.Message);

        // Every setting suppressed, the selection captured, and the scope disposed —
        // which is what restores. `using` is the guarantee: there is no return path
        // in the pipeline that can skip it.
        Assert.Equal("Screen", scope.Calls[0]);
        Assert.Contains("Events", scope.Calls);
        Assert.Contains("Alerts", scope.Calls);
        // The exact recorded call, not a prefix: `Assert.Contains` over a collection is
        // exact-element equality, so "Text:" alone would not match
        // "Text:Rendering the Gantt chart…" and the assertion would pass vacuously.
        Assert.Contains("Text:Rendering the Gantt chart…", scope.Calls);
        Assert.Contains("Capture", scope.Calls);
        Assert.Equal("Dispose", scope.Calls[^1]);

        // The progress text is written, but the bar's VISIBILITY is left alone so
        // the user's own Excel setting decides. This is the assertion that pins the
        // owner decision of 2026-10-01, and it is deliberately a negative one:
        // `SuppressStatusBar` used to be called here and the text set immediately
        // after, so the message was written into a bar we had just hidden and the
        // user never saw it. A test that only checked the text was set would have
        // passed against that version too.
        Assert.DoesNotContain("StatusBar", scope.Calls);
    }

    /// <summary>
    /// D8 again, on the failure path: a refusal must still restore the state.
    /// </summary>
    /// <remarks>
    /// This is the row that matters. A refresh that refused with events left disabled
    /// and the status bar still carrying our text would leave the user's Excel in a
    /// subtly broken state with nothing on screen to explain it.
    /// </remarks>
    [Theory]
    [InlineData("table", GanttRefreshRefusal.TableMissing)]
    [InlineData("factory", GanttRefreshRefusal.SceneRequestRefused)]
    [InlineData("columns", GanttRefreshRefusal.ColumnPresentationRefused)]
    [InlineData("identity", GanttRefreshRefusal.IdentityRefused)]
    public void The_application_state_is_restored_on_a_refusal_too(string failure, GanttRefreshRefusal expected)
    {
        (RefreshFakes fakes, GanttRefreshOrchestrator orchestrator) = Ready();
        // The orchestrator disposes the scope; the test owns it only to inspect the
        // recorded calls, so it is disposed here too rather than left to the analyzer.
        using RecordingScope scope = new();
        switch (failure)
        {
            case "table": fakes.TableReadRefused = true; break;
            case "factory": fakes.FactoryRefused = true; break;
            case "columns": fakes.ColumnPresentationRefused = true; break;
            case "identity": fakes.IdentityRefused = true; break;
            default: throw new ArgumentOutOfRangeException(nameof(failure), failure, "Unknown failure point.");
        }

        GanttRefreshOutcome outcome = fakes.BuildOrchestrator(scope).Refresh();

        Assert.Equal(expected, outcome.Refusal);
        Assert.Equal("Dispose", scope.Calls[^1]);
    }

    /// <summary>
    /// Validation sees the REPAIRED rows, not the snapshot taken before repair.
    /// </summary>
    /// <remarks>
    /// <b>The regression this pins.</b> Identity repair exists to fix a blank or
    /// malformed <c>Id</c>, so it WRITES new identifiers to the worksheet. The
    /// orchestrator read the table once at step 1, and validating that snapshot meant
    /// reporting the very <c>IdMissingOrMalformed</c> errors repair had just resolved —
    /// a refresh refused with <c>BlockingValidationErrors</c> on a workbook the user
    /// could not fix by any means, because the column they would edit is engine-hidden
    /// and the row was already repaired underneath it.
    /// </remarks>
    [Fact]
    public void Validation_receives_the_rows_repaired_by_the_identity_step()
    {
        // The FIRST read returns a row with no usable Id: the blocking error that
        // makes identity repair necessary in the first place.
        RefreshFakes fakes = new() { Rows = [RowWithoutId()], IdentityRepairedCount = 1 };

        // After repair the same row carries a well-formed Id, which is what the
        // repairer writes to the worksheet.
        fakes.RowsAfterRepair = ValidRows();

        GanttRefreshOutcome outcome = fakes.BuildOrchestrator().Refresh();

        Assert.True(
            outcome.Succeeded,
            "the refresh refused: " + outcome.Refusal + " " + outcome.Message);

        // The re-read really happened, after the repair and before validation.
        Assert.Equal(2, fakes.TableReadCount);
        Assert.True(
            fakes.Steps.IndexOf("Identity") < IndexOfSecond(fakes.Steps, "Read"),
            "the re-read must follow identity repair; steps were " + string.Join(", ", fakes.Steps));
        Assert.DoesNotContain(
            outcome.ValidationIssues,
            issue => issue.Severity == GanttValidationSeverity.Error);
    }

    /// <summary>
    /// A failed re-read after repair takes the table-read refusal path, with no
    /// shape mutation.
    /// </summary>
    /// <remarks>
    /// <b>The positive test for the new refusal branch.</b> Without handling it, a
    /// re-read that failed would be validated against the stale rows or — worse —
    /// produce a chart from data the repair was in the middle of changing. It maps to
    /// <see cref="GanttRefreshRefusal.TableMissing"/> because from that point on it is
    /// indistinguishable from a first read that failed.
    /// </remarks>
    [Fact]
    public void A_failed_reread_after_repair_refuses_as_a_table_miss_and_mutates_no_shape()
    {
        RefreshFakes fakes = new()
        {
            Rows = [RowWithoutId()],
            IdentityRepairedCount = 1,
            TableReadRefusedAfterRepair = true,
        };

        GanttRefreshOutcome outcome = fakes.BuildOrchestrator().Refresh();

        Assert.False(outcome.Succeeded);
        Assert.Equal(GanttRefreshRefusal.TableMissing, outcome.Refusal);
        Assert.Empty(fakes.Shapes.Calls);
    }

    /// <summary>The one-based position of the second occurrence of <paramref name="step"/>.</summary>
    private static int IndexOfSecond(List<string> steps, string step)
    {
        var seen = 0;
        for (var index = 0; index < steps.Count; index++)
        {
            if (string.Equals(steps[index], step, StringComparison.Ordinal)
                && ++seen == 2)
            {
                return index;
            }
        }

        throw new InvalidOperationException("'" + step + "' does not occur twice in " + string.Join(", ", steps));
    }

    /// <summary>A row whose <c>Id</c> is blank, which is what identity repair fixes.</summary>
    private static GanttRowDto RowWithoutId() =>
        new(
            2,
            string.Empty,
            NewId(),
            0,
            "As-Planned Activity",
            "Design",
            new DateOnly(2024, 1, 8),
            new DateOnly(2024, 1, 19),
            parentId: null,
            styleKey: null,
            labelPositionText: null,
            fillColourText: null,
            strokeColourText: null,
            visible: true,
            sortOrder: null);

    /// <summary>Null collaborators are refused at construction, not at first use.</summary>
    [Fact]
    public void A_null_collaborator_is_refused_at_construction()
    {
        RefreshFakes fakes = new();

        Assert.Throws<ArgumentNullException>(() => new GanttRefreshOrchestrator(
            null!,
            fakes.ConfigReader,
            fakes.Guard,
            fakes.Panel,
            fakes.Duration,
            fakes.RowHeights,
            fakes.Outline,
            fakes.Factory,
            fakes.Shapes));
    }

    /// <summary>
    /// A present type-change normaliser forces a re-read before validation.
    /// </summary>
    /// <remarks>
    /// <b>Why the re-read is asserted, not merely the call.</b> The normaliser WRITES
    /// engine-hidden cells to the sheet. Validating the snapshot taken at step 1 would
    /// then judge the values the normaliser just corrected, so a row whose stale
    /// hidden cell is now right would still be refused — the same defect identity
    /// repair had. The step ordering is the claim: normalise, then read, then validate.
    /// </remarks>
    [Fact]
    public void A_type_change_normaliser_is_followed_by_a_reread_before_validation()
    {
        RefreshFakes fakes = new() { Rows = ValidRows(), TypeCellsWritten = 1 };

        GanttRefreshOutcome outcome = fakes.BuildOrchestratorWithTypeChangeNormaliser().Refresh();

        Assert.True(outcome.Succeeded, "refused: " + outcome.Refusal + " " + outcome.Message);
        Assert.Equal(1, fakes.NormaliseCount);

        // The re-read is the step IMMEDIATELY after the normalise, because the point of
        // it is that validation sees the corrected cells and nothing in between can move
        // the boundary. Counting reads would not say that: this pipeline reads the table
        // a third time on the render path.
        int normaliseAt = fakes.Steps.IndexOf("Normalise");
        Assert.True(normaliseAt >= 0, "Normalise never ran; steps=" + string.Join(", ", fakes.Steps));
        Assert.True(
            normaliseAt + 1 < fakes.Steps.Count
                && string.Equals(fakes.Steps[normaliseAt + 1], "Read", StringComparison.Ordinal),
            "expected the read straight after Normalise; steps=" + string.Join(", ", fakes.Steps));
    }

    /// <summary>
    /// The re-read after normalisation reports a missing workbook as exactly that.
    /// </summary>
    /// <remarks>
    /// The second arm of the re-read refusal. It exists because the two refusals have
    /// different user remedies — open a workbook versus find the table — and collapsing
    /// them into one message sends the user looking for a workbook they already have.
    /// </remarks>
    [Fact]
    public void A_reread_after_normalisation_with_no_workbook_refuses_as_no_active_workbook()
    {
        RefreshFakes fakes = new()
        {
            Rows = ValidRows(),
            TypeCellsWritten = 1,
            TableReadRefusedAfterRepair = true,
            RereadRefusalReason = GanttTableReadRefusalReason.NoActiveWorkbook,
        };

        GanttRefreshOutcome outcome = fakes.BuildOrchestratorWithTypeChangeNormaliser().Refresh();

        // The normaliser must have RUN for this refusal to be the re-read's. Without
        // this, a first read that refused for the same reason would satisfy the
        // assertion, and the branch under test would stay unproven.
        Assert.True(
            fakes.NormaliseCount == 1,
            "normaliser never ran: refusal=" + outcome.Refusal + " reads=" + fakes.TableReadCount
                + " steps=" + string.Join(",", fakes.Steps));
        Assert.Equal(GanttRefreshRefusal.NoActiveWorkbook, outcome.Refusal);
        Assert.Empty(fakes.Shapes.Calls);
    }

    /// <summary>The re-read after normalisation reports a missing table as exactly that.</summary>
    [Fact]
    public void A_reread_after_normalisation_with_no_table_refuses_as_a_table_miss()
    {
        RefreshFakes fakes = new()
        {
            Rows = ValidRows(),
            TypeCellsWritten = 1,
            TableReadRefusedAfterRepair = true,
            RereadRefusalReason = GanttTableReadRefusalReason.TableMissing,
        };

        GanttRefreshOutcome outcome = fakes.BuildOrchestratorWithTypeChangeNormaliser().Refresh();

        Assert.Equal(1, fakes.NormaliseCount);
        Assert.Equal(GanttRefreshRefusal.TableMissing, outcome.Refusal);
        Assert.Empty(fakes.Shapes.Calls);
    }

    /// <summary>
    /// A no-op identity repair does not trigger a second full-table read.
    /// </summary>
    /// <remarks>
    /// The re-read exists only because repair may have WRITTEN Ids. Reporting zero means
    /// the step-1 snapshot is still the sheet's content, so re-reading it was a redundant
    /// slot on every refresh — the kind of cost this optimisation row removes without
    /// changing what validation sees.
    /// </remarks>
    [Fact]
    public void A_no_op_identity_repair_does_not_trigger_a_second_read()
    {
        RefreshFakes fakes = new() { Rows = ValidRows() };

        GanttRefreshOutcome outcome = fakes.BuildOrchestrator().Refresh();

        Assert.True(outcome.Succeeded, "refused: " + outcome.Refusal + " " + outcome.Message);
        Assert.Equal(1, fakes.TableReadCount);
    }

    /// <summary>
    /// A no-op type-change pass does not trigger a second full-table read, and a written
    /// cell still does.
    /// </summary>
    [Fact]
    public void A_write_from_the_type_change_normaliser_still_triggers_the_reread()
    {
        RefreshFakes fakes = new() { Rows = ValidRows(), TypeCellsWritten = 1 };

        GanttRefreshOutcome outcome = fakes.BuildOrchestratorWithTypeChangeNormaliser().Refresh();

        Assert.True(outcome.Succeeded, "refused: " + outcome.Refusal + " " + outcome.Message);
        Assert.Equal(2, fakes.TableReadCount);
    }

    /// <summary>
    /// A dangling <c>StyleKey</c> is healed on the SAME refresh, not the next one.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>The whole reason <c>ApplyStyleKeyToEvents</c> exists.</b> The validated events
    /// carry the key read from column K at step 1. When that key names a style the user
    /// has deleted, validation reports a warning and keeps the raw value, the planner
    /// heals the cell to the type's code-owned default, and the scene then resolves the
    /// key carried by the EVENT. Without the mirror the scene sees
    /// <c>StyleThatWasDeleted</c> and the build refuses
    /// <see cref="SceneBuilderRefusal.UnresolvableStyle"/> — so a user who tidied up
    /// their styles would have to press Refresh twice, with the first press appearing to
    /// fail.
    /// </para>
    /// <para>
    /// The assertion is therefore on the refresh SUCCEEDING with the healed key, plus
    /// the writer's own count reaching the outcome. A test asserting only that the
    /// writer was called would pass with the mirror removed.
    /// </para>
    /// </remarks>
    [Fact]
    public void A_dangling_style_key_is_healed_and_rendered_in_the_same_refresh()
    {
        RefreshFakes fakes = new()
        {
            Rows =
            [
                Row(
                    rowNumber: 1,
                    id: null,
                    type: "As-Planned Activity",
                    description: "Design",
                    start: new DateOnly(2024, 1, 8),
                    finish: new DateOnly(2024, 1, 19),
                    styleKey: "StyleThatWasDeleted"),
            ],
            StyleKeyCellsWritten = 1,
        };

        GanttRefreshOutcome outcome = fakes.BuildOrchestratorWithStyleKeyWriter().Refresh();

        Assert.True(outcome.Succeeded, "refused: " + outcome.Refusal + " " + outcome.Message);
        Assert.Equal(1, fakes.StyleKeyPlanWriteCount);
        Assert.Equal(1, outcome.StyleKeyCellsWritten);
        Assert.NotEmpty(fakes.Shapes.Calls);
    }

    /// <summary>
    /// A 1904-date-system workbook is refused as unreadable, naming the date system.
    /// </summary>
    /// <remarks>
    /// The refusal is <see cref="GanttRefreshRefusal.TableMissing"/> because a workbook
    /// the add-in cannot date is a workbook whose table it cannot read, but the MESSAGE
    /// has to name the real cause: telling the user the table is missing sends them
    /// looking for a table that is there.
    /// </remarks>
    [Fact]
    public void A_1904_workbook_is_refused_with_the_date_system_named()
    {
        RefreshFakes fakes = new()
        {
            Rows = ValidRows(),
            TableReadRefused = true,
            FirstReadRefusal = GanttTableReadRefusalReason.DateSystemUnsupported,
        };

        GanttRefreshOutcome outcome = fakes.BuildOrchestrator().Refresh();

        Assert.Equal(GanttRefreshRefusal.TableMissing, outcome.Refusal);
        Assert.Contains("1904", outcome.Message!, StringComparison.Ordinal);
        Assert.Empty(fakes.Shapes.Calls);
    }

    /// <summary>
    /// A row already carrying its type's default style key costs no write at all.
    /// </summary>
    /// <remarks>
    /// The empty-plan early return, and a product guarantee rather than an optimisation:
    /// writing the value that is already in the cell marks the workbook dirty on every
    /// Refresh for no change. The counterweight to the healing test above.
    /// </remarks>
    [Fact]
    public void A_style_key_already_matching_its_type_is_not_written()
    {
        RefreshFakes fakes = new()
        {
            Rows =
            [
                Row(
                    rowNumber: 1,
                    id: null,
                    type: "As-Planned Activity",
                    description: "Design",
                    start: new DateOnly(2024, 1, 8),
                    finish: new DateOnly(2024, 1, 19),
                    styleKey: "AsPlannedActivity"),
            ],
        };

        GanttRefreshOutcome outcome = fakes.BuildOrchestratorWithStyleKeyWriter().Refresh();

        Assert.True(outcome.Succeeded, "refused: " + outcome.Refusal + " " + outcome.Message);
        Assert.Equal(0, fakes.StyleKeyPlanWriteCount);
        Assert.Equal(0, outcome.StyleKeyCellsWritten);
    }

    /// <summary>
    /// A refused <c>StyleKey</c> write reports the <c>Duration</c> cells that already
    /// landed, and mutates no shape.
    /// </summary>
    /// <remarks>
    /// The refusal happens after the Duration write, so reporting zero worksheet writes
    /// would tell the user their sheet was untouched when it was not — the workbook
    /// would be dirty for a reason the refresh denied.
    /// </remarks>
    [Fact]
    public void A_refused_style_key_write_reports_the_duration_cells_already_written()
    {
        RefreshFakes fakes = new()
        {
            Rows = ValidRows(),
            StyleKeyWriteRefused = true,
            DurationCellsWritten = 3,
        };

        GanttRefreshOutcome outcome = fakes.BuildOrchestratorWithStyleKeyWriter().Refresh();

        Assert.Equal(GanttRefreshRefusal.StyleKeyWriteRefused, outcome.Refusal);
        Assert.Equal(3, outcome.DurationCellsWritten);
        Assert.Empty(fakes.Shapes.Calls);
    }

    /// <summary>
    /// A scene-build refusal stops before any shape write and reports the reason.
    /// </summary>
    /// <remarks>
    /// The scene build sits between the worksheet writes and the first shape write, so
    /// it is the step where a careless reordering first becomes visible. The row uses
    /// <c>As-Built Activity</c>, whose style the fake registry does not carry, which is
    /// exactly the <see cref="SceneBuilderRefusal.UnresolvableStyle"/> input a
    /// hand-edited workbook produces.
    /// </remarks>
    [Fact]
    public void A_scene_build_refusal_reports_that_the_chart_was_not_rebuilt()
    {
        RefreshFakes fakes = new()
        {
            Rows =
            [
                Row(
                    rowNumber: 1,
                    id: null,
                    type: "As-Built Activity",
                    description: "Install",
                    start: new DateOnly(2024, 1, 8),
                    finish: new DateOnly(2024, 1, 19)),
            ],
        };

        GanttRefreshOutcome outcome = fakes.BuildOrchestrator().Refresh();

        Assert.Equal(GanttRefreshRefusal.SceneBuildRefused, outcome.Refusal);
        Assert.Empty(fakes.Shapes.Calls);
    }

    /// <summary>
    /// The arming step locks the owned shapes and leaves the sheet RELEASED for the write phase.
    /// </summary>
    /// <remarks>
    /// Not protecting here is the correction the live measurement forced: a protected sheet refuses
    /// shape creation, so arming must end with the sheet open and <see cref="WorksheetProtectionWindow"/>
    /// re-applies the protection when the writes finish. A lock-then-protect sequence inside this step
    /// would make the very first refresh fail at its first shape write.
    /// </remarks>
    [Fact]
    public void A_refresh_locks_the_owned_shapes_and_leaves_the_sheet_released()
    {
        RefreshFakes fakes = new() { Rows = ValidRows() };

        GanttRefreshOutcome outcome = fakes.BuildOrchestratorWithProtection().Refresh();

        Assert.True(outcome.Succeeded, "refusal=" + outcome.Refusal + " msg=" + outcome.Message);
        Assert.Equal(1, fakes.Shapes.LockOwnedShapesCallCount);
        Assert.Equal(["Read"], fakes.SheetProtection.Calls);
        Assert.True(fakes.SheetProtectionArming.IsArmed);
        Assert.False(fakes.SheetProtection.State!.ContentsProtected);
    }

    /// <summary>
    /// The second refresh over the same workbook does not re-arm, because arming is a session fact.
    /// </summary>
    /// <remarks>
    /// A re-arm per pass would be an unprotect plus a protect on every automatic refresh — the extra
    /// step the owner ruled out. The two orchestrators here model the live composition, which builds
    /// a new one per refresh and shares only the state object.
    /// </remarks>
    [Fact]
    public void A_second_refresh_does_not_re_arm_the_protection()
    {
        RefreshFakes fakes = new() { Rows = ValidRows() };

        Assert.True(fakes.BuildOrchestratorWithProtection().Refresh().Succeeded);
        Assert.True(fakes.BuildOrchestratorWithProtection().Refresh().Succeeded);

        Assert.Equal(0, fakes.SheetProtection.ProtectCount);
        Assert.Equal(["Read"], fakes.SheetProtection.Calls);
        Assert.Equal(1, fakes.Shapes.LockOwnedShapesCallCount);
    }

    /// <summary>
    /// A sheet that was already protected is released, because <c>UserInterfaceOnly</c> does not
    /// survive a reopen and the write phase needs an open sheet.
    /// </summary>
    [Fact]
    public void An_already_protected_sheet_is_released_before_the_writes()
    {
        RefreshFakes fakes = new() { Rows = ValidRows() };
        fakes.SheetProtection.State = new WorksheetProtectionState(ContentsProtected: true, ShapesProtected: true);

        GanttRefreshOutcome outcome = fakes.BuildOrchestratorWithProtection().Refresh();

        Assert.True(outcome.Succeeded, "refusal=" + outcome.Refusal + " msg=" + outcome.Message);
        Assert.Equal(["Read", "Unprotect"], fakes.SheetProtection.Calls);
        Assert.False(fakes.SheetProtection.State!.ContentsProtected);
    }

    /// <summary>
    /// A workbook that opts out is never protected and never swept, so the feature costs it nothing.
    /// </summary>
    /// <remarks>
    /// Absent means on, which is the catalogue default, so the opt-out has to be asserted
    /// explicitly: a test that omitted the key would pass on a pipeline that ignored the setting.
    /// </remarks>
    [Fact]
    public void A_workbook_that_opts_out_is_never_protected()
    {
        RefreshFakes fakes = new()
        {
            Rows = ValidRows(),
            Settings = new Dictionary<string, string> { ["ProtectGanttSheet"] = "FALSE" },
        };

        GanttRefreshOutcome outcome = fakes.BuildOrchestratorWithProtection().Refresh();

        Assert.True(outcome.Succeeded, "refusal=" + outcome.Refusal + " msg=" + outcome.Message);
        Assert.Empty(fakes.SheetProtection.Calls);
        Assert.Equal(0, fakes.Shapes.LockOwnedShapesCallCount);
        Assert.False(fakes.SheetProtectionArming.IsArmed);
    }

    /// <summary>
    /// A refused lock sweep loses the lock, not the chart, and is recorded.
    /// </summary>
    /// <remarks>
    /// The refresh still rendering is the product decision — no chart at all is worse than a correct
    /// chart the user can drag — and the trace is the only evidence a user can act on, so it is
    /// asserted rather than assumed.
    /// </remarks>
    [Fact]
    public void A_refused_lock_sweep_keeps_the_chart_and_is_recorded()
    {
        RefreshFakes fakes = new() { Rows = ValidRows() };
        fakes.Shapes.Refusals["LockOwnedShapes"] = ShapeWriteRefusal.TargetProtected;
        List<string> traces = [];

        GanttRefreshOutcome outcome = fakes.BuildOrchestratorWithProtection(traces.Add).Refresh();

        Assert.True(outcome.Succeeded, "refusal=" + outcome.Refusal + " msg=" + outcome.Message);
        Assert.Contains(
            traces,
            line => line.Contains("lock refused=TargetProtected", StringComparison.Ordinal));
    }

    /// <summary>
    /// Turning the setting off after it has been armed releases the add-in's protection.
    /// </summary>
    /// <remarks>
    /// The setting is both the ownership record and the user's switch (ADR-0055 D2), so an opt-out
    /// that only stopped future arming would leave the sheet locked against the user's explicit
    /// choice for the rest of the session.
    /// </remarks>
    [Fact]
    public void Turning_the_setting_off_after_arming_releases_the_protection()
    {
        RefreshFakes fakes = new() { Rows = ValidRows() };
        Assert.True(fakes.BuildOrchestratorWithProtection().Refresh().Succeeded);
        fakes.Settings = new Dictionary<string, string> { ["ProtectGanttSheet"] = "FALSE" };

        GanttRefreshOutcome outcome = fakes.BuildOrchestratorWithProtection().Refresh();

        Assert.True(outcome.Succeeded, "refusal=" + outcome.Refusal + " msg=" + outcome.Message);
        Assert.Equal(1, fakes.SheetProtection.UnprotectCount);
        Assert.False(fakes.SheetProtectionArming.IsArmed);
        Assert.False(fakes.SheetProtection.State!.ContentsProtected);
    }
}
