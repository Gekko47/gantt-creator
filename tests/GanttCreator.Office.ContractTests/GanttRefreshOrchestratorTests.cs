using GanttCreator.Core;

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
    /// <summary>One valid planned activity, enough for a scene to be built.</summary>
    private static GanttRowDto[] ValidRows() =>
    [
        Row(
            rowNumber: 2,
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
    private static GanttRowDto Row(
        int rowNumber,
        string? id,
        string type,
        string description,
        DateOnly? start,
        DateOnly? finish) =>
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
            styleKey: null,
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

        GanttRefreshOutcome outcome = orchestrator.Refresh();

        Assert.True(outcome.Succeeded, "SCAFFOLD refusal=" + outcome.Refusal + " msg=" + outcome.Message);
        Assert.Equal(
            [
                "Read",
                "Config",
                "Guard",
                "Columns",
                "Identity",
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
                Row(2, string.Empty, "As-Planned Activity", "Design", new DateOnly(2024, 1, 8), new DateOnly(2024, 1, 19)),
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
}
