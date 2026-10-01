using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office;

/// <summary>
/// The pipeline-shaped <see cref="IGanttRefreshOrchestrator"/>: it composes the
/// already-landed R4.7 and R4.8 contracts in the fixed order R4.8A D2 requires.
/// </summary>
/// <remarks>
/// <para>
/// <b>The order is the product, not an implementation detail.</b> Validation must
/// precede mutation, the plot must be resolved before the scene is built, and the
/// scene must be complete before a single shape moves. Each of those is a rule a user
/// would notice breaking — a half-rendered chart, a chart sized against stale rows —
/// so the sequence lives in one named type rather than in a Ribbon callback whose
/// ordering is a property of how it was written.
/// </para>
/// <para>
/// <b>What "before any shape mutation" means concretely.</b> Steps 1-13 below issue
/// no shape write at all: they read, validate, plan, and build. The
/// <see cref="IShapeWritePort"/> is first touched by
/// <see cref="ShapeReconciler.Reconcile"/> in the final step, and its
/// <c>ListOwned</c> read is itself skipped when the refresh is blocked. A test can
/// therefore inject a failure at any step and assert the shape port recorded zero
/// calls — which is D3, and it is only assertable because every collaborator is a
/// port.
/// </para>
/// <para>
/// <b>The two writes that are not shape mutations.</b> <c>Duration</c> and the outline
/// are worksheet writes that happen before the reconciliation, because both are
/// derived from validated rows and both must be current before the geometry is
/// measured. They are reported separately from the shape count so a caller can tell
/// "the chart was not touched" from "the sheet was not touched at all".
/// </para>
/// <para>
/// <b>Partial refresh is reported, not rolled back.</b> REV5 §16 defers transactional
/// Refresh to R4.10, so a mid-reconciliation refusal is surfaced as
/// <see cref="GanttRefreshRefusal.PartialReconciliation"/> with the count of
/// operations that did land. Pretending a partial write did not happen would be worse
/// than the tear itself.
/// </para>
/// </remarks>
public sealed class GanttRefreshOrchestrator(
    IGanttTableReader tableReader,
    IConfigCatalogueReader configReader,
    IWorksheetProtectionGuard protectionGuard,
    IPanelGridMeasurementPort panelMeasurement,
    IDurationWritePort durationWriter,
    IRowHeightNormalisationPort rowHeightNormaliser,
    IOutlineGroupPort outlineWriter,
    ISceneBuildRequestFactory requestFactory,
    IShapeWritePort shapeWriter,
    IApplicationStateScope? stateScope = null)
    : IGanttRefreshOrchestrator
{
    private readonly IGanttTableReader _tableReader = tableReader ?? throw new ArgumentNullException(nameof(tableReader));
    private readonly IConfigCatalogueReader _configReader = configReader ?? throw new ArgumentNullException(nameof(configReader));
    private readonly IWorksheetProtectionGuard _protectionGuard = protectionGuard ?? throw new ArgumentNullException(nameof(protectionGuard));
    private readonly IPanelGridMeasurementPort _panelMeasurement = panelMeasurement ?? throw new ArgumentNullException(nameof(panelMeasurement));
    private readonly IDurationWritePort _durationWriter = durationWriter ?? throw new ArgumentNullException(nameof(durationWriter));
    private readonly IRowHeightNormalisationPort _rowHeightNormaliser = rowHeightNormaliser ?? throw new ArgumentNullException(nameof(rowHeightNormaliser));
    private readonly IOutlineGroupPort _outlineWriter = outlineWriter ?? throw new ArgumentNullException(nameof(outlineWriter));
    private readonly ISceneBuildRequestFactory _requestFactory = requestFactory ?? throw new ArgumentNullException(nameof(requestFactory));
    private readonly IShapeWritePort _shapeWriter = shapeWriter ?? throw new ArgumentNullException(nameof(shapeWriter));

    /// <summary>
    /// The application-state scope, or <see langword="null"/> for the no-op scope.
    /// </summary>
    /// <remarks>
    /// Optional rather than required, so a caller with no Excel application — and a
    /// test that is not about restoration — does not have to construct one.
    /// </remarks>
    private readonly IApplicationStateScope? _stateScope = stateScope;

    /// <summary>The scene-to-shape renderer, held as a field so a test can observe
    /// that the orchestrator owns the translation step rather than a caller.</summary>
    private readonly SceneShapeRenderer _renderer = new(ChartOriginDelta.Identity);

    /// <inheritdoc />
    public GanttRefreshOutcome Refresh()
    {
        // D8: the five application settings and the selection are captured and
        // restored on EVERY path, success or failure, by the ADR-0020 scope. A
        // refresh turns off screen updating and events for speed and then adds
        // hundreds of shapes; if it refused half way, the user would be left with
        // events disabled and a status bar still carrying our text, with no way to
        // tell that anything is wrong. The `using` is what makes that guarantee
        // total: every return below runs Dispose, so there is no path that can
        // skip the restore.
        using IApplicationStateScope scope = _stateScope ?? NullApplicationStateScope.Instance;

        scope.SuppressScreenUpdating();
        scope.SuppressEvents();
        scope.SuppressAlerts();
        scope.SuppressStatusBar();
        scope.SetStatusBarText("Rendering the Gantt chart…");
        scope.CaptureSelection();

        return RunRefresh();
    }

    /// <summary>
    /// The pipeline itself, with the application state already suppressed.
    /// </summary>
    /// <returns>The typed outcome.</returns>
    private GanttRefreshOutcome RunRefresh()
    {
        // Step 1-2. Read the user's rows and the engine's configuration. Both are
        // read-only, so a refusal here has mutated nothing at all.
        GanttTableReadOutcome rows = _tableReader.Read();
        if (!rows.Succeeded)
        {
            return Refuse(GanttRefreshRefusal.TableMissing, Describe(rows.Refusal!.Value));
        }

        ConfigReadOutcome config = _configReader.Read();
        if (!config.Succeeded)
        {
            return Refuse(
                GanttRefreshRefusal.ConfigurationUnreadable,
                "The Gantt Creator configuration could not be read from the hidden configuration sheet.");
        }

        // Step 3. The mutation boundary, consulted before any write. A protected
        // worksheet would otherwise be discovered halfway through, leaving a
        // half-written sheet behind.
        ProtectionGuardOutcome protection = _protectionGuard.Query();
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return Refuse(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? GanttRefreshRefusal.NoActiveWorkbook
                    : GanttRefreshRefusal.TargetProtected,
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? "There is no active workbook to refresh."
                    : "The Gantt worksheet is protected, so the chart cannot be updated. Unprotect it and try again.");
        }

        // Step 4-5. Validate every row. A blocking error means the refresh is not
        // attempted at all; the previous chart stays exactly as it was.
        GanttValidationOutcome validation = GanttRowValidator.Validate(rows.Rows);
        if (!validation.IsValid)
        {
            return GanttRefreshOutcome.Refused(
                GanttRefreshRefusal.BlockingValidationErrors,
                "The table has errors, so the chart was not rebuilt. The existing chart is unchanged.",
                validation.Issues);
        }

        // Step 6. Worksheet presentation. The outline is rebuilt from the validated
        // hierarchy, never read back, so collapsing a group stays presentation.
        OutlineGroupOutcome outline = _outlineWriter.Apply(validation.Events);
        if (!outline.Succeeded)
        {
            return Refuse(GanttRefreshRefusal.OutlineRefused, "The row outline groups could not be updated.");
        }

        // Step 7. Duration. Calculated in Core and written in one bulk pass; an
        // already-correct column produces no write at all.
        DurationWritePlan plan = DurationCalculator.Plan(
            validation.Events,
            ReadCurrentDurations(rows.Rows));
        DurationWriteOutcome duration = _durationWriter.Write(plan);
        if (!duration.Succeeded)
        {
            return Refuse(GanttRefreshRefusal.DurationWriteRefused, "The Duration column could not be written.");
        }

        // Step 8. Row heights, so the measured lane geometry is meaningful.
        //
        // The three targets come from the token catalogue rather than from literals
        // written here. This call previously passed (15, 6, 6) against catalogue
        // defaults of 18 / 18 / 9, so the sheet was normalised to a height the tokens
        // do not describe — and because R4.7D removed lane auto-growth and derives the
        // lane height from the measured row, the worksheet and the chart were being
        // driven to disagree by the row that exists to stop them disagreeing
        // (ADR-0026 D2/D3).
        RowHeightNormalisationOutcome heights = _rowHeightNormaliser.Normalise(
            GanttCatalogues.MetricDefault("GanttRowHeightPt"),
            GanttCatalogues.MetricDefault("SplitterHeightPt"),
            GanttCatalogues.MetricDefault("SpacerHeightPt"));
        if (!heights.Succeeded)
        {
            return Refuse(GanttRefreshRefusal.RowHeightRefused, "The row heights could not be normalised.");
        }

        // Step 9-10. Measure the live panel, then resolve the scene inputs. The
        // measurement is a read; the factory performs no host call of its own.
        PanelGridOutcome grid = _panelMeasurement.Measure(ExcelSceneBuildRequestFactory.MeasuredColumns);
        if (!grid.Succeeded)
        {
            return Refuse(GanttRefreshRefusal.MeasurementRefused, "The worksheet columns could not be measured.");
        }

        // The style registry comes from the configuration read, not from
        // GanttCatalogues. The config sheet is the single source of the effective
        // named styles — including any user-authored ones — and rebuilding a
        // registry from the built-in catalogue here would silently drop them, so a
        // custom style would validate and then fail to resolve at render time.
        SceneBuildRequestOutcome request = _requestFactory.Create(
            validation.Events,
            config.Settings,
            config.Styles,
            grid.Grid!);
        if (!request.Succeeded)
        {
            return Refuse(GanttRefreshRefusal.SceneRequestRefused, request.Message!);
        }

        // Step 11. Build the scene. A refusal here is still before any shape write.
        SceneBuildOutcome scene = SceneBuilder.TryBuild(request.Request);
        if (!scene.Succeeded)
        {
            return Refuse(
                GanttRefreshRefusal.SceneBuildRefused,
                "The chart could not be composed: " + scene.Refusal + ".");
        }

        // Step 12. Translate the scene into host shape requests.
        SceneTranslationOutcome translated = _renderer.Translate(scene.Result!.Scene);
        if (translated.Refusals.Count > 0)
        {
            return Refuse(
                GanttRefreshRefusal.TranslationRefused,
                translated.Refusals.Count + " chart element(s) could not be drawn, so the chart was not rebuilt.");
        }

        // Step 13. The first and only shape mutation of the whole pipeline. Every
        // validation, measurement, and scene construction above has completed; from
        // here the reconciliation stops at the first refusal (D3).
        ShapeReconcileOutcome reconcile = ShapeReconciler.Reconcile(
            _shapeWriter,
            new ShapeReconcileRequest(translated.Requests, HasBlockingErrors: false));

        if (reconcile.Succeeded)
        {
            return GanttRefreshOutcome.Ok(validation.Issues, reconcile.CompletedCount, duration.CellsWritten);
        }

        // A reconciliation that stopped after at least one write leaves the sheet in
        // a state the user can see: part old scene, part new. That is reported
        // distinctly rather than as a plain refusal, because "the existing chart is
        // unchanged" would be false. REV5 defers transactional Refresh to R4.10, so
        // the tear is surfaced, not undone.
        return reconcile.CompletedCount > 0
            ? GanttRefreshOutcome.Refused(
                GanttRefreshRefusal.PartialReconciliation,
                "The chart was partly updated before Excel refused a change ("
                    + reconcile.CompletedCount
                    + " element(s) written), so it may show a mixture of the old and new layout. Refresh again to finish.",
                validation.Issues)
            : GanttRefreshOutcome.Refused(
                GanttRefreshRefusal.ReconciliationRefused,
                "Excel refused the first chart change, so the existing chart is unchanged.",
                validation.Issues);
    }

    /// <summary>
    /// Reads the <c>Duration</c> text each row currently shows, so the planner can
    /// exclude unchanged rows instead of rewriting the whole column.
    /// </summary>
    private static Dictionary<int, string?> ReadCurrentDurations(IReadOnlyList<GanttRowDto> rows)
    {
        Dictionary<int, string?> current = [];
        foreach (GanttRowDto row in rows)
        {
            current[row.RowNumber] = row.DurationCell.HasValue ? row.DurationCell.Value : null;
        }

        return current;
    }

    private static GanttRefreshOutcome Refuse(GanttRefreshRefusal refusal, string message) =>
        GanttRefreshOutcome.Refused(refusal, message);

    private static string Describe(GanttTableReadRefusalReason refusal) =>
        refusal switch
        {
            GanttTableReadRefusalReason.NoActiveWorkbook =>
                "There is no active workbook to refresh.",
            GanttTableReadRefusalReason.TableMissing =>
                "The Gantt table was not found on the active sheet.",
            GanttTableReadRefusalReason.DateSystemUnsupported =>
                "This workbook uses the 1904 date system, which Gantt Creator does not support.",
            _ => "The Gantt table could not be read.",
        };
}
