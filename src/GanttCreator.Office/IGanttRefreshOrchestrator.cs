using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office;

/// <summary>Why a refresh refused, as one flat, reportable reason.</summary>
/// <remarks>
/// <para>
/// One enum rather than a per-step type hierarchy. A Refresh has a single user-facing
/// outcome — the chart on the sheet is unchanged, and here is why — so the caller
/// needs one value to switch on and one message to show. Each member names the step
/// that refused, which is what makes the outcome actionable without the caller
/// re-deriving the pipeline.
/// </para>
/// </remarks>
public enum GanttRefreshRefusal
{
    /// <summary>The application object or active workbook was absent.</summary>
    NoActiveWorkbook = 0,

    /// <summary>The Gantt worksheet or its table was missing.</summary>
    TableMissing = 1,

    /// <summary>The configuration could not be read from the very-hidden sheet.</summary>
    ConfigurationUnreadable = 2,

    /// <summary>The live panel-grid measurement refused.</summary>
    MeasurementRefused = 3,

    /// <summary>The worksheet is protected, so the refresh cannot write.</summary>
    TargetProtected = 4,

    /// <summary>
    /// The rows carry a blocking validation error, so the refresh is not attempted.
    /// </summary>
    /// <remarks>
    /// The rows themselves are NOT a separate refusal reason: a blocking error is
    /// reported to the user through the validation report with its own row, field and
    /// code, which is far more actionable than "something in the table is wrong". This
    /// member says only that the refresh therefore did not run.
    /// </remarks>
    BlockingValidationErrors = 5,

    /// <summary>The scene-request factory refused; see its own typed reason.</summary>
    SceneRequestRefused = 6,

    /// <summary>The scene builder refused; see <see cref="SceneBuilderRefusal"/>.</summary>
    SceneBuildRefused = 7,

    /// <summary>The scene-to-shape translation refused or deferred a primitive.</summary>
    TranslationRefused = 8,

    /// <summary>The duration column could not be written.</summary>
    DurationWriteRefused = 9,

    /// <summary>The outline groups could not be maintained.</summary>
    OutlineRefused = 10,

    /// <summary>The row heights could not be normalised.</summary>
    RowHeightRefused = 11,

    /// <summary>The shape reconciliation refused part-way through.</summary>
    ReconciliationRefused = 12,

    /// <summary>
    /// The refresh completed some shape operations and then stopped.
    /// </summary>
    /// <remarks>
    /// Reported separately from <see cref="ReconciliationRefused"/> because the state
    /// it leaves behind is different: the chart on the sheet is now partly from the
    /// old scene and partly from the new one. The user needs to know their sheet is
    /// in that state; a plain refusal implies the previous chart is still intact and
    /// merely out of date. REV5 §16 defers transactional Refresh to R4.10, so this
    /// row reports the partial outcome rather than attempting a rollback.
    /// </remarks>
    PartialReconciliation = 13,

    /// <summary>The row identifiers could not be seeded or repaired.</summary>
    IdentityRefused = 14,

    /// <summary>
    /// The managed columns could not be restored to their hidden/locked
    /// classification (R4.7C D1/D2).
    /// </summary>
    /// <remarks>
    /// Reported as its own reason rather than folded into a generic "the sheet could
    /// not be prepared" because the user's remedy is specific: the workbook is in a
    /// state the add-in considers inconsistent, and the usual cause is that the sheet
    /// was protected or the table was edited by hand.
    /// </remarks>
    ColumnPresentationRefused = 15,

    /// <summary>
    /// The add-in's own cell notes could not be brought into line with the current
    /// validation findings.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The notes are the visible evidence of what the refresh found, so a refusal to
    /// maintain them is reported rather than swallowed. Reported before anything is
    /// rendered, which means the chart is untouched: the user keeps the last valid
    /// chart and sees a message explaining that the sheet's notes could not be
    /// updated.
    /// </para>
    /// <para>
    /// The usual cause is the same as <see cref="ColumnPresentationRefused"/> — a
    /// protected sheet or a hand-edited table — so the user's remedy is to unprotect
    /// the sheet and refresh again.
    /// </para>
    /// </remarks>
    ValidationNotesRefused = 16,
}

/// <summary>The typed result of one whole-sheet refresh.</summary>
/// <param name="Refusal">Why the refresh did not complete, or <see langword="null"/> on success.</param>
/// <param name="Message">A user-facing explanation, present whenever the refresh did not succeed.</param>
/// <param name="ValidationIssues">
/// The validation issues found, on both the success and the refusal path, so a caller
/// can report warnings even when the chart rendered.
/// </param>
/// <param name="ShapesWritten">How many shape operations the reconciliation performed.</param>
/// <param name="DurationCellsWritten">How many <c>Duration</c> cells were written.</param>
public sealed record GanttRefreshOutcome(
    GanttRefreshRefusal? Refusal,
    string? Message,
    IReadOnlyList<GanttValidationIssue> ValidationIssues,
    int ShapesWritten,
    int DurationCellsWritten)
{
    /// <summary>Gets whether the refresh completed.</summary>
    public bool Succeeded => Refusal is null;

    /// <summary>
    /// Gets whether the sheet was left in a partly-applied state that the user must
    /// know about.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This is <see cref="GanttRefreshRefusal.PartialReconciliation"/> <em>or</em> any
    /// refusal reached after a worksheet write landed. It previously named only
    /// <c>PartialReconciliation</c>, which understated the second case: a refusal
    /// after the Duration write left the user's Duration column changed even though no
    /// shape had been touched, and this property said the sheet was clean.
    /// </para>
    /// <para>
    /// The test is stated on the counts rather than on the refusal member, because the
    /// counts are the observable fact — they cannot be true while nothing was written.
    /// </para>
    /// </remarks>
    public bool LeftPartiallyApplied =>
        Refusal == GanttRefreshRefusal.PartialReconciliation
        || ShapesWritten > 0
        || DurationCellsWritten > 0;

    /// <summary>Creates a successful outcome.</summary>
    /// <param name="issues">The validation issues found, warnings included.</param>
    /// <param name="shapesWritten">How many shape operations ran.</param>
    /// <param name="durationCellsWritten">How many duration cells were written.</param>
    /// <returns>The successful outcome.</returns>
    public static GanttRefreshOutcome Ok(
        IReadOnlyList<GanttValidationIssue> issues,
        int shapesWritten,
        int durationCellsWritten) => new(null, null, issues, shapesWritten, durationCellsWritten);

    /// <summary>Creates a refusal.</summary>
    /// <param name="refusal">Why the refresh did not complete.</param>
    /// <param name="message">The explanation shown to the user.</param>
    /// <param name="issues">The validation issues found, which may be empty.</param>
    /// <param name="durationCellsWritten">
    /// How many <c>Duration</c> cells were written before the refusal. Zero is correct
    /// only for a refusal reached before any worksheet write.
    /// </param>
    /// <param name="shapesWritten">How many shape operations ran before the refusal.</param>
    /// <returns>The refusal outcome.</returns>
    /// <remarks>
    /// The counts used to be hardcoded to zero, which made the class's own promise —
    /// that a caller can tell "the chart was not touched" from "the sheet was not
    /// touched at all" — false for every refusal after the Duration write. A caller now
    /// passes what actually happened.
    /// </remarks>
    public static GanttRefreshOutcome Refused(
        GanttRefreshRefusal refusal,
        string message,
        IReadOnlyList<GanttValidationIssue>? issues = null,
        int durationCellsWritten = 0,
        int shapesWritten = 0) =>
        // NOTE: the record's positional order is (ShapesWritten, DurationCellsWritten),
        // which is the REVERSE of this method's parameter order. Named arguments are
        // used deliberately: passing them positionally swapped the two counts, which
        // is exactly the class of defect this method was widened to fix.
        new(refusal, message, issues ?? [], ShapesWritten: shapesWritten, DurationCellsWritten: durationCellsWritten);
}

/// <summary>
/// Assembles a whole Refresh from workbook state: read, validate, resolve, build the
/// scene, and reconcile the chart (R4.8A D1).
/// </summary>
/// <remarks>
/// <para>
/// <b>Why a service and not command logic.</b> D1: <c>RefreshSheetCommand</c> checks
/// availability, invokes this, and presents the outcome. The ordering below is a
/// contract with consequences — validation must precede mutation, the plot must be
/// resolved before the scene is built — and a rule that only holds when a Ribbon
/// callback happens to write it in the right order is not a rule.
/// </para>
/// <para>
/// <b>Every step is an injected port.</b> The orchestrator names no worksheet, no COM
/// proxy, and no interop type, so the whole pipeline is testable without Excel: a
/// test supplies fakes, injects a failure at each step, and asserts that no shape was
/// mutated. That is the property D3 exists to guarantee, and it is only observable
/// if the pipeline is reachable without a workbook.
/// </para>
/// <para>
/// <b>D3 — all validation, layout and scene construction completes before any shape
/// mutation.</b> A failure in any of them leaves the last valid chart untouched
/// rather than replacing it with a half-built one. The steps are therefore ordered so
/// that the only writes after the preflight are the reconciliation's own.
/// </para>
/// <para>
/// <b>COM and thread affinity.</b> A live implementation runs on the Excel main STA
/// thread and never marshals or offloads.
/// </para>
/// </remarks>
public interface IGanttRefreshOrchestrator
{
    /// <summary>Refreshes the Gantt chart from the current worksheet state.</summary>
    /// <returns>
    /// The typed outcome. On any refusal before the reconciliation the previous chart
    /// is untouched; see <see cref="GanttRefreshOutcome.LeftPartiallyApplied"/> for
    /// the one refusal that does not have that property.
    /// </returns>
    GanttRefreshOutcome Refresh();
}
