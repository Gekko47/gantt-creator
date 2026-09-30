using System.Globalization;

namespace GanttCreator.Core;

/// <summary>
/// The destructive-command policy ratified by ADR-0008: no undo, warned
/// confirmation, protected-sheet guard first. This file owns the shared Core
/// contract every future mutating command must consult — classification,
/// confirmation text, and the evidence ledger. It is pure: no Office,
/// no interop, no UI.
/// </summary>
/// <remarks>
/// <para>
/// ADR-0008 D1 defines "destructive command" as any Office-mutating adapter
/// path that clears, overwrites, or recreates user-visible cell content or
/// table geometry in <c>tblGanttData</c> or the configuration worksheets.
/// Normal worksheet edits, Type/colour/label-position/selection changes, and
/// Refresh/render are not destructive commands under this ADR and do not need
/// this policy.
/// </para>
/// <para>
/// ADR-0008 D2 (no undo): the add-in does not expose an Excel undo-stack entry
/// or a transactional rollback for destructive commands. The confirmation text
/// states this explicitly and the exact phrasing is pinned by test.
/// </para>
/// <para>
/// ADR-0008 D4 (guard ordering): the workbook-protection guard is the first
/// check in every mutating adapter; the confirmation is the user-facing guard
/// that follows when the operation is user-initiated. This class does not own
/// the protection guard — that lives in <c>GanttCreator.Office</c> — but it
/// classifies commands and builds the confirmation text that the guard's caller
/// presents.
/// </para>
/// </remarks>
public static class DestructiveCommandPolicy
{
    /// <summary>
    /// The dialog caption presented for every destructive-command confirmation.
    /// Constant across all command classes, so it is declared once here rather
    /// than repeated per command.
    /// </summary>
    public const string ConfirmationCaption = "Gantt Creator";

    /// <summary>
    /// The dialog main instruction presented for every destructive-command
    /// confirmation. Constant across all command classes.
    /// </summary>
    public const string ConfirmationMainInstruction =
        "Gantt Creator is about to make a destructive change";

    /// <summary>
    /// The recognised destructive-command classes. Each future mutating command
    /// picks exactly one; the classification drives whether a confirmation is
    /// required and what the "cannot be undone" line must say.
    /// </summary>
    public enum CommandClass
    {
        /// <summary>
        /// Clears the visible <c>tblGanttData</c> body (row data) without
        /// recreating the table geometry. Under ADR-0008 this is destructive
        /// and must be confirmed.
        /// </summary>
        ClearTableBody = 0,

        /// <summary>
        /// Recolumnises <c>tblGanttData</c>: drops and recreates the table
        /// with the current schema columns. Under ADR-0008 this is destructive
        /// and must be confirmed.
        /// </summary>
        RecolumniseTable = 1,

        /// <summary>
        /// Resets the <c>_GanttCreatorConfig</c> catalogues back to the
        /// code-owned defaults, preserving user-authored style rows where the
        /// catalogue contract allows it. Under ADR-0008 this is destructive
        /// and must be confirmed.
        /// </summary>
        ResetCatalogues = 2,
    }

    /// <summary>
    /// Whether a destructive command presents a confirmation dialog and what
    /// kind. Every destructive command in the supported set is destructive-but-
    /// confirmed (D3): the user is warned and the dialog states the change
    /// cannot be undone. The enum is the stable shape of that decision so
    /// future commands cannot silently change it.
    /// </summary>
    public enum ConfirmationKind
    {
        /// <summary>
        /// The command is destructive and presents a warned confirmation that
        /// states the change cannot be undone (ADR-0008 D3/D2). This is the
        /// classification for every supported destructive command today.
        /// </summary>
        DestructiveButConfirmed = 0,
    }

    /// <summary>
    /// Classifies a destructive command for policy purposes: whether it needs a
    /// confirmation dialog and what kind. Pure lookup — no side effects.
    /// </summary>
    /// <param name="commandClass">
    /// The destructive-command class to classify.
    /// </param>
    /// <returns>
    /// The confirmation kind the command must present. For every supported
    /// command today this is <see cref="ConfirmationKind.DestructiveButConfirmed"/>.
    /// </returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when <paramref name="commandClass"/> is not one of the three
    /// recognised values. The default <c>0</c> value is a valid, named class
    /// (<see cref="CommandClass.ClearTableBody"/>) and does not trigger this
    /// exception.
    /// </exception>
    /// <remarks>
    /// <para>
    /// The guard is intentional: a destructive command that reaches the policy
    /// layer without a class is a missing-classification defect, not a routine
    /// "no confirmation" case. The guard makes that defect loud at the call
    /// site instead of silently skipping the confirmation.
    /// </para>
    /// <para>
    /// It is an <see cref="ArgumentOutOfRangeException"/> rather than an
    /// <see cref="ArgumentNullException"/> because <paramref name="commandClass"/>
    /// is a non-nullable value type: a caller cannot pass "no class" at all,
    /// only a value outside the supported set.
    /// </para>
    /// </remarks>
    public static ConfirmationKind Classify(CommandClass commandClass) =>
        commandClass switch
        {
            CommandClass.ClearTableBody or
            CommandClass.RecolumniseTable or
            CommandClass.ResetCatalogues => ConfirmationKind.DestructiveButConfirmed,

            _ => throw new ArgumentOutOfRangeException(
                nameof(commandClass),
                commandClass,
                "Unrecognised destructive-command class; every destructive command must be classified."),
        };

    /// <summary>
    /// Builds the user-facing confirmation text for a destructive command. Pure
    /// and deterministic: the exact body text is pinned by test so the
    /// "cannot be undone" phrasing (ADR-0008 D3) cannot drift between the Core
    /// builder and any dialog that presents it.
    /// </summary>
    /// <param name="commandClass">
    /// The destructive-command class being confirmed.
    /// </param>
    /// <returns>
    /// The full confirmation body, including the no-undo line. Ready to hand to
    /// the confirmation dialog.
    /// </returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when <paramref name="commandClass"/> is unrecognised, per
    /// <see cref="Classify(CommandClass)"/>.
    /// </exception>
    public static ConfirmationText BuildConfirmationText(CommandClass commandClass)
    {
        // Validate the classification first so the text is always built for a
        // recognised command.
        _ = Classify(commandClass);

        var commandName = CommandDisplayName(commandClass);
        var subject = CommandSubject(commandClass);

        // ADR-0008 D3/D2: the body must explicitly state the change cannot be
        // undone, using exactly that phrasing. The exact line is pinned in test
        // so it cannot drift.
        const string noUndoLine = "This change cannot be undone.";

        // Count-only body: the text names the command and its subject but never
        // enumerates row counts, table sizes, or catalogue row counts from the
        // workbook. That keeps the builder Office-free and deterministic.
        var body = string.Create(
            CultureInfo.InvariantCulture,
            $"{commandName} {subject}. {noUndoLine}");

        return new ConfirmationText(ConfirmationCaption, ConfirmationMainInstruction, body);
    }

    /// <summary>
    /// Returns the short display name of a destructive command, for use in the
    /// confirmation body. Pure, culture-invariant.
    /// </summary>
    /// <param name="commandClass">The destructive-command class.</param>
    /// <returns>
    /// The display name, e.g. "Clear table body".
    /// </returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when <paramref name="commandClass"/> is not one of the three
    /// recognised values, per <see cref="Classify(CommandClass)"/>.
    /// </exception>
    public static string CommandDisplayName(CommandClass commandClass) =>
        commandClass switch
        {
            CommandClass.ClearTableBody => "Clear table body",
            CommandClass.RecolumniseTable => "Recolumnise table",
            CommandClass.ResetCatalogues => "Reset configuration catalogues",
            _ => throw new ArgumentOutOfRangeException(
                nameof(commandClass),
                commandClass,
                "Unrecognised destructive-command class."),
        };

    /// <summary>
    /// Returns the short subject phrase for a destructive command, for use in
    /// the confirmation body. Pure, culture-invariant.
    /// </summary>
    /// <param name="commandClass">The destructive-command class.</param>
    /// <returns>
    /// The subject phrase, e.g. "will delete all rows from the Gantt data table."
    /// </returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when <paramref name="commandClass"/> is not one of the three
    /// recognised values, per <see cref="Classify(CommandClass)"/>.
    /// </exception>
    public static string CommandSubject(CommandClass commandClass) =>
        commandClass switch
        {
            CommandClass.ClearTableBody =>
                "will delete all rows from the Gantt data table.",
            CommandClass.RecolumniseTable =>
                "will recreate the Gantt data table with the current columns.",
            CommandClass.ResetCatalogues =>
                "will restore the configuration catalogues to their defaults.",
            _ => throw new ArgumentOutOfRangeException(
                nameof(commandClass),
                commandClass,
                "Unrecognised destructive-command class."),
        };
}

/// <summary>
/// The three fields a destructive-command confirmation dialog renders, built
/// by <see cref="DestructiveCommandPolicy.BuildConfirmationText(DestructiveCommandPolicy.CommandClass)"/>.
/// </summary>
/// <remarks>
/// A record rather than a single formatted string: the dialog needs three
/// separate fields, and encoding them into one newline-delimited string would
/// make the boundary ambiguous the moment any field could contain a newline.
/// <see cref="Caption"/> and <see cref="MainInstruction"/> are the same for
/// every command; only <see cref="Body"/> varies.
/// </remarks>
/// <param name="Caption">The dialog caption.</param>
/// <param name="MainInstruction">The dialog's main instruction line.</param>
/// <param name="Body">
/// The dialog body, naming the command and its subject and stating the exact
/// no-undo line (ADR-0008 D3/D2).
/// </param>
public sealed record ConfirmationText(
    string Caption,
    string MainInstruction,
    string Body);
