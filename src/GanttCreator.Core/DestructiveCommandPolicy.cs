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
    /// <exception cref="ArgumentNullException">
    /// Thrown when <paramref name="commandClass"/> is the default
    /// <c>0</c> value and no command class was specified — callers that
    /// construct an unclassified command have made a programmer error.
    /// </exception>
    /// <remarks>
    /// <para>
    /// The <see cref="ArgumentNullException"/> guard is intentional: a
    /// destructive command that reaches the policy layer without a class is a
    /// missing-classification defect, not a routine "no confirmation" case. The
    /// guard makes that defect loud at the call site instead of silently
    /// skipping the confirmation.
    /// </para>
    /// </remarks>
    public static ConfirmationKind Classify(CommandClass commandClass) =>
        commandClass switch
        {
            CommandClass.ClearTableBody or
            CommandClass.RecolumniseTable or
            CommandClass.ResetCatalogues => ConfirmationKind.DestructiveButConfirmed,

            _ => throw new ArgumentNullException(
                nameof(commandClass),
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
    /// <exception cref="ArgumentNullException">
    /// Thrown when <paramref name="commandClass"/> is unrecognised, per
    /// <see cref="Classify(CommandClass)"/>.
    /// </exception>
    public static string BuildConfirmationText(CommandClass commandClass)
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
        const string caption = "Gantt Creator";
        const string mainInstruction =
            "Gantt Creator is about to make a destructive change";

        // Count-only body: the text names the command and its subject but never
        // enumerates row counts, table sizes, or catalogue row counts from the
        // workbook. That keeps the builder Office-free and deterministic.
        var body = string.Create(
            CultureInfo.InvariantCulture,
            $"{commandName} {subject}. {noUndoLine}");

        return string.Create(
            CultureInfo.InvariantCulture,
            $"caption={caption}\nmainInstruction={mainInstruction}\nbody={body}");
    }

    /// <summary>
    /// Returns the short display name of a destructive command, for use in the
    /// confirmation body. Pure, culture-invariant.
    /// </summary>
    /// <param name="commandClass">The destructive-command class.</param>
    /// <returns>
    /// The display name, e.g. "Clear table body".
    /// </returns>
    public static string CommandDisplayName(CommandClass commandClass) =>
        commandClass switch
        {
            CommandClass.ClearTableBody => "Clear table body",
            CommandClass.RecolumniseTable => "Recolumnise table",
            CommandClass.ResetCatalogues => "Reset configuration catalogues",
            _ => throw new ArgumentNullException(
                nameof(commandClass),
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
    public static string CommandSubject(CommandClass commandClass) =>
        commandClass switch
        {
            CommandClass.ClearTableBody =>
                "will delete all rows from the Gantt data table.",
            CommandClass.RecolumniseTable =>
                "will recreate the Gantt data table with the current columns.",
            CommandClass.ResetCatalogues =>
                "will restore the configuration catalogues to their defaults.",
            _ => throw new ArgumentNullException(
                nameof(commandClass),
                "Unrecognised destructive-command class."),
        };
}
