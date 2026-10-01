namespace GanttCreator.Core;

/// <summary>
/// The single rule deciding whether a table row is a semantically blank entry
/// row rather than an entity. Immutable; no state.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why this exists.</b> Two private, inconsistent blank-row predicates had
/// already grown in the Office layer: the table reader's treated
/// <see langword="null"/> and whitespace as blank, while the row inserter's
/// additionally treated <c>Missing</c> and <c>DBNull</c> as blank. A row one
/// considered reusable and the other considered populated is exactly the defect
/// R4.7A D3 names. There is now one rule and both layers reach it.
/// </para>
/// <para>
/// <see cref="Test"/> and <see cref="TestStrict"/> check exactly four cells: the
/// <c>Type</c> and <c>Description</c> texts and the <c>Start</c> and <c>Finish</c>
/// dates. Every other cell — <c>Id</c>, <c>ParentId</c>, <c>StyleKey</c>, the
/// colour overrides, <c>LabelPosition</c>, <c>SortOrder</c>, and the engine's own
/// <c>LaneId</c>/<c>StackIndex</c>/<c>SiblingOrder</c>/<c>Visible</c> — is
/// deliberately outside the rule, so a row carrying only such a value is
/// <b>blank</b>.
/// </para>
/// <para>
/// <b>Why those four, and why engine-maintained cells are excluded.</b> Type,
/// Description, Start and Finish are the user's authoring columns, so those are
/// the ones whose absence means the user has not entered anything. Engine columns
/// are the opposite: they are written <em>because</em> a row is an entity, or
/// scaffolded in advance of one. Treating a non-empty <c>LaneId</c> or
/// <c>StackIndex</c> as meaningful would mean a freshly scaffolded blank row is
/// never blank; blank-row detection would then fail on every new row and the
/// "exactly one blank entry row" invariant would be unsatisfiable. The Id,
/// ParentId, StyleKey, colour, label-position and SortOrder cells sit in the same
/// group for the same reason — the add-in populates them for a row it has already
/// decided is an entity, so their presence is not independent evidence that the
/// user did anything.
/// </para>
/// </remarks>
public static class IsSemanticallyEmptyRow
{
    /// <summary>
    /// Whether the row holds no value in any of the four authoring cells —
    /// <c>Type</c>, <c>Description</c>, <c>Start</c> or <c>Finish</c>. Every other
    /// cell is outside the rule; see the type remarks. Judge on the normalized
    /// views, where a whitespace-only string is already absent.
    /// </summary>
    /// <param name="row">The row to test. A <see langword="null"/> row is blank.</param>
    /// <returns><see langword="true"/> when all four authoring cells are absent.</returns>
    public static bool Test(GanttRowDto? row) =>
        row is null
        || (IsBlank(row.TypeText)
            && IsBlank(row.Description)
            && row.Start is null
            && row.Finish is null);

    /// <summary>
    /// The stricter form, judged on raw cell state so an Excel error value in one
    /// of the four authoring cells is not mistaken for absence. This is the one the
    /// validator uses: a cell holding <c>#N/A</c> is not blank, because the user
    /// put something there and must be told about it rather than have the row
    /// silently skipped.
    /// </summary>
    /// <param name="row">The row to test. A <see langword="null"/> row is blank.</param>
    /// <returns><see langword="true"/> when all four authoring cells are genuinely empty.</returns>
    public static bool TestStrict(GanttRowDto? row) =>
        row is null
        || (IsBlankText(row.TypeCell)
            && IsBlankText(row.DescriptionCell)
            && IsAbsent(row.StartCell)
            && IsAbsent(row.FinishCell));

    /// <summary>
    /// Whether the row carries any user-authored content, the exact negation of
    /// <see cref="Test"/>. Provided so call sites read as the question they ask.
    /// </summary>
    /// <param name="row">The row to test.</param>
    /// <returns><see langword="true"/> when the row is an entity.</returns>
    public static bool IsEntity(GanttRowDto? row) => !Test(row);

    private static bool IsAbsent<T>(GanttCell<T> cell) => cell.State == GanttCellState.Empty;

    /// <summary>
    /// A string value counts as absent when it is null or entirely whitespace.
    /// Excel writes an empty string to a cell a user has cleared, and a stray
    /// space is not user content; treating either as meaningful would create an
    /// entity with an empty description and consume an identity for it.
    /// </summary>
    private static bool IsBlank(string? value) => string.IsNullOrWhiteSpace(value);

    /// <summary>
    /// The strict form of the same test on a raw text cell: a whitespace-only
    /// value is a <see cref="GanttCellState.Value"/> cell, so the state alone
    /// would report it as content.
    /// </summary>
    private static bool IsBlankText(GanttCell<string> cell) =>
        cell.State == GanttCellState.Empty
        || (cell.State == GanttCellState.Value && IsBlank(cell.Value));
}
