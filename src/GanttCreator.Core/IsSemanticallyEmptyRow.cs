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
/// <b>What "semantically empty" means.</b> Every cell the <em>user</em> authors
/// must be absent. A row carrying a Type, a description, dates, an Id, a
/// ParentId, a StyleKey, colours, a label position or a SortOrder is
/// <b>meaningful</b>, whatever the rest of the row looks like.
/// </para>
/// <para>
/// <b>Why engine-maintained cells are excluded.</b> Treating a non-empty
/// <c>LaneId</c> or <c>StackIndex</c> as meaningful would mean a freshly
/// scaffolded blank row is never blank. Blank-row detection would then fail on
/// every new row and the "exactly one blank entry row" invariant would be
/// unsatisfiable. Engine columns are written <em>because</em> a row is an
/// entity, not before it is one.
/// </para>
/// </remarks>
public static class IsSemanticallyEmptyRow
{
    /// <summary>
    /// Whether the row holds no user-authored value. Judge on the normalized
    /// views, where a whitespace-only string is already absent.
    /// </summary>
    /// <param name="row">The row to test. A <see langword="null"/> row is blank.</param>
    /// <returns><see langword="true"/> when no user-authored value is present.</returns>
    public static bool Test(GanttRowDto? row) =>
        row is null
        || (IsBlank(row.TypeText)
            && IsBlank(row.Description)
            && row.Start is null
            && row.Finish is null);

    /// <summary>
    /// The stricter form, judged on raw cell state so an Excel error value in a
    /// user-authored cell is not mistaken for absence. This is the one the
    /// validator uses: a cell holding <c>#N/A</c> is not blank, because the user
    /// put something there and must be told about it rather than have the row
    /// silently skipped.
    /// </summary>
    /// <param name="row">The row to test. A <see langword="null"/> row is blank.</param>
    /// <returns><see langword="true"/> when every user-authored cell is genuinely empty.</returns>
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
