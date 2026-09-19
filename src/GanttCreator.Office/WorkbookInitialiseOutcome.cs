namespace GanttCreator.Office;

/// <summary>
/// Why an initialise attempt refused to change the workbook.
/// </summary>
/// <remarks>
/// Every refusal is a typed, expected outcome — routine, not an exception
/// (docs/02-ARCHITECTURE.md "Error handling": expected user/data errors return
/// typed results; do not throw for routine bad user input). Unexpected COM
/// failures are not refusals and propagate to the command boundary.
/// </remarks>
public enum InitialiseRefusalReason
{
    /// <summary>The Excel application object or the active workbook was absent.</summary>
    NoActiveWorkbook = 0,

    /// <summary>The selected target worksheet already contains a table named <c>tblGanttData</c>.</summary>
    TableExists = 1,

    /// <summary>
    /// The workbook already contains the <c>_GanttCreatorConfig</c>
    /// configuration worksheet; repair belongs to the R2.10 workflow.
    /// </summary>
    ConfigSheetExists = 2,

    /// <summary>The target worksheet is protected; renaming or populating it is not possible.</summary>
    TargetProtected = 3,
}

/// <summary>
/// The typed outcome of one initialise attempt: which path was taken
/// (adopted, created, refused) and the final label of the Gantt worksheet
/// when the attempt succeeded.
/// </summary>
/// <param name="Path">The path the attempt took.</param>
/// <param name="SheetName">
/// The final label of the Gantt worksheet on success; <see langword="null"/>
/// on a refusal.
/// </param>
/// <param name="Refusal">
/// The refusal reason when <paramref name="Path"/> is
/// <see cref="WorkbookInitialisePath.Refused"/>; otherwise
/// <see langword="null"/>.
/// </param>
public sealed record WorkbookInitialiseOutcome(
    WorkbookInitialisePath Path,
    string? SheetName,
    InitialiseRefusalReason? Refusal)
{
    /// <summary>
    /// Initialises an outcome for the adopt path: the blank active worksheet
    /// was renamed and populated in place.
    /// </summary>
    /// <param name="sheetName">The final label of the adopted worksheet.</param>
    /// <returns>The success outcome.</returns>
    public static WorkbookInitialiseOutcome Adopted(string sheetName) => new(
        WorkbookInitialisePath.Adopted, sheetName, null);

    /// <summary>
    /// Initialises an outcome for the create path: a new worksheet was created,
    /// labelled, and populated; existing sheets were untouched.
    /// </summary>
    /// <param name="sheetName">The final label of the created worksheet.</param>
    /// <returns>The success outcome.</returns>
    public static WorkbookInitialiseOutcome CreatedNew(string sheetName) => new(
        WorkbookInitialisePath.CreatedNew, sheetName, null);

    /// <summary>
    /// Initialises an outcome for the refusal path. The workbook is unchanged;
    /// the caller may surface the reason to the user.
    /// </summary>
    /// <param name="refusal">Why nothing was changed.</param>
    /// <returns>The refusal outcome.</returns>
    public static WorkbookInitialiseOutcome Refused(InitialiseRefusalReason refusal) => new(
        WorkbookInitialisePath.Refused, null, refusal);

    /// <summary>Gets a value indicating whether the workbook was initialised.</summary>
    public bool Succeeded => Path != WorkbookInitialisePath.Refused;
}

/// <summary>Which structural path one initialise attempt took.</summary>
public enum WorkbookInitialisePath
{
    /// <summary>The blank active worksheet was adopted, renamed, and populated.</summary>
    Adopted = 0,

    /// <summary>A new worksheet was created, labelled, and populated.</summary>
    CreatedNew = 1,

    /// <summary>Nothing was changed; the workbook is exactly as it was.</summary>
    Refused = 2,
}
