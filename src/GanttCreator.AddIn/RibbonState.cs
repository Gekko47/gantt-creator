namespace GanttCreator.AddIn;

/// <summary>
/// Immutable snapshot of the host facts the dynamic Ribbon getters read. The
/// state service captures it only from an explicit refresh, so every getter is a
/// pure read of one value — deterministic, side-effect-free, and never touching
/// Excel.
/// </summary>
/// <param name="HasActiveWorkbook">True when Excel has an active workbook.</param>
/// <param name="LogAvailable">True when the rolling log exposes an active file path.</param>
/// <param name="SheetInitialised">True when the active workbook has a Gantt sheet that has been initialised (tblGanttData exists).</param>
/// <param name="PlotStartAuto">True when the plot start derives from the data range (R5.1).</param>
/// <param name="PlotFinishAuto">True when the plot finish derives from the data range (R5.1).</param>
/// <param name="PlotStartDate">The explicit plot start date, or empty when automatic (R5.1).</param>
/// <param name="PlotFinishDate">The explicit plot finish date, or empty when automatic (R5.1).</param>
/// <param name="EffectivePlotStartDate">
/// The date the start edit box displays: the explicit date in explicit mode,
/// the derived data range start in automatic mode, or empty when neither is
/// known (fix plan ruling 1 — manual-entry boxes also display the current
/// effective plot date, even when disabled).
/// </param>
/// <param name="EffectivePlotFinishDate">
/// The date the finish edit box displays: the explicit date in explicit mode,
/// the derived data range finish in automatic mode, or empty when neither is
/// known (fix plan ruling 1).
/// </param>
internal sealed record RibbonState(
    bool HasActiveWorkbook,
    bool LogAvailable,
    bool SheetInitialised,
    bool PlotStartAuto,
    bool PlotFinishAuto,
    string PlotStartDate,
    string PlotFinishDate,
    string EffectivePlotStartDate,
    string EffectivePlotFinishDate)
{
    /// <summary>
    /// The snapshot before any successful capture. No fact is known, so every
    /// gated control starts disabled and is enabled by the first refresh
    /// (docs/03-ROADMAP.md R1.5).
    /// </summary>
    internal static RibbonState Initial { get; } = new(false, false, false, true, true, string.Empty, string.Empty, string.Empty, string.Empty);

    /// <summary>
    /// Gets whether the control with <paramref name="controlId"/> is enabled
    /// under this snapshot: the pure truth table behind the Ribbon's
    /// <c>getEnabled</c> callback.
    /// </summary>
    /// <param name="controlId">The Ribbon control ID, or null when it could not be probed.</param>
    /// <returns>
    /// The control's enabled state. Unknown, null, or whitespace IDs are
    /// <see langword="true"/> (fail-open): a permanently grey button is a worse
    /// failure than a briefly enabled one, and the command boundary still
    /// validates at execution time.
    /// </returns>
    internal bool IsEnabled(string? controlId) => controlId switch
    {
        RibbonControlIds.Diagnostics => HasActiveWorkbook,
        RibbonControlIds.OpenLog => LogAvailable,
        RibbonControlIds.InitialiseSheet => HasActiveWorkbook,
        RibbonControlIds.ValidateSheet => HasActiveWorkbook,
        RibbonControlIds.RefreshSheet => HasActiveWorkbook && SheetInitialised,
        RibbonControlIds.RepairConfig => HasActiveWorkbook,
        RibbonControlIds.AddActivity => HasActiveWorkbook && SheetInitialised,
        RibbonControlIds.AddMilestone => HasActiveWorkbook && SheetInitialised,
        RibbonControlIds.AddDelineator => HasActiveWorkbook && SheetInitialised,
        RibbonControlIds.PlotStartAuto => HasActiveWorkbook && SheetInitialised,
        RibbonControlIds.PlotFinishAuto => HasActiveWorkbook && SheetInitialised,
        RibbonControlIds.PlotStartDate => HasActiveWorkbook && SheetInitialised && !PlotStartAuto,
        RibbonControlIds.PlotFinishDate => HasActiveWorkbook && SheetInitialised && !PlotFinishAuto,
        _ => true,
    };
}
