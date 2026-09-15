namespace GanttCreator.AddIn;

/// <summary>
/// Immutable snapshot of the host facts the dynamic Ribbon getters read. The
/// state service captures it only from an explicit refresh, so every getter is a
/// pure read of one value — deterministic, side-effect-free, and never touching
/// Excel.
/// </summary>
/// <param name="HasActiveWorkbook">True when Excel has an active workbook.</param>
/// <param name="LogAvailable">True when the rolling log exposes an active file path.</param>
internal sealed record RibbonState(bool HasActiveWorkbook, bool LogAvailable)
{
    /// <summary>
    /// The snapshot before any successful capture. No fact is known, so every
    /// gated control starts disabled and is enabled by the first refresh
    /// (docs/03-ROADMAP.md R1.5).
    /// </summary>
    internal static RibbonState Initial { get; } = new(false, false);

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
        _ => true,
    };
}
