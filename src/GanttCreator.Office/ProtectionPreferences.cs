namespace GanttCreator.Office;

/// <summary>
/// Reads the user's two protection toggles: whether the chart's shapes are locked, and whether the
/// sheet's cells outside the visible table are unlocked.
/// </summary>
/// <remarks>
/// A read-only view, so an adapter can consult the toggles without being able to change them. The
/// mutable side is <see cref="ProtectionPreferences"/>.
/// </remarks>
public interface IProtectionPreferenceSource
{
    /// <summary>
    /// Gets a value indicating whether the chart's shapes are locked.
    /// </summary>
    /// <remarks>
    /// <see langword="true"/> means <c>Protect(DrawingObjects:=true)</c> and every owned shape carries
    /// <c>Locked = true</c>. <see langword="false"/> means the user asked to be able to drag chart
    /// elements.
    /// </remarks>
    bool ChartLocked { get; }

    /// <summary>
    /// Gets a value indicating whether the sheet's cells OUTSIDE the visible table are editable.
    /// </summary>
    /// <remarks>
    /// <see langword="false"/> (the default) means <c>Protect(Contents:=true)</c> and every cell
    /// outside the table stays read-only. <see langword="true"/> turns contents protection off, which
    /// is what makes the rest of the sheet clickable and editable. The visible table is never locked
    /// either way — that is a separate, permanent rule.
    /// </remarks>
    bool CellsUnlocked { get; }
}

/// <summary>
/// The Excel session's two protection toggles, with the default state of each.
/// </summary>
/// <remarks>
/// <para>
/// <b>Owner ruling 2026-11-09.</b> The ribbon carries one toggle for the CHART and one for the
/// SHEET CELLS, both visibly reflecting the state they are in. The visible table is a third thing and
/// has no toggle: it is always editable, so the user can copy, paste and format schedule data at any
/// time. Cells outside the table are locked unless the cell toggle is switched on.
/// </para>
/// <para>
/// <b>Why this is separate from <see cref="ShapeUnlockState"/>.</b> That type recorded which single
/// shapes the user had unlocked out of the reconciler's reach — a per-shape, transient exemption. This
/// type is the durable, sheet-wide pair of switches the ribbon presents: it outlives a refresh, and a
/// refresh re-applies protection FROM it rather than overwriting it.
/// </para>
/// <para>
/// Mutable and explicitly shared, like <see cref="ProtectionArmingState"/> and
/// <see cref="ShapeUnlockState"/>: it lives as long as the Excel session, is touched only on the main
/// thread, and is reset when the active workbook changes because it describes one sheet's protection.
/// </para>
/// </remarks>
public sealed class ProtectionPreferences : IProtectionPreferenceSource
{
    /// <summary>The Excel session's single pair of toggles.</summary>
    public static ProtectionPreferences Shared { get; } = new();

    /// <summary>
    /// Gets or sets a value indicating whether the chart's shapes are locked. Defaults to
    /// <see langword="true"/>, which is the state every workbook has shipped in since ADR-0055.
    /// </summary>
    public bool ChartLocked { get; set; } = true;

    /// <summary>
    /// Gets or sets a value indicating whether cells outside the visible table are editable. Defaults
    /// to <see langword="false"/>: the rest of the sheet is read-only, so a stray click cannot damage
    /// the layout around the chart.
    /// </summary>
    public bool CellsUnlocked { get; set; }

    /// <summary>
    /// Gets a value indicating whether the user has released anything, so the automatic refresh must
    /// keep its hands off.
    /// </summary>
    /// <remarks>
    /// The scheduler reads this before queueing an automatic refresh. A manual shape drag is in
    /// flight whenever the chart is unlocked, and rebuilding the chart underneath it would discard the
    /// edit the user is making — the same reason <see cref="ShapeUnlockState.IsAnyUnlocked"/> exists.
    /// </remarks>
    public bool IsAnyUnlocked => !ChartLocked || CellsUnlocked;

    /// <summary>Restores both toggles to their defaults.</summary>
    /// <remarks>
    /// Called when the active workbook changes and when the user opts out of protection, because the
    /// toggles describe one sheet's state rather than the session's.
    /// </remarks>
    public void Reset()
    {
        ChartLocked = true;
        CellsUnlocked = false;
    }
}
