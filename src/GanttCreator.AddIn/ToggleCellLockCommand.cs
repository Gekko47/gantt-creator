using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// Flips the CELL toggle: makes the sheet's cells outside the visible table editable, or locks them
/// again.
/// </summary>
/// <remarks>
/// <para>
/// <b>What it does and does not govern (owner ruling 2026-11-09).</b> The visible table is always
/// editable — copy, paste and formatting are never blocked there — so this toggle is about the rest
/// of the sheet: the frame around the chart, the area below the table, and every other cell. Locked is
/// the default, so a stray click cannot damage the layout.
/// </para>
/// <para>
/// <b>One flag.</b> The toggle maps onto <c>Protect(Contents:=...)</c> and nothing else, which is what
/// keeps it independent of the chart toggle: releasing the cells must not release the shapes, and
/// vice versa. <see cref="ExcelWorksheetProtectionPort"/> reads both flags from
/// <see cref="ProtectionPreferences"/> on every protect, so a refresh re-applies the user's choice.
/// </para>
/// </remarks>
internal static class ToggleCellLockCommand
{
    /// <summary>Runs the production toggle in the current Excel session.</summary>
    /// <param name="locked">
    /// The state the toggle has moved to: <see langword="true"/> to lock the cells outside the table.
    /// RibbonX passes this to a <c>toggleButton</c>'s <c>onAction</c>.
    /// </param>
    internal static void RunForExcel(bool locked)
    {
        object? application = ExcelDna.Integration.ExcelDnaUtil.Application;

        Run(
            new ExcelWorksheetProtectionPort(application),
            ProtectionPreferences.Shared,
            locked,
            CommandErrorDialog.Show);

        RibbonStateService.Instance.InvalidateControl(RibbonControlIds.ToggleChartLock);
        RibbonStateService.Instance.InvalidateControl(RibbonControlIds.ToggleCellLock);
    }

    /// <summary>Runs an injected toggle, for tests and for the R4.9 wiring.</summary>
    /// <param name="protection">Re-applies the sheet protection with the new flags.</param>
    /// <param name="preferences">The session's toggles; the cell one is set from <paramref name="locked"/>.</param>
    /// <param name="locked">The requested state of the cells outside the visible table.</param>
    /// <param name="presenter">The user-safe refusal presenter.</param>
    internal static void Run(
        IWorksheetProtectionPort protection,
        ProtectionPreferences preferences,
        bool locked,
        Action<string> presenter)
    {
        ArgumentNullException.ThrowIfNull(protection);
        ArgumentNullException.ThrowIfNull(preferences);
        ArgumentNullException.ThrowIfNull(presenter);

        // Set first, so the protection below is applied from the state the button shows even if the
        // host refuses the write. Assigning the requested state keeps the command idempotent.
        bool unlockNow = !locked;
        preferences.CellsUnlocked = unlockNow;

        ProtectionWriteOutcome outcome = protection.Protect();
        if (outcome.Refusal is { } refusal)
        {
            presenter(
                unlockNow
                    ? "The cells outside the table could not be released: " + refusal + "."
                    : "The cells outside the table could not be locked again: " + refusal + ".");
        }
    }
}
