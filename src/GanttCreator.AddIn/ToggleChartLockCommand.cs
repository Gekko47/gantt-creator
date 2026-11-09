using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// Flips the CHART toggle: locks every owned shape and enforces <c>DrawingObjects</c>, or releases
/// both so the user can drag chart elements.
/// </summary>
/// <remarks>
/// <para>
/// <b>One button, two states (owner ruling 2026-11-09).</b> This replaces the separate "Unlock shape"
/// and "Lock shapes" pair. The old pair acted on one selected shape and could never enable itself;
/// the ruling is a single control that reflects the chart's state and changes it, so the enablement
/// question disappears: the button is always available on an initialised sheet, and what the user
/// sees is the state they are in.
/// </para>
/// <para>
/// <b>The toggle is session state, not a one-shot.</b> <see cref="ProtectionPreferences.ChartLocked"/>
/// is what <see cref="ExcelWorksheetProtectionPort"/> reads when it applies protection, and what
/// <see cref="WorksheetProtectionArming"/> reads before it re-locks anything. A refresh therefore
/// re-applies the user's choice instead of overwriting it.
/// </para>
/// <para>
/// <b>Suspension.</b> While the chart is unlocked <see cref="ProtectionPreferences.IsAnyUnlocked"/> is
/// true, so the automatic refresh stays out of the way of a manual drag. An explicit Refresh still
/// reconciles, which is how a discarded manual edit comes back.
/// </para>
/// </remarks>
internal static class ToggleChartLockCommand
{
    /// <summary>Runs the production toggle in the current Excel session.</summary>
    internal static void RunForExcel()
    {
        object? application = ExcelDna.Integration.ExcelDnaUtil.Application;

        Run(
            new ExcelShapeWriter(application, new ExcelWorksheetProtectionGuard(application)),
            new ExcelWorksheetProtectionPort(application),
            ProtectionPreferences.Shared,
            CommandErrorDialog.Show);

        RibbonStateService.Instance.InvalidateControl(RibbonControlIds.ToggleChartLock);
        RibbonStateService.Instance.InvalidateControl(RibbonControlIds.ToggleCellLock);
    }

    /// <summary>Runs an injected toggle, for tests and for the R4.9 wiring.</summary>
    /// <param name="shapes">Applies or clears <c>Shape.Locked</c> on the owned shapes.</param>
    /// <param name="protection">Re-applies the sheet protection with the new flags.</param>
    /// <param name="preferences">The session's toggles; the chart one is flipped.</param>
    /// <param name="presenter">The user-safe refusal presenter.</param>
    internal static void Run(
        IShapeWritePort shapes,
        IWorksheetProtectionPort protection,
        ProtectionPreferences preferences,
        Action<string> presenter)
    {
        ArgumentNullException.ThrowIfNull(shapes);
        ArgumentNullException.ThrowIfNull(protection);
        ArgumentNullException.ThrowIfNull(preferences);
        ArgumentNullException.ThrowIfNull(presenter);

        // Flip first: the writes below must agree with the state the button is about to show, even if
        // one of them is refused by the host.
        bool lockNow = !preferences.ChartLocked;
        preferences.ChartLocked = lockNow;

        if (lockNow)
        {
            _ = shapes.LockOwnedShapes();
        }
        else
        {
            foreach (string primitiveId in shapes.ListOwned())
            {
                _ = shapes.UnlockShape(primitiveId);
            }
        }

        ProtectionWriteOutcome outcome = protection.Protect();
        if (outcome.Refusal is { } refusal)
        {
            presenter(
                lockNow
                    ? "The chart was locked but the sheet protection could not be re-applied: " + refusal + "."
                    : "The chart shapes were released but the sheet protection could not be updated: " + refusal + ".");
        }
    }
}
