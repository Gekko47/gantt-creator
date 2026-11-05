using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// Re-locks every owned shape, re-applies the chart's protection, and forgets the unlocks — the
/// manual counterpart of what an explicit Refresh does (ADR-0055 clause 6).
/// </summary>
/// <remarks>
/// <para>
/// <b>Why this can be a command rather than only a refresh.</b> The unlock is transient and the
/// scheduler is suspended while it is in effect, so a user who has finished dragging needs a way to
/// put the lock back without triggering a full rebuild. This is that way; an explicit Refresh does the
/// same and additionally re-asserts every shape's geometry from the scene.
/// </para>
/// <para>
/// It does not refresh. Re-locking is not a reason to re-read the whole table, and a user who wants the
/// manual edit reverted asks for Refresh.
/// </para>
/// </remarks>
internal static class LockShapesCommand
{
    /// <summary>Runs the production lock in the current Excel session.</summary>
    internal static void RunForExcel()
    {
        object? application = ExcelDna.Integration.ExcelDnaUtil.Application;

        Run(
            new ExcelShapeWriter(application, new ExcelWorksheetProtectionGuard(application)),
            new ExcelWorksheetProtectionPort(application),
            ShapeUnlockState.Shared,
            ProtectionArmingState.Shared,
            CommandErrorDialog.Show);

        RibbonStateService.Instance.InvalidateControl(RibbonControlIds.UnlockShape);
        RibbonStateService.Instance.InvalidateControl(RibbonControlIds.LockShapes);
    }

    /// <summary>Runs an injected lock, for tests and for the R4.9 wiring.</summary>
    /// <param name="shapes">Locks every owned shape.</param>
    /// <param name="protection">Re-applies the chart's protection.</param>
    /// <param name="unlockState">Forgotten, which re-enables the automatic refresh.</param>
    /// <param name="arming">Marked armed, because this command is what applied the protection.</param>
    /// <param name="presenter">The user-safe refusal presenter.</param>
    internal static void Run(
        IShapeWritePort shapes,
        IWorksheetProtectionPort protection,
        ShapeUnlockState unlockState,
        ProtectionArmingState arming,
        Action<string> presenter)
    {
        ArgumentNullException.ThrowIfNull(shapes);
        ArgumentNullException.ThrowIfNull(protection);
        ArgumentNullException.ThrowIfNull(unlockState);
        ArgumentNullException.ThrowIfNull(arming);
        ArgumentNullException.ThrowIfNull(presenter);

        ShapeWriteOutcome locked = shapes.LockOwnedShapes();
        if (!locked.Succeeded)
        {
            Present(presenter, "The chart's shapes could not be locked. Try again.");
            return;
        }

        // The unlocks are forgotten only once the shapes are locked again: clearing first would
        // re-enable the automatic refresh while a shape is still writable.
        _ = unlockState.Clear();
        arming.MarkArmed();

        ProtectionWriteOutcome applied = protection.Protect();
        if (!applied.Succeeded)
        {
            Present(
                presenter,
                "The shapes are locked, but the chart could not be protected. " +
                "Refresh to try again.");
        }
    }

    /// <summary>Shows the message, degrading to no dialog if the host refuses one.</summary>
    /// <param name="presenter">The user-facing presenter.</param>
    /// <param name="message">The message to show.</param>
    private static void Present(Action<string> presenter, string message)
    {
        // CA1031: a presenter failure is outside command behaviour and must not escape into Excel.
#pragma warning disable CA1031
        try
        {
            presenter(message);
        }
        catch
        {
            // Intentionally empty: the message is lost; the command has already reported.
        }
#pragma warning restore CA1031
    }
}
