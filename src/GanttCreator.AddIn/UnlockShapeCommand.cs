using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// Releases the chart's protection and clears one owned shape's lock, so the user can drag that
/// shape deliberately (ADR-0055 clause 6).
/// </summary>
/// <remarks>
/// <para>
/// <b>Transient by design.</b> The unlock is recorded in <see cref="ShapeUnlockState"/>, which
/// suspends the automatic refresh while it is non-empty; the next explicit Refresh, or the Lock
/// shapes command, re-locks, re-protects and re-asserts the shape's geometry — reverting a manual edit
/// exactly as the ruling accepted. Ownership partiality is what ADR-0055 rejected: the invariant "every
/// owned shape equals the scene" stays total, and only the trigger is suspended.
/// </para>
/// <para>
/// The command refuses a selection it cannot prove is ours, because unlocking is the one operation
/// that hands the user a writable chart element.
/// </para>
/// </remarks>
internal static class UnlockShapeCommand
{
    /// <summary>Runs the production unlock in the current Excel session.</summary>
    internal static void RunForExcel()
    {
        object? application = ExcelDna.Integration.ExcelDnaUtil.Application;

        Run(
            new ExcelSelectedShapeReader(application),
            new ExcelShapeWriter(application, new ExcelWorksheetProtectionGuard(application)),
            new ExcelWorksheetProtectionPort(application),
            ShapeUnlockState.Shared,
            ProtectionArmingState.Shared,
            CommandErrorDialog.Show);

        RibbonStateService.Instance.InvalidateControl(RibbonControlIds.UnlockShape);
        RibbonStateService.Instance.InvalidateControl(RibbonControlIds.LockShapes);
    }

    /// <summary>Runs an injected unlock, for tests and for the R4.9 wiring.</summary>
    /// <param name="selection">Resolves the selected owned shape.</param>
    /// <param name="shapes">Clears the shape's lock.</param>
    /// <param name="protection">Releases the sheet so the shape can actually be dragged.</param>
    /// <param name="unlockState">Records the unlock, which suspends the automatic refresh.</param>
    /// <param name="arming">Cleared so the next refresh re-arms and re-locks.</param>
    /// <param name="presenter">The user-safe refusal presenter.</param>
    internal static void Run(
        ISelectedShapePort selection,
        IShapeWritePort shapes,
        IWorksheetProtectionPort protection,
        ShapeUnlockState unlockState,
        ProtectionArmingState arming,
        Action<string> presenter)
    {
        ArgumentNullException.ThrowIfNull(selection);
        ArgumentNullException.ThrowIfNull(shapes);
        ArgumentNullException.ThrowIfNull(protection);
        ArgumentNullException.ThrowIfNull(unlockState);
        ArgumentNullException.ThrowIfNull(arming);
        ArgumentNullException.ThrowIfNull(presenter);

        SelectedShapeOutcome selected = selection.Read();
        if (!selected.Succeeded)
        {
            Present(presenter, Describe(selected.Refusal ?? SelectedShapeRefusalReason.ShapeNotOwned));
            return;
        }

        string primitiveId = selected.PrimitiveId!;
        ShapeWriteOutcome unlocked = shapes.UnlockShape(primitiveId);
        if (!unlocked.Succeeded)
        {
            Present(
                presenter,
                unlocked.Refusal == ShapeWriteRefusal.TargetProtected
                    ? "The chart's protection refused the change, so the shape was not unlocked."
                    : "That shape was not drawn by Gantt Creator, so it cannot be unlocked.");
            return;
        }

        // Order matters: record the unlock BEFORE releasing the sheet, so a release the host refuses
        // still leaves the scheduler suspended rather than rebuilding the chart over a half-done edit.
        _ = unlockState.Add(primitiveId);
        arming.Invalidate();

        ProtectionWriteOutcome released = protection.Unprotect();
        if (!released.Succeeded)
        {
            Present(
                presenter,
                "The chart's protection could not be released, so the shape is still locked. " +
                "Unprotect the sheet by hand and try again.");
        }
    }

    private static string Describe(SelectedShapeRefusalReason refusal) => refusal switch
    {
        SelectedShapeRefusalReason.NoApplication =>
            "Excel is not available, so the selection could not be read.",
        SelectedShapeRefusalReason.SelectionIsNotASingleShape =>
            "Select one chart shape that Gantt Creator drew, then try again.",
        SelectedShapeRefusalReason.ShapeNotOwned =>
            "That shape was not drawn by Gantt Creator, so it cannot be unlocked.",
        SelectedShapeRefusalReason.HostRejected =>
            "The selected shape could not be read. Try again.",
        _ => "Select one chart shape that Gantt Creator drew, then try again.",
    };

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
