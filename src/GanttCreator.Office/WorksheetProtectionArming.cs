namespace GanttCreator.Office;

/// <summary>
/// The policy that applies the add-in's own protection to the Gantt worksheet, once per session
/// (ADR-0055).
/// </summary>
/// <remarks>
/// <para>
/// <b>Why this is its own type rather than a method on the orchestrator.</b> It is the one place in
/// the add-in that mutates protection state, and the architecture gate's rule is that a file holding
/// raw workbook mutations is an adapter that must be classified. Keeping the calls here means the
/// orchestrator stays a pipeline — it contains no mutation pattern of its own — and this file is
/// declared as the protection-control exception for the reason below.
/// </para>
/// <para>
/// <b>It deliberately does not consult <see cref="IWorksheetProtectionGuard"/>.</b> The guard refuses
/// a protected worksheet, and the whole job here is to recover exactly that state: a sheet protected
/// in an earlier session, whose <c>UserInterfaceOnly</c> did not survive the reopen, must be released
/// and protected again before anything else can write. A guard-first rule would make the recovery
/// unreachable. This is the same declared exception recorded for
/// <see cref="ExcelWorksheetProtectionPort"/>.
/// </para>
/// <para>
/// It touches COM only through the two ports it is given, so the policy — order, idempotence,
/// release-on-opt-out, and the decision not to refuse the refresh — is contract-testable without a
/// host.
/// </para>
/// </remarks>
public sealed class WorksheetProtectionArming(
    IWorksheetProtectionPort port,
    IShapeWritePort shapeWriter,
    ProtectionArmingState state,
    Action<string>? trace = null)
{
    private readonly IWorksheetProtectionPort _port = port ?? throw new ArgumentNullException(nameof(port));
    private readonly IShapeWritePort _shapeWriter = shapeWriter ?? throw new ArgumentNullException(nameof(shapeWriter));
    private readonly ProtectionArmingState _state = state ?? throw new ArgumentNullException(nameof(state));
    private readonly Action<string>? _trace = trace;

    /// <summary>Gets whether the add-in's protection is already in force this session.</summary>
    /// <remarks>
    /// The refresh claims ownership of the sheet's protection only when this is true: a protected
    /// sheet the add-in has not armed is the user's, and the existing `TargetProtected` refusal is
    /// the right answer for it.
    /// </remarks>
    public bool IsArmed => _state.IsArmed;

    /// <summary>
    /// Applies the add-in's own protection when the workbook asks for it, or releases it when the
    /// workbook has opted out and it was armed.
    /// </summary>
    /// <param name="protects">Whether the configuration the refresh already read asks for it.</param>
    /// <remarks>
    /// <para>
    /// The order is <b>read → release (if protected) → lock the shapes → protect</b>. The release is
    /// required because <c>UserInterfaceOnly</c> is not persisted, so a sheet protected in an earlier
    /// session still refuses this add-in's writes. The lock runs between the release and the protect
    /// because the shape writer's own guard refuses a protected sheet — this is the only window in
    /// which shapes created before ADR-0055 can be reached.
    /// </para>
    /// <para>
    /// Nothing here refuses the refresh: a protection write the host rejects loses the lock, not the
    /// chart, and the state stays unarmed so the next pass tries again. An opt-out after the sheet
    /// was armed releases it, because the sheet is then the user's to manage.
    /// </para>
    /// </remarks>
    public void Arm(bool protects)
    {
        if (!protects)
        {
            if (_state.IsArmed)
            {
                Record(_port.Unprotect(), "release");
                _state.Invalidate();
            }

            return;
        }

        if (_state.IsArmed)
        {
            return;
        }

        WorksheetProtectionState? live = _port.Read();
        if (live is { } current && (current.ContentsProtected || current.ShapesProtected))
        {
            Record(_port.Unprotect(), "reprotect");
        }

        Record(_shapeWriter.LockOwnedShapes(), "lock");

        ProtectionWriteOutcome applied = _port.Protect();
        Record(applied, "protect");
        if (applied.Succeeded)
        {
            _state.MarkArmed();
        }
    }

    /// <summary>Records a protection write that did not succeed, or nothing when it did.</summary>
    /// <param name="outcome">The typed result of the write.</param>
    /// <param name="step">The step name, for example <c>protect</c>.</param>
    private void Record(ProtectionWriteOutcome outcome, string step)
    {
        if (outcome.Refusal is { } refusal)
        {
            _trace?.Invoke("gantt.protection " + step + " refused=" + refusal);
        }
    }

    /// <summary>Records a shape-lock sweep that did not succeed, or nothing when it did.</summary>
    /// <param name="outcome">The typed result of the sweep.</param>
    /// <param name="step">The step name, for example <c>lock</c>.</param>
    private void Record(ShapeWriteOutcome outcome, string step)
    {
        if (outcome.Refusal is { } refusal)
        {
            _trace?.Invoke("gantt.protection " + step + " refused=" + refusal);
        }
    }
}
