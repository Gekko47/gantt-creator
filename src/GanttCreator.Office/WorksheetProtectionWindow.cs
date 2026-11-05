namespace GanttCreator.Office;

/// <summary>
/// Releases the add-in's own sheet protection for the duration of a write, and re-applies it
/// afterwards (ADR-0055).
/// </summary>
/// <remarks>
/// <para>
/// <b>Why a window rather than "protect and let code through".</b> Measured live 2026-11-04 on Excel
/// 16.0.20527.20112: with the Gantt sheet protected for the interface only, a cell write lands but
/// <c>Shapes.AddShape</c> and <c>Shape.Left</c> raise "The specified value is out of range" and a
/// table-row insert is refused. <c>UserInterfaceOnly</c> therefore exempts cells only, so a refresh —
/// which creates and moves shapes — and every row-structure command must run with the sheet
/// RELEASED. The evidence is <c>ProtectionArmingIntegrationTests</c>.
/// </para>
/// <para>
/// <b>The cost is two host calls per write phase, and only when the feature is on.</b>
/// <see cref="Open"/> is inert when the add-in has not armed protection, so a workbook that opted out
/// pays a single bool read. <see cref="ProtectionArmingState.IsArmed"/> keeps its meaning — "this
/// protection is ours" — which is what makes the re-apply correct: the state is not cleared when the
/// sheet is released, only when the add-in stops owning it.
/// </para>
/// <para>
/// Re-entrant by depth, because a command can call the refresh (Add activity inserts a row, then
/// renders). Only the outermost scope releases and re-applies, so a nested window cannot leave the
/// sheet unprotected when the outer one finishes.
/// </para>
/// </remarks>
public sealed class WorksheetProtectionWindow(
    IWorksheetProtectionPort port,
    ProtectionArmingState state,
    Action<string>? trace = null)
{
    private readonly IWorksheetProtectionPort _port = port ?? throw new ArgumentNullException(nameof(port));
    private readonly ProtectionArmingState _state = state ?? throw new ArgumentNullException(nameof(state));
    private readonly Action<string>? _trace = trace;
    private int _depth;

    /// <summary>
    /// Opens the window: the inner scope runs with the sheet released, and the outermost scope
    /// re-applies protection when it disposes.
    /// </summary>
    /// <returns>The scope that closes the window.</returns>
    /// <remarks>
    /// <para>
    /// The scope is returned even when nothing is armed, because arming can happen INSIDE the window:
    /// the first refresh that arms releases the sheet as part of arming, and it still has to be
    /// protected again when that refresh finishes. When nothing is armed at either end, the cost is
    /// two bool reads and no host call.
    /// </para>
    /// <para>
    /// Disposing twice is harmless, and nesting is by depth: only the outermost scope releases and
    /// re-applies, so a command that calls the refresh cannot leave the sheet unprotected when the
    /// statement that started it finishes. A refused release is recorded and the write continues —
    /// the host then refuses the write and reports it, which is a truer signal than silence here.
    /// </para>
    /// </remarks>
    public IDisposable Open()
    {
        if (_depth == 0 && _state.IsArmed)
        {
            ProtectionWriteOutcome released = _port.Unprotect();
            Record(released, "window release");
            _trace?.Invoke("gantt.protection window open released=" + released.Succeeded);
        }

        _depth++;
        return new Scope(this);
    }

    private void Close()
    {
        if (_depth == 0)
        {
            return;
        }

        _depth--;
        if (_depth == 0 && _state.IsArmed)
        {
            ProtectionWriteOutcome protectedOutcome = _port.Protect();
            Record(protectedOutcome, "window protect");
            _trace?.Invoke("gantt.protection window close protected=" + protectedOutcome.Succeeded);
        }
    }

    /// <summary>Records a protection write that did not succeed, or nothing when it did.</summary>
    /// <param name="outcome">The typed result of the write.</param>
    /// <param name="step">The step name, for example <c>window release</c>.</param>
    private void Record(ProtectionWriteOutcome outcome, string step)
    {
        if (outcome.Refusal is { } refusal)
        {
            _trace?.Invoke("gantt.protection " + step + " refused=" + refusal);
        }
    }

    /// <summary>One window scope.</summary>
    private sealed class Scope(WorksheetProtectionWindow owner) : IDisposable
    {
        private WorksheetProtectionWindow? _owner = owner;

        /// <inheritdoc />
        public void Dispose()
        {
            WorksheetProtectionWindow? owner = _owner;
            _owner = null;
            owner?.Close();
        }
    }

}
