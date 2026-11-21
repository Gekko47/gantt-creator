using System.Diagnostics;
using System.Globalization;

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
    /// Whether the outermost scope unlocks the table's cells when it re-applies protection.
    /// </summary>
    /// <remarks>
    /// Recorded at depth 0, so a nested window cannot change the decision the outermost one made. An
    /// automatic pass declines it: the unlock measured 7.3 ms median of a 9.7 ms close, nothing
    /// re-locks those cells, and the explicit Refresh repairs a cell the user locked.
    /// </remarks>
    private bool _repairCellsOnClose = true;

    /// <summary>
    /// Whether the outermost scope may release protection the add-in did not arm, and must therefore
    /// put it back on the way out.
    /// </summary>
    /// <remarks>
    /// Recorded at depth 0 for the same reason as <see cref="_repairCellsOnClose"/>. Only the explicit
    /// Fix command asks for it (ADR-0011 amendment 2026-11-21): a workbook reopened from a previous
    /// session comes up protected — <c>UserInterfaceOnly</c> is not persisted — so the one command that
    /// exists to repair that workbook would otherwise demand a manual unprotect first.
    /// </remarks>
    private bool _releaseForeignProtection;

    /// <summary>
    /// Whether this window actually released protection it did not arm, so its close re-applies it.
    /// </summary>
    /// <remarks>
    /// Set only from a release that succeeded. A refused release leaves the flag clear, so the close
    /// does not protect a sheet it never opened — the caller sees the host refuse its write instead,
    /// and the repairer turns that into its own typed refusal before writing anything.
    /// </remarks>
    private bool _releasedForeignProtection;

    /// <summary>
    /// Opens the window: the inner scope runs with the sheet released, and the outermost scope
    /// re-applies protection when it disposes.
    /// </summary>
    /// <param name="repairCells">
    /// Whether the close unlocks the visible table's cells first, passed through to
    /// <see cref="IWorksheetProtectionPort.Protect"/>.
    /// </param>
    /// <param name="releaseForeignProtection">
    /// <see langword="true"/> to also release protection the add-in has not armed, and re-apply it on
    /// close. Used only by the explicit Fix command (ADR-0011 amendment 2026-11-21).
    /// </param>
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
    /// <para>
    /// <b>Foreign protection.</b> Arming answers "is this protection ours"; the repair path cannot ask
    /// that question, because the protection it must get past belongs to a previous session. It
    /// therefore passes <paramref name="releaseForeignProtection"/> and the window releases whatever
    /// protection the port reports — the add-in protects without a password, so a release it cannot
    /// perform is a protection it must not silently write through.
    /// </para>
    /// </remarks>
    public IDisposable Open(bool repairCells = true, bool releaseForeignProtection = false)
    {
        if (_depth == 0)
        {
            _repairCellsOnClose = repairCells;
            _releaseForeignProtection = releaseForeignProtection;
            _releasedForeignProtection = false;
        }

        bool armed = _state.IsArmed;
        bool foreign = _depth == 0 && _releaseForeignProtection && !armed && IsPortProtected();
        if (_depth == 0 && (armed || foreign))
        {
            // The release is one host call and it is on the click-to-plot path, so its cost is
            // recorded rather than inferred: `command - total - post` attributed ~31-54 ms per
            // converged pass to "the window or the ribbon", and this separates the two.
            long started = Stopwatch.GetTimestamp();
            ProtectionWriteOutcome released = _port.Unprotect();
            Record(released, "window release");
            _releasedForeignProtection = foreign && released.Succeeded;
            _trace?.Invoke(string.Create(
                CultureInfo.InvariantCulture,
                $"gantt.protection window open released={released.Succeeded} ms={Stopwatch.GetElapsedTime(started).TotalMilliseconds:0.#}"));
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
        if (_depth == 0 && (_state.IsArmed || _releasedForeignProtection))
        {
            long started = Stopwatch.GetTimestamp();
            ProtectionWriteOutcome protectedOutcome = _port.Protect(_repairCellsOnClose);
            Record(protectedOutcome, "window protect");
            _trace?.Invoke(string.Create(
                CultureInfo.InvariantCulture,
                $"gantt.protection window close protected={protectedOutcome.Succeeded} ms={Stopwatch.GetElapsedTime(started).TotalMilliseconds:0.#}"));
        }
    }

    /// <summary>Whether the port reports a protected sheet, without changing it.</summary>
    /// <returns><see langword="true"/> when either protection flag is set.</returns>
    /// <remarks>
    /// Read only when the caller asked for a foreign release: the arm path already knows the answer
    /// (the add-in armed it), so it spends no host call on this.
    /// </remarks>
    private bool IsPortProtected() =>
        _port.Read() is { } state && (state.ContentsProtected || state.ShapesProtected);

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
