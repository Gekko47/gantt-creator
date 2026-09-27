using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// The workbook recalculation modes the application-state scope can switch
/// between.
/// </summary>
/// <remarks>
/// The PIA's <c>XlCalculation</c> is an interop enum, and the AddIn must never
/// name an interop type (the CS0433 hazard documented on
/// <see cref="IExcelApplicationAdapter"/>). This is the Office-owned,
/// interop-free mirror, and <see cref="ExcelApplicationStateScope"/> maps it by
/// an exhaustive switch rather than by a numeric cast, so the two vocabularies
/// cannot drift silently.
/// </remarks>
public enum GanttCalculationMode
{
    /// <summary>Excel recalculates on every change.</summary>
    Automatic = 0,

    /// <summary>Excel recalculates automatically except for data tables.</summary>
    Semiautomatic = 1,

    /// <summary>Excel recalculates only when the user requests it.</summary>
    Manual = 2,
}

/// <summary>
/// One application-state setting captured for restoration, paired with the
/// write that changes it.
/// </summary>
/// <remarks>
/// A capture is recorded only when a value is <em>actually changed</em>, which
/// is the whole point of the scope: an unchanged property is never written on the
/// way in and never written on the way out, so a command cannot perturb a setting
/// it did not touch.
/// </remarks>
internal sealed record ApplicationStateCapture<T>(T OriginalValue, Action<T> Restore);

/// <summary>
/// Saves and restores the Excel application settings a rendering command must
/// not leave changed: <c>ScreenUpdating</c>, <c>EnableEvents</c>,
/// <c>DisplayAlerts</c>, the calculation mode, the status bar, and the
/// selection.
/// </summary>
/// <remarks>
/// <para>
/// <b>D1 — capture on change.</b> Each setting is captured only when the command
/// actually changes it, and only a captured setting is restored. An unchanged
/// property is never written.
/// </para>
/// <para>
/// <b>D2 — <c>Dispose</c> has <c>finally</c> semantics.</b> Disposing restores
/// everything captured so far whether the caller used <c>using</c> or an
/// explicit <c>try</c>/<c>finally</c>, and a restore that throws cannot skip the
/// remaining restores.
/// </para>
/// <para>
/// <b>D3 — selection restore is best-effort and ordered last.</b>
/// <c>Application.Selection</c> is read-only (probed 2026-09-27), so restore is
/// a <c>Range.Select()</c> call, which is a COM call the host can reject and has
/// no property-write alternative. It therefore degrades with a technical record
/// and never masks an earlier restore.
/// </para>
/// <para>
/// <b>D4 — every property read and write is an injectable seam</b>, so the
/// contract tests can inject a failure at each one.
/// </para>
/// <para>
/// This type must run on the Excel main STA thread. It never marshals, never
/// offloads, and never wraps a call in <c>Task.Run</c>.
/// </para>
/// <para>
/// The type is deliberately unsealed so the contract tests can derive it and
/// substitute the <c>internal virtual</c> property seams, matching
/// <see cref="ExcelWorkbookInitialiser"/> and <see cref="ExcelWorksheetProtectionGuard"/>.
/// </para>
/// </remarks>
/// <param name="application">
/// The Excel application object (for example <c>ExcelDnaUtil.Application</c>), or
/// <see langword="null"/> when the host supplied none. A foreign object fails the
/// interface cast, and every member then degrades to "no change and nothing to
/// restore" rather than throwing.
/// </param>
/// <param name="technicalRecord">
/// Receives a technical record when a restore degrades, per D3. The caller owns
/// the operation ID; the scope reports what it could not put back. A
/// <see langword="null"/> sink discards the record, which is the safe direction:
/// a logging failure must never mask the restore.
/// </param>
public class ExcelApplicationStateScope(object? application, Action<string>? technicalRecord = null)
    : IDisposable
{
    private readonly List<Action> _restores = [];
    private readonly Excel.Application? _application = application as Excel.Application;
    private bool _disposed;

    /// <summary>Gets whether the scope found a real Excel application.</summary>
    internal bool IsAvailable => _application is not null;

    /// <summary>
    /// Turns screen updating off, capturing the previous value only when the host
    /// is currently updating.
    /// </summary>
    public void SuppressScreenUpdating()
    {
        if (!IsAvailable)
        {
            return;
        }

        if (GetScreenUpdating())
        {
            Record(GetScreenUpdating, SetScreenUpdating);
            SetScreenUpdating(false);
        }
    }

    /// <summary>
    /// Turns event dispatch off, capturing the previous value only when events
    /// are currently on.
    /// </summary>
    public void SuppressEvents()
    {
        if (!IsAvailable)
        {
            return;
        }

        if (GetEnableEvents())
        {
            Record(GetEnableEvents, SetEnableEvents);
            SetEnableEvents(false);
        }
    }

    /// <summary>
    /// Turns alert dialogs off, capturing the previous value only when alerts
    /// are currently shown.
    /// </summary>
    public void SuppressAlerts()
    {
        if (!IsAvailable)
        {
            return;
        }

        if (GetDisplayAlerts())
        {
            Record(GetDisplayAlerts, SetDisplayAlerts);
            SetDisplayAlerts(false);
        }
    }

    /// <summary>
    /// Hides the status bar, capturing the previous value only when it is
    /// currently visible.
    /// </summary>
    /// <remarks>
    /// Only <c>DisplayStatusBar</c> is touched. The <c>StatusBar</c> text is a
    /// separate setting with a separate contract: a command that wants to show
    /// progress sets it through <see cref="SetStatusBarText"/>, and a command
    /// that does not must leave both alone.
    /// </remarks>
    public void SuppressStatusBar()
    {
        if (!IsAvailable)
        {
            return;
        }

        if (GetDisplayStatusBar())
        {
            Record(GetDisplayStatusBar, SetDisplayStatusBar);
            SetDisplayStatusBar(false);
        }
    }

    /// <summary>
    /// Sets the status-bar text, capturing the previous text only when this call
    /// would actually change it.
    /// </summary>
    /// <param name="text">The text to display, or <see langword="null"/> to clear it.</param>
    public void SetStatusBarText(string? text)
    {
        if (!IsAvailable)
        {
            return;
        }

        var current = GetStatusBarText();
        if (string.Equals(current as string, text, StringComparison.Ordinal))
        {
            return;
        }

        Record(GetStatusBarText, SetStatusBarTextCore);
        SetStatusBarTextCore(text);
    }

    /// <summary>
    /// Sets the calculation mode, capturing the previous mode only when the host
    /// is not already in the requested one.
    /// </summary>
    /// <param name="mode">The mode to switch to.</param>
    public void SetCalculationMode(GanttCalculationMode mode)
    {
        // The argument is validated BEFORE the availability check: an undefined
        // mode is a caller defect, and returning silently because no Excel host
        // happens to be present would hide it. The reverse order would make the
        // guard unreachable in exactly the situation a unit test exercises.
        if (!Enum.IsDefined(mode))
        {
            throw new ArgumentOutOfRangeException(nameof(mode), mode, "Undefined calculation mode.");
        }

        if (!IsAvailable)
        {
            return;
        }

        GanttCalculationMode current = GetCalculationMode();
        if (current == mode)
        {
            return;
        }

        Record(GetCalculationMode, SetCalculationModeCore);
        SetCalculationModeCore(mode);
    }

    /// <summary>
    /// Captures the current selection so it can be restored after the command.
    /// </summary>
    /// <remarks>
    /// This is a separate call rather than something <see cref="SuppressAlerts"/>
    /// and friends do implicitly, because D3 orders the selection restore last and
    /// separately: a command that never touched the selection must not cause a
    /// restore that could move the user's cursor.
    /// </remarks>
    public void CaptureSelection()
    {
        if (!IsAvailable || _selectionCaptured)
        {
            // Already captured: the FIRST selection is the one the user had, and
            // re-capturing after the command has already moved the cursor would
            // make the scope "restore" wherever the command left it.
            return;
        }

        var selection = GetSelection();
        if (selection is null)
        {
            return;
        }

        _selection = selection;
        _selectionCaptured = true;
    }

    private object? _selection;
    private bool _selectionCaptured;

    /// <summary>
    /// Reads a setting and registers the restore that puts it back.
    /// </summary>
    /// <typeparam name="T">The setting's value type.</typeparam>
    /// <param name="read">Reads the current value.</param>
    /// <param name="write">Writes a value back.</param>
    /// <remarks>
    /// The restore is registered <em>before</em> the command changes the setting,
    /// so a failure in the very next line still leaves the setting restorable. A
    /// read that throws leaves nothing registered, which is correct: nothing was
    /// changed, so nothing needs putting back.
    /// </remarks>
    private void Record<T>(Func<T> read, Action<T> write)
    {
        T original = read();
        _restores.Add(() => write(original));
    }

    /// <summary>
    /// Restores every captured setting, then the selection last.
    /// </summary>
    /// <remarks>
    /// D2: each restore is independently guarded, so one that throws cannot skip
    /// the rest. D3: the selection is restored after every other setting, and its
    /// failure is recorded rather than propagated, because it is the one restore
    /// with no property-write path and therefore the one the host is most likely
    /// to reject. The first failure is preserved for the technical record; later
    /// ones do not mask it.
    /// </remarks>
    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Performs the restore.
    /// </summary>
    /// <param name="disposing">
    /// <see langword="true"/> when called from <see cref="Dispose()"/>. The scope
    /// holds no unmanaged resource of its own, so the value is not branched on;
    /// the member exists because the type is unsealed for the test seams.
    /// </param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        // Reverse order: the last setting changed is the first one put back, which
        // matches the COM ownership rule of releasing in reverse acquisition order.
        for (var index = _restores.Count - 1; index >= 0; index--)
        {
            Restore(_restores[index], "an Excel application setting");
        }

        _restores.Clear();

        if (_selectionCaptured && _selection is not null)
        {
            // D3: the selection is restored after every other setting, and its
            // failure is recorded rather than propagated, because it is the one
            // restore with no property-write path and therefore the one the host is
            // most likely to reject.
            // CA1031: the selection restore is best-effort by contract (D3). It
            // runs last, inside the same Dispose, and its failure must not skip
            // the restores above or escape into Excel.
#pragma warning disable CA1031
            try
            {
                RestoreSelection(_selection);
            }
            catch (Exception exception)
            {
                Report("Gantt Creator could not restore the Excel selection: " + exception.Message);
            }
#pragma warning restore CA1031
            finally
            {
                _selection = null;
                _selectionCaptured = false;
            }
        }
    }

    /// <summary>
    /// Runs one restore, degrading to a technical record when it throws.
    /// </summary>
    /// <param name="restore">The restore to run.</param>
    /// <param name="subject">What the restore puts back, for the record.</param>
    /// <remarks>
    /// D2: this guard is what makes one failing restore unable to skip the rest. A
    /// rejected COM call surfaces as a <c>COMException</c>, but a proxy can
    /// also report it as <see cref="InvalidOperationException"/> or
    /// <see cref="NotSupportedException"/>, so the catch is deliberately broader
    /// than one type.
    /// </remarks>
    private void Restore(Action restore, string subject)
    {
        // CA1031: a restore failure must never skip the remaining restores, must
        // never escape a finally block into Excel, and degrades to a record.
#pragma warning disable CA1031
        try
        {
            restore();
        }
        catch (Exception exception)
        {
            Report("Gantt Creator could not restore " + subject + ": " + exception.Message);
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Writes a technical record, degrading to silence when the sink fails.
    /// </summary>
    /// <param name="message">The record to write.</param>
    /// <remarks>
    /// CA1031: a logging failure must never mask the restore it is reporting, and
    /// must never escape a finally block into Excel.
    /// </remarks>
    private void Report(string message)
    {
#pragma warning disable CA1031
        try
        {
            technicalRecord?.Invoke(message);
        }
        catch
        {
            // Intentionally empty: the safe direction is "no record" rather than
            // an exception escaping a finally block into Excel.
        }
    }
    // ---- Test seams (internal virtual, so each property read/write is injectable) ----

    /// <summary>Reads <c>Application.ScreenUpdating</c>. Test seam.</summary>
    /// <returns>The current screen-updating state.</returns>
    internal virtual bool GetScreenUpdating() => _application!.ScreenUpdating;

    /// <summary>Writes <c>Application.ScreenUpdating</c>. Test seam.</summary>
    /// <param name="value">The value to write.</param>
    internal virtual void SetScreenUpdating(bool value) => _application!.ScreenUpdating = value;

    /// <summary>Reads <c>Application.EnableEvents</c>. Test seam.</summary>
    /// <returns>The current event state.</returns>
    internal virtual bool GetEnableEvents() => _application!.EnableEvents;

    /// <summary>Writes <c>Application.EnableEvents</c>. Test seam.</summary>
    /// <param name="value">The value to write.</param>
    internal virtual void SetEnableEvents(bool value) => _application!.EnableEvents = value;

    /// <summary>Reads <c>Application.DisplayAlerts</c>. Test seam.</summary>
    /// <returns>The current alert state.</returns>
    internal virtual bool GetDisplayAlerts() => _application!.DisplayAlerts;

    /// <summary>Writes <c>Application.DisplayAlerts</c>. Test seam.</summary>
    /// <param name="value">The value to write.</param>
    internal virtual void SetDisplayAlerts(bool value) => _application!.DisplayAlerts = value;

    /// <summary>Reads <c>Application.DisplayStatusBar</c>. Test seam.</summary>
    /// <returns>Whether the status bar is visible.</returns>
    internal virtual bool GetDisplayStatusBar() => _application!.DisplayStatusBar;

    /// <summary>Writes <c>Application.DisplayStatusBar</c>. Test seam.</summary>
    /// <param name="value">The value to write.</param>
    internal virtual void SetDisplayStatusBar(bool value) => _application!.DisplayStatusBar = value;

    /// <summary>Reads <c>Application.StatusBar</c>. Test seam.</summary>
    /// <returns>The status-bar text, or <see langword="null"/> when the host holds none.</returns>
    /// <remarks>
    /// The PIA types this member <c>Object</c>, not <c>string</c>, so the comparison
    /// happens at this boundary and the captured value keeps its original form.
    /// </remarks>
    internal virtual object? GetStatusBarText() => _application!.StatusBar;

    /// <summary>Writes <c>Application.StatusBar</c>. Test seam.</summary>
    /// <param name="value">The value to write.</param>
    internal virtual void SetStatusBarTextCore(object? value) => _application!.StatusBar = value;

    /// <summary>Reads <c>Application.Calculation</c> as the interop-free mode. Test seam.</summary>
    /// <returns>The current calculation mode.</returns>
    /// <remarks>
    /// Read through the <c>_Application</c> interface reference, never the
    /// coclass. The installed PIA declares <c>Calculation</c> only on the
    /// interface, and reading an enum-typed property through the
    /// <c>Application</c> coclass proxy raises
    /// <c>DISP_E_TYPEMISMATCH (0x80020005)</c> on the live host — observed
    /// 2026-09-27 by this row's Office gate, which is exactly the class of
    /// defect the live gate exists to catch and which no metadata probe can
    /// reveal, because the member is present and read/write on paper.
    /// </remarks>
    internal virtual GanttCalculationMode GetCalculationMode() =>
        ToMode(InteropApplication.Calculation);

    /// <summary>Writes <c>Application.Calculation</c>. Test seam.</summary>
    /// <param name="mode">The mode to write.</param>
    /// <remarks>
    /// Written through the same <c>_Application</c> interface reference as the
    /// read, for the same live-host reason.
    /// </remarks>
    internal virtual void SetCalculationModeCore(GanttCalculationMode mode) =>
        InteropApplication.Calculation = ToInterop(mode);

    /// <summary>
    /// The application as its PIA <em>interface</em>, which is the form the
    /// calculation members must be reached through.
    /// </summary>
    /// <exception cref="InvalidOperationException">
    /// The host supplied an object that is not an Excel application interface.
    /// Reaching this point means the scope was constructed with a usable
    /// application, so the cast cannot legitimately fail.
    /// </exception>
    private Excel._Application InteropApplication =>
        _application as Excel._Application
        ?? throw new InvalidOperationException("The application object does not expose the Excel interface.");

    /// <summary>Reads <c>Application.Selection</c>. Test seam.</summary>
    /// <returns>The current selection, or <see langword="null"/>.</returns>
    internal virtual object? GetSelection() => _application!.Selection;

    /// <summary>
    /// Restores the selection. Test seam over <c>Range.Select()</c>.
    /// </summary>
    /// <param name="selection">The previously captured selection.</param>
    /// <remarks>
    /// <c>Application.Selection</c> has no setter (probed 2026-09-27:
    /// <c>canWrite = false</c>), and neither <c>Application.Select</c> nor
    /// <c>Worksheet.Select</c> exists on the PIA interfaces, so the only restore
    /// path is <c>Range.Select()</c> on the captured range.
    /// </remarks>
    internal virtual void RestoreSelection(object selection)
    {
        if (selection is Excel.Range range)
        {
            range.Select();
        }
    }

    /// <summary>Maps the interop calculation enum by name, never by a numeric cast.</summary>
    /// <param name="value">The interop enum value.</param>
    /// <returns>The interop-free mode.</returns>
    private static GanttCalculationMode ToMode(Excel.XlCalculation value) => value switch
    {
        Excel.XlCalculation.xlCalculationAutomatic => GanttCalculationMode.Automatic,
        Excel.XlCalculation.xlCalculationSemiautomatic => GanttCalculationMode.Semiautomatic,
        Excel.XlCalculation.xlCalculationManual => GanttCalculationMode.Manual,
        _ => throw new ArgumentOutOfRangeException(nameof(value), value, "Unknown Excel calculation mode."),
    };

    /// <summary>Maps the interop-free mode onto the interop calculation enum.</summary>
    /// <param name="mode">The interop-free mode.</param>
    /// <returns>The interop enum value.</returns>
    private static Excel.XlCalculation ToInterop(GanttCalculationMode mode) => mode switch
    {
        GanttCalculationMode.Automatic => Excel.XlCalculation.xlCalculationAutomatic,
        GanttCalculationMode.Semiautomatic => Excel.XlCalculation.xlCalculationSemiautomatic,
        GanttCalculationMode.Manual => Excel.XlCalculation.xlCalculationManual,
        _ => throw new ArgumentOutOfRangeException(nameof(mode), mode, "Undefined calculation mode."),
    };
}
