namespace GanttCreator.Office;

/// <summary>
/// The application-state settings a render command must not leave changed
/// (ADR-0020, R4.8A D8).
/// </summary>
/// <remarks>
/// <para>
/// <b>Why this is a port rather than a direct <see cref="ExcelApplicationStateScope"/>
/// call.</b> The live scope is a COM-facing class that cannot be constructed without
/// an Excel application, so depending on it directly would make the orchestrator
/// untestable without a host — and the whole point of extracting the orchestrator
/// was that D3 (validate before mutate) is only provable when the pipeline is
/// reachable from a test. The port is the seam; the live scope implements it.
/// </para>
/// <para>
/// <b>Why the operations are separate rather than one <c>Begin</c>/<c>End</c> pair.</b>
/// The orchestrator needs to suppress three settings, write a status message, and
/// capture the selection, and a caller that forgot one of those would silently skip
/// a restore. Named operations make an omission visible at the call site, and each
/// is independently assertable — a test can prove the status bar was restored
/// without reasoning about the others.
/// </para>
/// <para>
/// <b>Not every member is invoked by every caller.</b> The scope supports five
/// settings; which of them a command uses is its own decision. Refresh, for example,
/// writes status-bar text without calling <see cref="SuppressStatusBar"/>, because
/// bar visibility is the user's Excel setting rather than the add-in's (ADR-0020 D4).
/// The members stay here so a caller that does want the bar hidden can ask, and so
/// the capture/restore of <c>DisplayStatusBar</c> remains tested.
/// </para>
/// <para>
/// Every member is best-effort by contract: a host that refuses one of them must
/// not abort the refresh, because the alternative is refusing to draw a chart
/// because the status bar was busy. Restoration failures are reported through the
/// scope's own technical-record sink, never by throwing.
/// </para>
/// </remarks>
public interface IApplicationStateScope : IDisposable
{
    /// <summary>Turns screen updating off, capturing the previous value.</summary>
    void SuppressScreenUpdating();

    /// <summary>Turns application events off, capturing the previous value.</summary>
    void SuppressEvents();

    /// <summary>Turns alerts off, capturing the previous value.</summary>
    void SuppressAlerts();

    /// <summary>Hides the status bar, capturing the previous value.</summary>
    void SuppressStatusBar();

    /// <summary>Sets the status-bar text, capturing the previous value.</summary>
    /// <param name="text">The text to display while the refresh runs.</param>
    void SetStatusBarText(string? text);

    /// <summary>Captures the current selection so it can be restored afterwards.</summary>
    void CaptureSelection();
}

/// <summary>
/// The no-op scope, used when no live Excel application is available and in tests
/// that are not about state restoration.
/// </summary>
/// <remarks>
/// A genuine no-op rather than a null check at each call site. The alternative —
/// nullable scope plus a null check on all six operations — spreads one fact across
/// six places and makes the refresh harder to read than the thing it restores.
/// </remarks>
internal sealed class NullApplicationStateScope : IApplicationStateScope
{
    /// <summary>The shared instance; the type is stateless.</summary>
    public static NullApplicationStateScope Instance { get; } = new();

    private NullApplicationStateScope()
    {
    }

    public void SuppressScreenUpdating()
    {
    }

    public void SuppressEvents()
    {
    }

    public void SuppressAlerts()
    {
    }

    public void SuppressStatusBar()
    {
    }

    public void SetStatusBarText(string? text)
    {
    }

    public void CaptureSelection()
    {
    }

    public void Dispose() =>
        // Nothing was captured, so there is nothing to restore. Dispose is still
        // called, so the `using` above behaves identically with this scope and a
        // live one.
        GC.SuppressFinalize(this);
}
