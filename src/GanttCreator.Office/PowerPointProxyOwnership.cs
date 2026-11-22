using System.Globalization;
using System.Runtime.InteropServices;

namespace GanttCreator.Office;

/// <summary>
/// Tracks the PowerPoint COM proxies this add-in acquires and releases each exactly once, in
/// reverse acquisition order, with <see cref="Marshal.ReleaseComObject"/> (R7.1 D2).
/// </summary>
/// <remarks>
/// <para>
/// <b>Release, never force-release.</b> The tracker calls
/// <see cref="Marshal.ReleaseComObject"/> because a PowerPoint proxy has stable COM identity
/// and the RCW may be shared; <see cref="Marshal.FinalReleaseComObject"/> is never used,
/// matching the Excel precedent on <see cref="ExcelGanttTableReader"/>.
/// </para>
/// <para>
/// <b>Exactly once, in reverse order.</b> <see cref="ReleaseAll"/> snapshots the tracked set,
/// clears it before releasing, and walks it newest-first, so a second call releases nothing
/// and a release that throws cannot cause a repeat.
/// </para>
/// <para>
/// <b>Nothing it does not own.</b> Only proxies passed to <see cref="Acquire{T}"/> are
/// released. A non-COM value is ignored by the default release, so a contract test can pass a
/// plain object without touching the host.
/// </para>
/// <para>
/// <b>Not thread-safe, by design.</b> The tracker is an STA-only, per-operation type; sharing
/// it across threads would break the ordering guarantee it exists to provide.
/// </para>
/// </remarks>
public class PowerPointProxyOwnership
{
    private readonly List<object> _proxies = [];
    private readonly Action<string>? _technicalRecord;

    /// <summary>Creates an empty tracker.</summary>
    public PowerPointProxyOwnership()
        : this(null)
    {
    }

    /// <summary>Creates an empty tracker that reports a refused release once.</summary>
    /// <param name="technicalRecord">Receives one technical record per refused release, or <see langword="null"/>.</param>
    internal PowerPointProxyOwnership(Action<string>? technicalRecord)
    {
        _technicalRecord = technicalRecord;
    }

    /// <summary>Gets the number of proxies acquired and not yet released.</summary>
    public int Count => _proxies.Count;

    /// <summary>Records a proxy this add-in owns and returns it for chaining.</summary>
    /// <typeparam name="T">The proxy's static type; no interop type is required by this signature.</typeparam>
    /// <param name="proxy">The proxy to release when the operation finishes.</param>
    /// <returns>The same <paramref name="proxy"/>, so an acquisition can wrap a host read.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="proxy"/> is <see langword="null"/>.</exception>
    public T Acquire<T>(T proxy)
        where T : class
    {
        ArgumentNullException.ThrowIfNull(proxy);
        _proxies.Add(proxy);
        return proxy;
    }

    /// <summary>
    /// Releases every tracked proxy in reverse acquisition order, exactly once.
    /// </summary>
    /// <remarks>
    /// A refused release is recorded once and does not escape: the operation that acquired the
    /// proxy has already produced its answer, and a teardown failure must not become a second
    /// error. The exception, never swallowed silently, carries the HRESULT into the record.
    /// </remarks>
    public void ReleaseAll()
    {
        object[] proxies = [.. _proxies];
        _proxies.Clear();

        for (int index = proxies.Length - 1; index >= 0; index--)
        {
            try
            {
                _ = ReleaseCore(proxies[index]);
            }
            catch (Exception exception) when (exception is ArgumentException or COMException)
            {
                _technicalRecord?.Invoke(
                    "powerpoint release refused hresult=0x"
                    + exception.HResult.ToString("X8", CultureInfo.InvariantCulture));
            }
        }
    }

    /// <summary>Releases one proxy; a non-COM value is ignored.</summary>
    /// <param name="proxy">The tracked proxy.</param>
    /// <returns>The host's remaining reference count, or zero when no COM call was made.</returns>
    /// <remarks>
    /// This is the test seam the contract tests override to record the release order without a
    /// host. Keeping the <c>IsComObject</c> guard here, rather than in
    /// <see cref="ReleaseAll"/>, is what keeps the ordering behaviour testable with plain
    /// objects while production still refuses to hand a non-COM value to the marshaller.
    /// </remarks>
    internal virtual int ReleaseCore(object proxy) =>
        Marshal.IsComObject(proxy) ? Marshal.ReleaseComObject(proxy) : 0;
}
