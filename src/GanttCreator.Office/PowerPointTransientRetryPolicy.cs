using System.Collections.Frozen;
using System.Globalization;

namespace GanttCreator.Office;

/// <summary>
/// R7.5 D2: the measured transient-rejection filter, the explicit attempt cap and the
/// diagnostics budget for PowerPoint COM calls.
/// </summary>
/// <remarks>
/// <para>
/// <b>Evidence-first.</b> The filter is a frozen set populated from the R7.5 Step-0 busy-host
/// probe on the reference host; the three <see cref="DocumentedCandidates"/> are the work
/// item's recorded candidate codes, and the live busy-host body confirms or corrects them. An
/// HRESULT that is not in the set is never retried, so an unobserved code fails safe to
/// <see cref="PowerPointTransferFailureKind.HostRejected"/> rather than looping.
/// </para>
/// <para>
/// <b>Bounded, never a sleep.</b> The cap is an attempt count, not a wall-clock deadline; the
/// elapsed time is recorded for diagnostics only and is never a loop condition (work item
/// Step 0). No <c>Thread.Sleep</c>, spin or <c>Task.Run</c> exists anywhere on this path.
/// </para>
/// </remarks>
public sealed class PowerPointTransientRetryPolicy
{
    /// <summary>The <c>RPC_E_CALL_REJECTED</c> HRESULT (<c>0x80010001</c>) — a rejected call never executed.</summary>
    public static readonly int RpcCallRejected = unchecked((int)0x80010001);

    /// <summary>The <c>RPC_E_SERVERCALL_RETRYLATER</c> HRESULT (<c>0x8001010A</c>) — retry later.</summary>
    public static readonly int RpcServerCallRetryLater = unchecked((int)0x8001010A);

    /// <summary>The <c>RPC_E_SERVERCALL_REJECTED</c> HRESULT (<c>0x8001010B</c>) — the server rejected the call.</summary>
    public static readonly int RpcServerCallRejected = unchecked((int)0x8001010B);

    /// <summary>The starting attempt cap (three), confirmed by the Step-0 probe.</summary>
    public const int DefaultMaxAttempts = 3;

    /// <summary>
    /// The work item's documented transient candidates with <see cref="DefaultMaxAttempts"/>.
    /// The Phase-4 busy-host live body is the measurement that confirms or corrects this set.
    /// </summary>
    public static PowerPointTransientRetryPolicy DocumentedCandidates { get; } = new(
        [RpcCallRejected, RpcServerCallRetryLater, RpcServerCallRejected],
        DefaultMaxAttempts);

    /// <summary>Creates a policy over an explicit transient set and attempt cap.</summary>
    /// <param name="transientHResults">The measured transient HRESULT codes.</param>
    /// <param name="maxAttempts">The attempt cap; at least one.</param>
    /// <exception cref="ArgumentNullException"><paramref name="transientHResults"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="maxAttempts"/> is less than one.</exception>
    public PowerPointTransientRetryPolicy(IEnumerable<int> transientHResults, int maxAttempts)
    {
        ArgumentNullException.ThrowIfNull(transientHResults);
        ArgumentOutOfRangeException.ThrowIfLessThan(maxAttempts, 1);

        TransientHResults = transientHResults.ToFrozenSet();
        MaxAttempts = maxAttempts;
    }

    /// <summary>Gets the measured transient HRESULT set.</summary>
    public IReadOnlySet<int> TransientHResults { get; }

    /// <summary>Gets the explicit attempt cap.</summary>
    public int MaxAttempts { get; }

    /// <summary>Whether <paramref name="hresult"/> is a measured transient rejection.</summary>
    /// <param name="hresult">The COM failure code.</param>
    /// <returns><see langword="true"/> when the code is in the measured set.</returns>
    public bool IsTransient(int hresult) => TransientHResults.Contains(hresult);

    /// <summary>
    /// Renders the filter for the diagnostics record: the codes in ascending order, or
    /// <c>none</c> when the set is empty.
    /// </summary>
    /// <returns>The filter description.</returns>
    public string DescribeFilter() =>
        TransientHResults.Count == 0
            ? "none"
            : string.Join(
                ',',
                TransientHResults
                    .Order()
                    .Select(code => string.Create(CultureInfo.InvariantCulture, $"0x{code:X8}")));
}
