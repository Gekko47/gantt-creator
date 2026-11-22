using System.Diagnostics;
using System.Globalization;

namespace GanttCreator.Office;

/// <summary>
/// One operation attempt's outcome for <see cref="PowerPointRetryLoop"/> (R7.5 D2/PD-3).
/// </summary>
/// <param name="Value">The operation result: a typed observation on success or refusal.</param>
/// <param name="FailureHResult">
/// The COM failure code when the attempt failed, or <see langword="null"/> when the attempt
/// produced its answer (including a typed absence with no COM failure).
/// </param>
/// <param name="MutationCommitted">
/// Whether the operation had already mutated the host when it failed. A mutation-bounded
/// operation (the paste) must not retry after this point, because re-invoking it would paste a
/// second group and orphan the first (PD-3).
/// </param>
/// <typeparam name="T">The operation's typed result.</typeparam>
public readonly record struct PowerPointRetryAttempt<T>(
    T Value,
    int? FailureHResult,
    bool MutationCommitted);

/// <summary>
/// R7.5 D2: the bounded retry loop, split by the operation's mutation boundary (PD-3).
/// </summary>
/// <remarks>
/// <para>
/// <b>Only measured transient rejections retry.</b> A failure whose HRESULT is not in the
/// policy's set is returned immediately — the non-transient positive test (D2).
/// </para>
/// <para>
/// <b>The mutation boundary.</b> A pure read never mutates, so it retries up to the cap. A
/// placement write is idempotent (four absolute members), so it retries post-mutation too. The
/// paste is not idempotent: a rejected <c>PasteSpecial</c> call never executed and is safe to
/// retry, but a failure after the paste committed is returned as-is, with the caller cleaning
/// the residue (PD-10).
/// </para>
/// <para>
/// <b>Diagnostics, never a loop condition.</b> The elapsed time appears only in the stop
/// record; the loop terminates on the attempt cap or a non-retryable result, never on a clock.
/// </para>
/// </remarks>
public static class PowerPointRetryLoop
{
    /// <summary>
    /// Runs <paramref name="attempt"/> until it produces a non-retryable result or the attempt
    /// cap is reached.
    /// </summary>
    /// <typeparam name="T">The operation's typed result.</typeparam>
    /// <param name="policy">The transient filter and attempt cap.</param>
    /// <param name="retryAfterMutation">
    /// Whether a failure that occurred after the host mutation may be retried (idempotent
    /// operations). The paste passes <see langword="false"/>; reads and the placement pass
    /// <see langword="true"/>.
    /// </param>
    /// <param name="attempt">One attempt, receiving the 1-based attempt number.</param>
    /// <param name="technicalRecord">Receives the retry records, or <see langword="null"/>.</param>
    /// <param name="stage">The stage name included in every retry record.</param>
    /// <returns>The final attempt's value.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="policy"/> or <paramref name="attempt"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException"><paramref name="stage"/> is blank.</exception>
    public static T Execute<T>(
        PowerPointTransientRetryPolicy policy,
        bool retryAfterMutation,
        Func<int, PowerPointRetryAttempt<T>> attempt,
        Action<string>? technicalRecord,
        string stage)
    {
        ArgumentNullException.ThrowIfNull(policy);
        ArgumentNullException.ThrowIfNull(attempt);
        ArgumentException.ThrowIfNullOrWhiteSpace(stage);

        Stopwatch stopwatch = Stopwatch.StartNew();
        for (int attemptNumber = 1; ; attemptNumber++)
        {
            PowerPointRetryAttempt<T> result = attempt(attemptNumber);
            if (result.FailureHResult is not int hresult)
            {
                return result.Value;
            }

            bool transient = policy.IsTransient(hresult);
            bool canRetry = transient
                && attemptNumber < policy.MaxAttempts
                && (retryAfterMutation || !result.MutationCommitted);
            if (canRetry)
            {
                technicalRecord?.Invoke(string.Create(
                    CultureInfo.InvariantCulture,
                    $"{stage} retry=continue attempt={attemptNumber} hresult=0x{hresult:X8}"));
                continue;
            }

            if (transient)
            {
                string head = string.Create(
                    CultureInfo.InvariantCulture,
                    $"{stage} retry=stop attempts={attemptNumber} maxattempts={policy.MaxAttempts}");
                string detail = string.Create(
                    CultureInfo.InvariantCulture,
                    $" filter={policy.DescribeFilter()} elapsedms={stopwatch.Elapsed.TotalMilliseconds:0.#} hresult=0x{hresult:X8}");
                technicalRecord?.Invoke(head + detail);
            }

            return result.Value;
        }
    }
}
