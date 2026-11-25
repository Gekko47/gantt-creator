using System.ComponentModel;
using System.Diagnostics;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>One identity-checked sample of an owned process's measurable counters (R7.7 D1).</summary>
/// <param name="ProcessId">The sampled process identifier.</param>
/// <param name="Alive">Whether the process was readable at sample time.</param>
/// <param name="StartTimeUtcTicks">
/// The process start time in UTC ticks; zero when unreadable. The soak compares it across
/// samples so a recycled PID is caught instead of being counted as the same process.
/// </param>
/// <param name="HandleCount">The process handle count (the .NET wrapper over <c>GetProcessHandleCount</c>); zero when unreadable.</param>
/// <param name="WorkingSetBytes">The process working set in bytes; zero when unreadable.</param>
internal sealed record SoakProcessCounters(
    int ProcessId,
    bool Alive,
    long StartTimeUtcTicks,
    int HandleCount,
    long WorkingSetBytes);

/// <summary>
/// Samples process counters without touching COM — the R7.7 D1 handles/working-set
/// evidence. A process that exits or denies access is reported
/// <see cref="SoakProcessCounters.Alive"/> = <see langword="false"/> rather than throwing:
/// the soak's own teardown can legitimately remove a process between samples, and the
/// orphan verdict is a separate end-state measurement.
/// </summary>
internal static class SoakProcessSampler
{
    /// <summary>Samples one process by identifier.</summary>
    /// <param name="processId">The process to sample.</param>
    /// <returns>The counters, or an unalive sample when the process cannot be read.</returns>
    internal static SoakProcessCounters Sample(int processId)
    {
        try
        {
            using Process process = Process.GetProcessById(processId);
            return new SoakProcessCounters(
                processId,
                Alive: true,
                process.StartTime.ToUniversalTime().Ticks,
                process.HandleCount,
                process.WorkingSet64);
        }
        catch (ArgumentException)
        {
            // The process exited between the caller's owned-set read and this sample.
            return new SoakProcessCounters(processId, Alive: false, 0, 0, 0);
        }
        catch (InvalidOperationException)
        {
            // Same.
            return new SoakProcessCounters(processId, Alive: false, 0, 0, 0);
        }
        catch (Win32Exception)
        {
            // The process exists but its counters are not readable.
            return new SoakProcessCounters(processId, Alive: false, 0, 0, 0);
        }
    }
}

/// <summary>The verdict of one counter trend across the soak series (R7.7 D1).</summary>
/// <param name="WithinTolerance">Whether the measured drift stays inside the recorded tolerance.</param>
/// <param name="Baseline">
/// The start statistic: the series minimum for handles, the first-third median for the
/// working set.
/// </param>
/// <param name="End">
/// The end statistic: the series maximum for handles, the last-third median for the
/// working set.
/// </param>
/// <param name="Delta">The measured spread or drift.</param>
internal sealed record SoakTrendOutcome(bool WithinTolerance, long Baseline, long End, long Delta);

/// <summary>
/// The R7.7 D1 trend evaluators. Both fail closed on a series too short to prove a flat
/// trend — an unmeasured trend is not a flat one.
/// </summary>
internal static class SoakTrend
{
    /// <summary>
    /// How many leading samples of the attach arm are warm-up rather than measurement.
    /// </summary>
    /// <remarks>
    /// The attach arm reuses one host across iterations, and the first passes allocate the
    /// steady state: on 2026-11-25 the series read <c>1892, 1951, 1956</c> and then settled at
    /// 1920-1956, so a whole-series max − min reported a 64 spread against a tolerance of 50 and
    /// failed a suite whose steady region spans 36. Two samples is the smallest exclusion that
    /// removes the allocation step; a leak cannot hide in it, because a leak keeps climbing
    /// through the steady region that is still gated.
    /// </remarks>
    internal const int HandleWarmUpSamples = 2;

    /// <summary>
    /// Handle-count trend: flat when the steady region's max − min stays within
    /// <paramref name="tolerance"/>. The handle count is the gated counter because a real
    /// proxy or handle leak is monotonic in it.
    /// </summary>
    /// <param name="samples">The per-iteration handle counts for one owned process.</param>
    /// <param name="tolerance">The measured maximum acceptable spread.</param>
    /// <param name="warmUpSamples">Leading samples to exclude as host warm-up.</param>
    /// <returns>The verdict and the spread measured over the steady region.</returns>
    internal static SoakTrendOutcome Handles(
        IReadOnlyList<int> samples,
        int tolerance,
        int warmUpSamples = HandleWarmUpSamples)
    {
        if (samples.Count == 0 || warmUpSamples < 0)
        {
            return new SoakTrendOutcome(WithinTolerance: false, 0, 0, 0);
        }

        // Fail closed when the steady region is too short to prove anything: an unmeasured
        // trend is not a flat one, the same discipline as the empty series.
        int[] steady = [.. samples.Skip(warmUpSamples)];
        if (steady.Length < 3)
        {
            return new SoakTrendOutcome(WithinTolerance: false, 0, 0, 0);
        }

        int minimum = steady.Min();
        int maximum = steady.Max();
        return new SoakTrendOutcome(maximum - minimum <= tolerance, minimum, maximum, maximum - minimum);
    }

    /// <summary>
    /// Working-set trend: compares the median of the first third against the median of the
    /// last third, so a single OS trim or GC spike cannot trip it and a monotonic leak
    /// cannot hide under one. Recorded rather than gated on the handle count's strictness
    /// (R7.7 review F5).
    /// </summary>
    /// <param name="samples">The per-iteration working sets for one owned process.</param>
    /// <param name="toleranceBytes">The measured maximum acceptable drift.</param>
    /// <returns>The verdict and the measured drift.</returns>
    internal static SoakTrendOutcome WorkingSet(IReadOnlyList<long> samples, long toleranceBytes)
    {
        if (samples.Count < 3)
        {
            return new SoakTrendOutcome(WithinTolerance: false, 0, 0, 0);
        }

        int third = samples.Count / 3;
        long baseline = Median(samples, offset: 0, count: third);
        long end = Median(samples, offset: samples.Count - third, count: third);
        long delta = Math.Abs(end - baseline);
        return new SoakTrendOutcome(delta <= toleranceBytes, baseline, end, delta);
    }

    private static long Median(IReadOnlyList<long> samples, int offset, int count)
    {
        long[] window = [.. samples.Skip(offset).Take(count).OrderBy(value => value)];
        int middle = window.Length / 2;
        return window.Length % 2 == 1
            ? window[middle]
            : (window[middle - 1] + window[middle]) / 2;
    }
}
