using GanttCreator.Core.Export;

namespace GanttCreator.Office;

/// <summary>
/// The Office fact gatherer for the R7.2 target decision: attaches to or starts PowerPoint
/// through the R7.1 port, reads the target observation, maps both to the Core facts, and calls
/// the Core table.
/// </summary>
/// <remarks>
/// <para>
/// <b>No COM of its own.</b> The resolver holds no proxies and performs no host call beyond
/// the two port members; every traversal and every release belongs to the port implementation.
/// </para>
/// <para>
/// <b>The observation's count is authoritative.</b> The attach call reports the presentation
/// count it saw, but a deck can be opened between the attach and the target read, so only the
/// observation's count reaches the table. A stale attach count must never turn an open deck
/// into a creation.
/// </para>
/// <para>
/// <b>Expected states are typed results.</b> A refused attach or a refused read maps to the
/// table's <c>AttachRefused</c> / <c>FactsUnreadable</c> outcomes; the resolver never throws
/// for an expected host state. It also never re-derives a case — the table is the one
/// authority (R7.2 D2).
/// </para>
/// </remarks>
public sealed class PowerPointTargetResolver
{
    private readonly IPowerPointApplicationPort _port;

    /// <summary>Creates a resolver over the PowerPoint fact port.</summary>
    /// <param name="port">The port the resolver drives.</param>
    /// <exception cref="ArgumentNullException"><paramref name="port"/> is <see langword="null"/>.</exception>
    public PowerPointTargetResolver(IPowerPointApplicationPort port)
    {
        ArgumentNullException.ThrowIfNull(port);
        _port = port;
    }

    /// <summary>Resolves the typed target decision for the current host state.</summary>
    /// <returns>The typed decision; expected host states are results, never exceptions.</returns>
    public PowerPointTargetDecision Resolve()
    {
        PowerPointApplicationFacts application = _port.AttachOrStart();
        PowerPointTargetObservation observation = _port.ReadTargetObservation();

        PowerPointTargetFacts facts = new(
            AttachAvailable: application.Available,
            AttachHResult: application.HResult,
            StartedByUs: application.StartedByUs,
            ReadSucceeded: observation.ReadSucceeded,
            ReadHResult: observation.HResult,
            PresentationCount: observation.PresentationCount,
            ProtectedViewWindowCount: observation.ProtectedViewWindowCount,
            ActiveProtectedViewWindowPresent: observation.ActiveProtectedViewWindowPresent,
            ActivePresentationReadOnly: observation.ActivePresentationReadOnly,
            ActiveWindowPresent: observation.ActiveWindowPresent,
            ActiveViewType: observation.ActiveViewType,
            SlideShowActive: observation.SlideShowActive,
            SlideCount: observation.SlideCount,
            ActiveSlideIndex: observation.ActiveSlideIndex);

        return PowerPointTargetDecision.Decide(facts);
    }
}
