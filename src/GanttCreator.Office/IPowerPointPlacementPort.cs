using GanttCreator.Core.Export;

namespace GanttCreator.Office;

/// <summary>What one read of the active slide's size reported (R7.4).</summary>
/// <param name="ReadSucceeded">
/// Whether the host answered. <see langword="false"/> is a typed fact, never an escaping
/// exception.
/// </param>
/// <param name="HResult">The COM failure code when <paramref name="ReadSucceeded"/> is <see langword="false"/>.</param>
/// <param name="WidthPt">The slide width in points; zero when the read failed.</param>
/// <param name="HeightPt">The slide height in points; zero when the read failed.</param>
public sealed record PowerPointSlideSizeFacts(bool ReadSucceeded, int? HResult, double WidthPt, double HeightPt);

/// <summary>The typed result of one placement write attempt (R7.4).</summary>
public enum PowerPointPlacementApplyOutcome
{
    /// <summary>The four box members were written and read back without raising.</summary>
    Applied = 0,

    /// <summary>The application, its active window, view or slide was unavailable.</summary>
    NoSlide = 1,

    /// <summary>The forwarded group identifier does not name an <c>msoGroup</c> on the active slide.</summary>
    GroupMissing = 2,

    /// <summary>The host refused a write or a read with a COM failure.</summary>
    HostRejected = 3,
}

/// <summary>What one <see cref="IPowerPointPlacementPort.ApplyPlacement"/> call reported (R7.4).</summary>
/// <param name="Outcome">The typed write outcome.</param>
/// <param name="HResult">
/// The COM failure code when <paramref name="Outcome"/> is
/// <see cref="PowerPointPlacementApplyOutcome.HostRejected"/>.
/// </param>
/// <param name="LeftPt">The measured absolute left edge after the write, in points.</param>
/// <param name="TopPt">The measured absolute top edge after the write, in points.</param>
/// <param name="WidthPt">The measured width after the write, in points.</param>
/// <param name="HeightPt">The measured height after the write, in points.</param>
public sealed record PowerPointPlacementObservation(
    PowerPointPlacementApplyOutcome Outcome,
    int? HResult,
    double LeftPt,
    double TopPt,
    double WidthPt,
    double HeightPt);

/// <summary>
/// The narrow R7.4 placement port: one slide-size fact read and one single-group write whose
/// result is read back as primitive facts.
/// </summary>
/// <remarks>
/// <para>
/// <b>Primitive surface.</b> No interop type appears in a signature, so the AddIn compilation
/// never names <c>Microsoft.Office.Interop.PowerPoint</c> — the same duplicate-type discipline as
/// <see cref="IPowerPointApplicationPort"/> and <see cref="IPowerPointPastePort"/>.
/// </para>
/// <para>
/// <b>One shape by identity.</b> <see cref="ApplyPlacement"/> names exactly one shape by its
/// forwarded identifier and applies the placement to it. The port exposes no range write, so the
/// work item D4 rule — resize the <c>msoGroup</c>, never the returned range that also lists the
/// group's flattened children — is structural rather than a convention.
/// </para>
/// <para>
/// <b>STA only.</b> Every member requires the Office main STA thread; there is no async wrapper
/// and no <c>Task.Run</c> around COM.
/// </para>
/// <para>
/// <b>Read the active slide, never activate.</b> The slide is read through
/// <c>Application.ActiveWindow</c> / <c>View.Slide</c>; nothing is activated (work item D4).
/// </para>
/// <para>
/// <b>The port cannot save.</b> No <c>Save</c>, <c>SaveAs</c>, <c>Close</c>, <c>Quit</c> or
/// <c>Activate</c> member exists on this surface, which the reflective architecture scan proves
/// structurally (work item D5).
/// </para>
/// <para>
/// <b>Ownership.</b> An implementation acquires every proxy it owns through
/// <see cref="PowerPointProxyOwnership"/> and releases it exactly once in reverse order; it never
/// saves, closes or quits a presentation.
/// </para>
/// </remarks>
public interface IPowerPointPlacementPort
{
    /// <summary>
    /// Reads the active slide's size in points without activating anything.
    /// </summary>
    /// <returns>The typed size facts; a refused read is reported, not thrown.</returns>
    PowerPointSlideSizeFacts ReadSlideSize();

    /// <summary>
    /// Applies the single placement to the <c>msoGroup</c> named by <paramref name="groupId"/> on
    /// the active slide and reports the measured result.
    /// </summary>
    /// <param name="groupId">The pasted group's identifier, forwarded by the R7.3 paste stage.</param>
    /// <param name="placement">The single placement produced by <see cref="SlideFit.Fit"/>.</param>
    /// <returns>
    /// The typed observation; a refused COM call is reported, never thrown. The group is located
    /// by identifier and type, and only that shape is written.
    /// </returns>
    PowerPointPlacementObservation ApplyPlacement(int groupId, SlidePlacement placement);
}
