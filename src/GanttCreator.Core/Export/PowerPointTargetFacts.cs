namespace GanttCreator.Core.Export;

/// <summary>
/// The primitive facts the PowerPoint target decision table consumes (R7.2).
/// </summary>
/// <remarks>
/// <para>
/// Every member is a primitive or a nullable primitive; no interop type appears, so the Core
/// table is testable without a host and the AddIn compilation never names the PowerPoint PIA.
/// </para>
/// <para>
/// The facts are assembled by the Office fact gatherer from the R7.1 application facts and the
/// R7.2 target observation. <paramref name="PresentationCount"/> is authoritative from the
/// target observation, not from the attach call: a presentation can be opened between the
/// attach call and the target read, and a stale count must never make the table choose
/// presentation creation over an open deck.
/// </para>
/// <para>
/// <paramref name="SlideShowActive"/> means a slide-show window exists — a show is running.
/// <paramref name="ActiveProtectedViewWindowPresent"/> means a protected-view window is the
/// on-top window, which is distinct from plain <paramref name="ProtectedViewWindowCount"/>
/// (a protected window can sit behind a normal deck).
/// </para>
/// </remarks>
/// <param name="AttachAvailable">Whether a typed PowerPoint application proxy was obtained.</param>
/// <param name="AttachHResult">The COM failure code when the attach was refused, or <see langword="null"/>.</param>
/// <param name="StartedByUs">Whether this add-in started the host process.</param>
/// <param name="ReadSucceeded">Whether the target read answered; a refused read is a typed fact.</param>
/// <param name="ReadHResult">The COM failure code when the target read was refused, or <see langword="null"/>.</param>
/// <param name="PresentationCount">The number of open presentations at the target read.</param>
/// <param name="ProtectedViewWindowCount">The number of open protected-view windows.</param>
/// <param name="ActiveProtectedViewWindowPresent">Whether a protected-view window is the on-top window.</param>
/// <param name="ActivePresentationReadOnly">The resolved active presentation's read-only state.</param>
/// <param name="ActiveWindowPresent">Whether the resolved active presentation has a document window.</param>
/// <param name="ActiveViewType">
/// The active view's <c>PpViewType</c> value, or <see langword="null"/> when there is no window.
/// The interpretation is the decision table's; Core never names the interop enum.
/// </param>
/// <param name="SlideShowActive">Whether a slide-show window exists (a show is running).</param>
/// <param name="SlideCount">The resolved active presentation's slide count.</param>
/// <param name="ActiveSlideIndex">
/// The active slide's one-based index, or <see langword="null"/> when the view exposes no slide.
/// </param>
public sealed record PowerPointTargetFacts(
    bool AttachAvailable,
    int? AttachHResult,
    bool StartedByUs,
    bool ReadSucceeded,
    int? ReadHResult,
    int PresentationCount,
    int ProtectedViewWindowCount,
    bool ActiveProtectedViewWindowPresent,
    bool ActivePresentationReadOnly,
    bool ActiveWindowPresent,
    int? ActiveViewType,
    bool SlideShowActive,
    int SlideCount,
    int? ActiveSlideIndex);
