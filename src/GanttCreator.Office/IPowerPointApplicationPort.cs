namespace GanttCreator.Office;

/// <summary>
/// What one attach-or-start observation reported about the desktop PowerPoint host (R7.1).
/// </summary>
/// <param name="Available">
/// Whether a typed PowerPoint application proxy was obtained. <see langword="false"/> is a
/// typed outcome, never an escaping exception.
/// </param>
/// <param name="StartedByUs">
/// Whether this call created the host process. A pre-existing PowerPoint is never owned.
/// </param>
/// <param name="ProcessId">The host process identifier, or zero when it could not be resolved.</param>
/// <param name="PresentationCount">The host's <c>Presentations.Count</c> at the observation.</param>
/// <param name="Build">The host's file version, or <see langword="null"/> when it is unreadable.</param>
/// <param name="HResult">The COM failure code when <paramref name="Available"/> is <see langword="false"/>.</param>
public sealed record PowerPointApplicationFacts(
    bool Available,
    bool StartedByUs,
    int ProcessId,
    int PresentationCount,
    string? Build,
    int? HResult);

/// <summary>
/// What one read of the active presentation, its window, its view and its slide reported
/// (R7.1). Every member is a fact; "proceed or refuse" is R7.2's Core decision table and is
/// deliberately not expressed here.
/// </summary>
/// <param name="ReadSucceeded">
/// Whether the host answered. A refused read is a typed fact so R7.2 can distinguish it from
/// a readable empty state.
/// </param>
/// <param name="HResult">The COM failure code when <paramref name="ReadSucceeded"/> is <see langword="false"/>.</param>
/// <param name="HasPresentation">Whether at least one presentation is open.</param>
/// <param name="ReadOnly">The active presentation's <c>ReadOnly</c> state, read without changing it.</param>
/// <param name="HasWindow">Whether the active presentation has a document window.</param>
/// <param name="SlideCount">The active presentation's slide count.</param>
/// <param name="ActiveViewType">
/// The active view's <c>PpViewType</c> value as an integer, or <see langword="null"/> when there
/// is no document window. The interpretation is R7.2's.
/// </param>
/// <param name="SlideShowActive">Whether a slide-show window is currently active.</param>
/// <param name="ActiveSlideIndex">
/// The active slide's one-based index, or <see langword="null"/> when the view exposes no slide.
/// </param>
public sealed record PowerPointPresentationFacts(
    bool ReadSucceeded,
    int? HResult,
    bool HasPresentation,
    bool ReadOnly,
    bool HasWindow,
    int SlideCount,
    int? ActiveViewType,
    bool SlideShowActive,
    int? ActiveSlideIndex);

/// <summary>
/// What one read of the paste target reported (R7.2): the presentation and protected-view
/// counts, and the resolved active presentation, window, view and slide facts. Every member is
/// a fact; "proceed or refuse" is R7.2's Core decision table and is deliberately not expressed
/// here.
/// </summary>
/// <param name="ReadSucceeded">
/// Whether the host answered. A refused read is a typed fact so the decision table can
/// distinguish it from a readable empty state.
/// </param>
/// <param name="HResult">The COM failure code when <paramref name="ReadSucceeded"/> is <see langword="false"/>.</param>
/// <param name="PresentationCount">The number of open presentations at the read.</param>
/// <param name="ProtectedViewWindowCount">The number of open protected-view windows.</param>
/// <param name="ActiveProtectedViewWindowPresent">Whether a protected-view window is the on-top window.</param>
/// <param name="ActivePresentationReadOnly">
/// The resolved active presentation's <c>ReadOnly</c> state, read without changing it.
/// </param>
/// <param name="ActiveWindowPresent">Whether the resolved active presentation has a document window.</param>
/// <param name="ActiveViewType">
/// The active view's <c>PpViewType</c> value as an integer, or <see langword="null"/> when there
/// is no document window. The interpretation is the decision table's.
/// </param>
/// <param name="SlideShowActive">Whether a slide-show window exists (a show is running).</param>
/// <param name="SlideCount">The resolved active presentation's slide count.</param>
/// <param name="ActiveSlideIndex">
/// The active slide's one-based index, or <see langword="null"/> when the view exposes no slide.
/// </param>
public sealed record PowerPointTargetObservation(
    bool ReadSucceeded,
    int? HResult,
    int PresentationCount,
    int ProtectedViewWindowCount,
    bool ActiveProtectedViewWindowPresent,
    bool ActivePresentationReadOnly,
    bool ActiveWindowPresent,
    int? ActiveViewType,
    bool SlideShowActive,
    int SlideCount,
    int? ActiveSlideIndex);

/// <summary>
/// What one create-presentation call reported (R7.5 PD-2). The D-R7.2 button contract creates a
/// presentation only when none is open, announced by the command; this observation reports
/// whether that creation happened and carries the blank slide the paste will target.
/// </summary>
/// <param name="Created">Whether a presentation and its blank slide were created.</param>
/// <param name="SlideIndex">The created blank slide's one-based index, or <see langword="null"/> on refusal.</param>
/// <param name="HResult">The COM failure code when the creation was refused, or <see langword="null"/>.</param>
public sealed record PowerPointPresentationCreation(bool Created, int? SlideIndex, int? HResult);

/// <summary>
/// The narrow PowerPoint application/presentation/slide fact port (R7.1).
/// </summary>
/// <remarks>
/// <para>
/// <b>Primitive surface.</b> No interop type appears in a signature, so the AddIn
/// compilation never names <c>Microsoft.Office.Interop.PowerPoint</c> — the CS0433
/// duplicate-type discipline documented on <see cref="IExcelApplicationAdapter"/>.
/// </para>
/// <para>
/// <b>STA only.</b> Every member requires the Office main STA thread; there is no async
/// wrapper and no <c>Task.Run</c> around COM (<c>AGENTS.md</c>, work item D3).
/// </para>
/// <para>
/// <b>Read, never activate.</b> The active presentation and slide are read through
/// <c>Application.ActiveWindow</c> and <c>View.Slide</c>; nothing is activated
/// (work item D4).
/// </para>
/// <para>
/// <b>The target read resolves the active deck, not the first one.</b>
/// <see cref="ReadTargetObservation"/> resolves the active presentation through
/// <c>Application.ActiveWindow.Presentation</c> when a document window exists, so with two or
/// more presentations open the facts describe the deck the user is working in. A windowless
/// deck falls back to <c>Presentations[1]</c> only to report its own facts, which the decision
/// table then refuses as <c>NoWindow</c>.
/// </para>
/// <para>
/// <b>The port cannot save.</b> No <c>Save</c>, <c>SaveAs</c>, <c>Close</c>, <c>Quit</c> or
/// <c>Activate</c> member exists on this surface or on any <c>IPowerPoint*</c> interface in
/// this assembly, which is proven structurally by an architecture test (work item D5).
/// </para>
/// <para>
/// <b>Ownership.</b> An implementation acquires every proxy it owns through
/// <see cref="PowerPointProxyOwnership"/> and releases it exactly once in reverse order with
/// <c>Marshal.ReleaseComObject</c>; it never terminates a pre-existing PowerPoint instance.
/// </para>
/// </remarks>
public interface IPowerPointApplicationPort
{
    /// <summary>
    /// Attaches to the running desktop PowerPoint, or starts it when none is running, and
    /// reports the application facts.
    /// </summary>
    /// <returns>The typed application facts; a refused COM call is reported, not thrown.</returns>
    /// <remarks>
    /// A typed <c>new Application()</c> both attaches to a running instance and starts a
    /// stopped one; the implementation records the <c>POWERPNT</c> process set before and
    /// after the call so <see cref="PowerPointApplicationFacts.StartedByUs"/> is measured,
    /// never assumed.
    /// </remarks>
    PowerPointApplicationFacts AttachOrStart();

    /// <summary>
    /// Reads the active presentation, window, view and slide facts without activating
    /// anything.
    /// </summary>
    /// <returns>The typed presentation facts; a refused read is reported, not thrown.</returns>
    PowerPointPresentationFacts ReadPresentationFacts();

    /// <summary>
    /// Reads the paste-target facts without activating anything: the presentation and
    /// protected-view counts, and the active presentation resolved through the active window
    /// (falling back to the first presentation only when the deck has no window).
    /// </summary>
    /// <returns>The typed target observation; a refused read is reported, not thrown.</returns>
    PowerPointTargetObservation ReadTargetObservation();

    /// <summary>
    /// Creates a presentation with one blank slide — the D-R7.2 branch taken only when no
    /// presentation is open. The caller announces the creation; the adapter reports whether it
    /// happened and never opens, closes or activates anything else.
    /// </summary>
    /// <returns>
    /// The typed creation facts; a refused creation is reported, not thrown, and the R7.5
    /// mapper maps it to <see cref="PowerPointTransferFailureKind.NoPresentation"/>.
    /// </returns>
    PowerPointPresentationCreation CreatePresentationWithBlankSlide();
}
