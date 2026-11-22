using System.Collections.Frozen;

namespace GanttCreator.Core.Export;

/// <summary>The action the PowerPoint target decision table resolves for one target state.</summary>
public enum PowerPointTargetAction
{
    /// <summary>Paste into the resolved active editable presentation and its active slide.</summary>
    ProceedWithActiveSlide = 0,

    /// <summary>No presentation is open; create one and announce it (D-R7.2).</summary>
    CreatePresentation = 1,

    /// <summary>Refuse with the typed reason on <see cref="PowerPointTargetDecision.Refusal"/>.</summary>
    Refuse = 2,
}

/// <summary>
/// The closed set of typed refusals an attach/start or active-slide resolution can produce
/// (R7.2 D1). Each member is one user-visible case the command translates to one actionable
/// message; the message text belongs to the command, never to this table.
/// </summary>
public enum PowerPointTargetRefusal
{
    /// <summary>The typed PowerPoint application proxy could not be obtained.</summary>
    AttachRefused = 0,

    /// <summary>The target read was refused by the host (for example a missing member).</summary>
    FactsUnreadable = 1,

    /// <summary>A slide-show window exists, so a shape paste is not a valid target.</summary>
    SlideShowActive = 2,

    /// <summary>
    /// The on-top surface is a protected-view window, or the only open document surface is a
    /// protected-view window.
    /// </summary>
    ProtectedView = 3,

    /// <summary>The resolved active presentation is read-only or password-protected.</summary>
    ReadOnlyPresentation = 4,

    /// <summary>A presentation is open but has no document window, so there is no active slide.</summary>
    NoWindow = 5,

    /// <summary>The active view is not slide, notes page or normal.</summary>
    UnsupportedView = 6,

    /// <summary>The resolved active presentation has no slides.</summary>
    NoSlide = 7,

    /// <summary>The view exposes no active slide.</summary>
    NoActiveSlide = 8,
}

/// <summary>
/// The typed outcome of the attach/start and active-slide decision matrix (R7.2 D1/D2).
/// </summary>
/// <param name="Action">The resolved action.</param>
/// <param name="Refusal">The typed refusal when <paramref name="Action"/> is <see cref="PowerPointTargetAction.Refuse"/>; otherwise <see langword="null"/>.</param>
/// <param name="StartedByUs">
/// Whether the add-in started the host process, echoed so the command can announce a start
/// (D3) from the decision alone.
/// </param>
/// <param name="ActiveSlideIndex">
/// The resolved active slide's one-based index when <paramref name="Action"/> is
/// <see cref="PowerPointTargetAction.ProceedWithActiveSlide"/>; otherwise <see langword="null"/>.
/// </param>
/// <param name="HResult">
/// The COM failure code, carried only for <see cref="PowerPointTargetRefusal.AttachRefused"/>
/// and <see cref="PowerPointTargetRefusal.FactsUnreadable"/>; otherwise <see langword="null"/>.
/// </param>
/// <remarks>
/// <para>
/// <b>One authority.</b> This static table is the single answer to "proceed or refuse" (D2).
/// The Office fact gatherer reads the host and calls it; no later row — including the
/// command — re-derives a case.
/// </para>
/// <para>
/// <b>Invariants, pinned by tests.</b> A refusal carries no slide index and only the two
/// unreadable refusals carry an <c>HResult</c>; a proceed carries neither a refusal nor an
/// <c>HResult</c>; a creation carries no slide index and no <c>HResult</c>.
/// </para>
/// <para>
/// <b>Precedence is total.</b> The order is: refused attach, refused read, slide show,
/// foreground protected view, protected view with no presentation, creation, read-only,
/// no window, unsupported or absent view, no slide, no active slide, proceed. Impossible
/// fact combinations resolve deterministically by that order.
/// </para>
/// </remarks>
public sealed record PowerPointTargetDecision(
    PowerPointTargetAction Action,
    PowerPointTargetRefusal? Refusal,
    bool StartedByUs,
    int? ActiveSlideIndex,
    int? HResult)
{
    /// <summary>
    /// The <c>PpViewType</c> values a shape paste may target: slide (1), notes page (3) and
    /// normal (9). Every other value — including a future or unknown value — is refused.
    /// </summary>
    /// <remarks>
    /// Core cannot name the interop enum, so the values are documented constants pinned by
    /// the Core tests against the documented <c>PpViewType</c> enumeration.
    /// </remarks>
    public static IReadOnlySet<int> PasteableViewTypes { get; } = FrozenSet.ToFrozenSet([1, 3, 9]);

    /// <summary>
    /// Resolves the typed outcome for one set of host facts.
    /// </summary>
    /// <param name="facts">The primitive facts the Office fact gatherer observed.</param>
    /// <returns>The typed decision; every case maps to exactly one outcome.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="facts"/> is <see langword="null"/>.</exception>
    public static PowerPointTargetDecision Decide(PowerPointTargetFacts facts)
    {
        ArgumentNullException.ThrowIfNull(facts);

        return facts switch
        {
            { AttachAvailable: false } => Refuse(
                PowerPointTargetRefusal.AttachRefused,
                facts,
                facts.AttachHResult),
            { ReadSucceeded: false } => Refuse(
                PowerPointTargetRefusal.FactsUnreadable,
                facts,
                facts.ReadHResult),
            { SlideShowActive: true } => Refuse(PowerPointTargetRefusal.SlideShowActive, facts),
            { ActiveProtectedViewWindowPresent: true } => Refuse(
                PowerPointTargetRefusal.ProtectedView,
                facts),
            { PresentationCount: 0, ProtectedViewWindowCount: > 0 } => Refuse(
                PowerPointTargetRefusal.ProtectedView,
                facts),
            { PresentationCount: 0 } => new PowerPointTargetDecision(
                PowerPointTargetAction.CreatePresentation,
                null,
                facts.StartedByUs,
                null,
                null),
            { ActivePresentationReadOnly: true } => Refuse(
                PowerPointTargetRefusal.ReadOnlyPresentation,
                facts),
            { ActiveWindowPresent: false } => Refuse(PowerPointTargetRefusal.NoWindow, facts),
            { ActiveViewType: null } => Refuse(PowerPointTargetRefusal.UnsupportedView, facts),
            { ActiveViewType: not (1 or 3 or 9) } => Refuse(
                PowerPointTargetRefusal.UnsupportedView,
                facts),
            { SlideCount: 0 } => Refuse(PowerPointTargetRefusal.NoSlide, facts),
            { ActiveSlideIndex: null } => Refuse(PowerPointTargetRefusal.NoActiveSlide, facts),
            _ => new PowerPointTargetDecision(
                PowerPointTargetAction.ProceedWithActiveSlide,
                null,
                facts.StartedByUs,
                facts.ActiveSlideIndex,
                null),
        };
    }

    private static PowerPointTargetDecision Refuse(
        PowerPointTargetRefusal refusal,
        PowerPointTargetFacts facts,
        int? hresult = null) =>
        new(PowerPointTargetAction.Refuse, refusal, facts.StartedByUs, null, hresult);
}
