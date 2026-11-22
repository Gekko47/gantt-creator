namespace GanttCreator.Office;

/// <summary>
/// What one <c>Shapes.PasteSpecial(ppPasteShape)</c> call reported (R7.3).
/// </summary>
/// <remarks>
/// <para>
/// A refused paste is a typed fact, never an escaping exception. The observation carries
/// either the range facts (when <paramref name="PasteSucceeded"/> is <see langword="true"/>)
/// or the COM failure code the stage classifies before validation.
/// </para>
/// <para>
/// The group's identity (<paramref name="GroupChildCount"/>, <paramref name="GroupName"/>,
/// <paramref name="GroupId"/>) is an Office-side member: D2 forwards it to R7.4 so the
/// validated group is positioned without re-querying PowerPoint state the stage already read.
/// </para>
/// <para>
/// <paramref name="GroupShapeType"/> is the located group's <c>Shape.Type</c> value, or
/// <see langword="null"/> when the range holds no group. The group is located by examining
/// every item's type, never assumed to be index 1 (ADR-0066 D3).
/// </para>
/// </remarks>
/// <param name="PasteSucceeded">Whether the host returned a usable shape range.</param>
/// <param name="HResult">
/// The COM failure code when the host raised, or <see langword="null"/> when the paste target
/// (the application or its active slide) was unavailable without a COM failure.
/// </param>
/// <param name="Count">The returned <c>ShapeRange.Count</c>; zero when the paste did not succeed.</param>
/// <param name="GroupCount">How many <c>msoGroup</c> shapes the range contains; zero when the paste did not succeed.</param>
/// <param name="GroupShapeType">The located group's <c>Shape.Type</c> value, or <see langword="null"/>.</param>
/// <param name="GroupChildCount">The located group's <c>GroupItems.Count</c>; zero when no group was located.</param>
/// <param name="GroupName">The located group's <c>Name</c>, or <see langword="null"/>.</param>
/// <param name="GroupId">The located group's <c>Id</c>, or <see langword="null"/>.</param>
/// <param name="WidthPt">The returned range width in points; zero when the paste did not succeed.</param>
/// <param name="HeightPt">The returned range height in points; zero when the paste did not succeed.</param>
public sealed record PowerPointPasteObservation(
    bool PasteSucceeded,
    int? HResult,
    int Count,
    int GroupCount,
    int? GroupShapeType,
    int GroupChildCount,
    string? GroupName,
    int? GroupId,
    double WidthPt,
    double HeightPt);

/// <summary>The typed result of one paste-residue delete attempt (R7.5 PD-10).</summary>
public enum PowerPointGroupDeleteOutcome
{
    /// <summary>The named group was found and deleted; its children went with it.</summary>
    Deleted = 0,

    /// <summary>No shape with the forwarded identifier exists on the active slide.</summary>
    GroupMissing = 1,

    /// <summary>The host refused the delete with a COM failure; the HRESULT is carried.</summary>
    HostRejected = 2,
}

/// <summary>What one residue-delete attempt reported (R7.5 PD-10).</summary>
/// <param name="Outcome">The typed delete outcome.</param>
/// <param name="HResult">The COM failure code when <paramref name="Outcome"/> is <see cref="PowerPointGroupDeleteOutcome.HostRejected"/>.</param>
public sealed record PowerPointGroupDeleteResult(PowerPointGroupDeleteOutcome Outcome, int? HResult);

/// <summary>
/// The narrow PowerPoint paste port (R7.3): one <c>PasteSpecial(ppPasteShape)</c> call whose
/// result is read as primitive facts, plus the R7.5 residue delete for a paste that must be
/// undone on a failure path.
/// </summary>
/// <remarks>
/// <para>
/// <b>A separate port, deliberately.</b> This is a mutation, so it does not belong on the
/// read-only PowerPoint fact port (<see cref="IPowerPointApplicationPort"/>); the reflective
/// architecture scan covers it automatically because its name starts with <c>IPowerPoint</c>.
/// </para>
/// <para>
/// <b>Primitive surface.</b> No interop type appears in a signature, so the AddIn compilation
/// never names <c>Microsoft.Office.Interop.PowerPoint</c> — the same duplicate-type
/// discipline as <see cref="IPowerPointApplicationPort"/>.
/// </para>
/// <para>
/// <b>STA only.</b> The member requires the Office main STA thread; there is no async wrapper
/// and no <c>Task.Run</c> around COM.
/// </para>
/// <para>
/// <b>Read the active slide, never activate.</b> The target slide is read through
/// <c>Application.ActiveWindow</c> / <c>View.Slide</c> at paste time; nothing is activated
/// (work item D4). A user switching slides between the R7.2 decision and the paste changes
/// the target, which is the no-activation contract's own behaviour.
/// </para>
/// <para>
/// <b>Ownership.</b> An implementation acquires every proxy it owns through
/// <see cref="PowerPointProxyOwnership"/> and releases it exactly once in reverse order; it
/// never saves, closes or quits a presentation.
/// </para>
/// </remarks>
public interface IPowerPointPastePort
{
    /// <summary>
    /// Calls <c>Shapes.PasteSpecial(ppPasteShape)</c> on the current active slide and reports
    /// the returned range as primitive facts.
    /// </summary>
    /// <returns>
    /// The typed observation; a refused COM call is reported, never thrown. The group is
    /// located by examining every item's type, because the measured result is the group plus
    /// its flattened children, not a single-item range (ADR-0066 D3).
    /// </returns>
    PowerPointPasteObservation PasteSpecialShape();

    /// <summary>
    /// Deletes the pasted group named by <paramref name="groupId"/> — the group-delete mechanic
    /// (ADR-0066 D3): delete the group, never the returned range that also lists its flattened
    /// children. This is the R7.5 paste-residue cleanup the R7.3 stage records for R7.5
    /// (PD-10); it is called only on a failure path after a paste materialised.
    /// </summary>
    /// <param name="groupId">The pasted group's identifier carried by the paste result.</param>
    /// <returns>
    /// The typed delete result; a missing group or a refused delete is reported, never thrown.
    /// A cleanup delete is attempted once and its refusal recorded — it is never retried.
    /// </returns>
    PowerPointGroupDeleteResult DeletePastedGroup(int groupId);
}
