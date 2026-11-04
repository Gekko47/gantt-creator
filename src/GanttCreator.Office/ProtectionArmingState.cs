namespace GanttCreator.Office;

/// <summary>
/// Records whether the add-in has already applied its own protection to the Gantt worksheet
/// in this Excel session (ADR-0055).
/// </summary>
/// <remarks>
/// <para>
/// <b>Why an object rather than a field on the orchestrator.</b> The live composition builds a
/// new <see cref="GanttRefreshOrchestrator"/> for every refresh, so a field would be
/// <see langword="false"/> on every pass and the arming would re-run each time — an
/// <c>Unprotect</c> plus a <c>Protect</c> per automatic refresh, which is exactly the extra step
/// the owner ruled out. This object outlives the orchestrator and carries the fact across passes.
/// </para>
/// <para>
/// <b>Why not "is this the same workbook" instead.</b> Excel exposes no property that says whether
/// <c>UserInterfaceOnly</c> is in force, and the worksheet proxy reached by one read is not
/// reference-equal to the one reached by the next, so there is no cheap identity to compare. The
/// arming is therefore a session fact that the add-in invalidates when the active workbook can have
/// changed, which is the only event that can make it wrong.
/// </para>
/// <para>
/// Mutable and explicitly shared, like <see cref="ProtectionOwnershipSource"/>: it is read and
/// written on the Excel main thread only, and the volatile field keeps a cross-thread read honest.
/// </para>
/// </remarks>
public sealed class ProtectionArmingState
{
    private volatile bool _armed;

    /// <summary>
    /// The Excel session's single arming memory, used by the live refresh composition.
    /// </summary>
    /// <remarks>
    /// The live composition builds a new orchestrator per refresh, so the memory has to be something
    /// that outlives it and something every refresh in the session shares. The add-in invalidates it
    /// when the active workbook changes, because the workbook the user switched to may be protected
    /// from an earlier session without <c>UserInterfaceOnly</c>.
    /// </remarks>
    public static ProtectionArmingState Shared { get; } = new();

    /// <summary>Gets whether the add-in has already applied its own protection this session.</summary>
    public bool IsArmed => _armed;

    /// <summary>Records that the add-in's protection is in force.</summary>
    /// <remarks>
    /// Called only after a successful <see cref="IWorksheetProtectionPort.Protect"/>. A refusal
    /// leaves the state unarmed so the next refresh tries again rather than assuming a lock that
    /// was never applied.
    /// </remarks>
    public void MarkArmed() => _armed = true;

    /// <summary>
    /// Records that the add-in's protection is no longer known to be in force.
    /// </summary>
    /// <remarks>
    /// Called when the active workbook changes (the new workbook may be protected from an earlier
    /// session without <c>UserInterfaceOnly</c>), when the user opts out of protection, and by the
    /// unlock command. The cost of being wrong is one re-arm on the next refresh, so this errs
    /// towards invalidating.
    /// </remarks>
    public void Invalidate() => _armed = false;
}
