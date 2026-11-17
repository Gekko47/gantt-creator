namespace GanttCreator.Office;

/// <summary>
/// Supplies the fact that the add-in owns the protection currently applied to the visible Gantt
/// worksheet (ADR-0055).
/// </summary>
/// <remarks>
/// The guard cannot discover this for itself: Excel reports <c>ProtectContents</c> for any
/// protection and does not say who applied it, and it does not report
/// <c>UserInterfaceOnly</c> at all. Ownership is therefore configured, from the
/// <c>ProtectGanttSheet</c> setting the refresh has already read.
/// </remarks>
public interface IProtectionOwnershipSource
{
    /// <summary>
    /// Gets whether the protection on the Gantt worksheet is the add-in's own.
    /// </summary>
    bool AddInOwnsProtection { get; }
}

/// <summary>
/// The in-memory ownership flag the add-in sets when it reads the configuration.
/// </summary>
/// <remarks>
/// <para>
/// <b>This exists so the guard adds no host call to the refresh.</b> The refresh reads the
/// configuration once, at step 2, and then hands the answer here; the guard reads a field when it
/// is asked, so authorising its own protection costs nothing measurable. Making the guard read the
/// configuration itself would put a second config read inside every command and every refresh,
/// which is precisely the step ADR-0055 forbids.
/// </para>
/// <para>
/// It is <see langword="false"/> until something sets it, so every caller that constructs a guard
/// without an ownership source keeps today's behaviour exactly: a protected sheet refuses.
/// </para>
/// </remarks>
public sealed class ProtectionOwnershipSource : IProtectionOwnershipSource
{
    private volatile bool _addInOwnsProtection;

    /// <summary>
    /// The Excel session's single ownership flag, used by every guard constructed without an
    /// explicit source.
    /// </summary>
    /// <remarks>
    /// One add-in runs in one Excel session and protects at most one visible Gantt worksheet at a
    /// time, while roughly two dozen adapters each construct their own guard; threading this
    /// instance through every one of those call sites would be noise with no behavioural gain. A
    /// test that needs isolation passes its own instance, which the guard's constructor accepts.
    /// It starts <see langword="false"/>, so nothing changes for a caller that never sets it.
    /// </remarks>
    public static ProtectionOwnershipSource Shared { get; } = new();

    /// <inheritdoc />
    public bool AddInOwnsProtection => _addInOwnsProtection;

    /// <summary>Records whether the add-in currently owns the Gantt worksheet's protection.</summary>
    /// <param name="addInOwnsProtection">The ownership fact, from the configuration just read.</param>
    public void Set(bool addInOwnsProtection) => _addInOwnsProtection = addInOwnsProtection;
}
