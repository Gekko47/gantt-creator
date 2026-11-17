namespace GanttCreator.Office;

/// <summary>
/// The Gantt worksheet's protection state, exactly as the host reports it.
/// </summary>
/// <param name="ContentsProtected">
/// <c>Worksheet.ProtectContents</c>: locked cells are read-only.
/// </param>
/// <param name="ShapesProtected">
/// <c>Worksheet.ProtectDrawingObjects</c>: <c>Shape.Locked</c> is enforced. This is the flag
/// that makes shape locking real, and it is only true when <c>Protect</c> was called with
/// <c>DrawingObjects:=true</c> — the argument defaults to false, so a sheet can report itself
/// protected while every shape stays draggable.
/// </param>
/// <remarks>
/// A record rather than two separate members because the two flags are read together, from the
/// same worksheet, in the same pass; a caller that read only one of them would be able to
/// conclude "protected" while the shapes it wants to lock are not.
/// </remarks>
public sealed record WorksheetProtectionState(bool ContentsProtected, bool ShapesProtected);

/// <summary>Why a protection write could not run.</summary>
/// <remarks>
/// Typed and expected, per the error-handling convention: these are workbook conditions, not
/// faults. An unexpected COM failure is not a refusal and is reported as
/// <see cref="HostRejected"/>, because from the caller's point of view the host declined.
/// </remarks>
public enum ProtectionWriteRefusal
{
    /// <summary>The Excel application object or the active workbook is absent.</summary>
    NoActiveWorkbook = 0,

    /// <summary>The workbook holds no worksheet carrying the Gantt table.</summary>
    WorksheetMissing = 1,

    /// <summary>The host raised while applying or removing protection.</summary>
    HostRejected = 2,
}

/// <summary>The typed result of one protection write.</summary>
/// <param name="Refusal">The refusal reason, or <see langword="null"/> on success.</param>
public sealed record ProtectionWriteOutcome(ProtectionWriteRefusal? Refusal)
{
    /// <summary>Gets whether the write completed.</summary>
    public bool Succeeded => Refusal is null;

    /// <summary>Creates a successful outcome.</summary>
    /// <returns>The successful outcome.</returns>
    public static ProtectionWriteOutcome Ok() => new((ProtectionWriteRefusal?)null);

    /// <summary>Creates a refusal.</summary>
    /// <param name="refusal">The refusal reason.</param>
    /// <returns>The refusal outcome.</returns>
    public static ProtectionWriteOutcome Refused(ProtectionWriteRefusal refusal) => new(refusal);
}

/// <summary>
/// Applies and removes protection of the visible Gantt worksheet, and reads whether it is
/// protected (ADR-0055).
/// </summary>
/// <remarks>
/// <para>
/// <b>This is the one port that writes protection state, and it deliberately does not consult
/// <see cref="IWorksheetProtectionGuard"/> first.</b> The guard refuses a protected worksheet, and
/// re-applying <c>UserInterfaceOnly</c> after a reopen is only possible on a protected one, so a
/// guard-first rule here would make the guard's own recovery path unreachable. The architecture
/// gate for ADR-0008 D4 therefore treats this adapter as a declared exception rather than an
/// omission, and the exception is named in that test.
/// </para>
/// <para>
/// The contract is deliberately narrow: apply the three arguments ADR-0055 names
/// (<c>DrawingObjects</c>, <c>Contents</c>, <c>UserInterfaceOnly</c>) with no password, remove
/// protection, and read the two flags back. Policy — when to arm, when to release, and what the
/// <c>ProtectGanttSheet</c> setting means — belongs to the caller.
/// </para>
/// </remarks>
public interface IWorksheetProtectionPort
{
    /// <summary>Reads the Gantt worksheet's protection state.</summary>
    /// <returns>The state, or <see langword="null"/> when no Gantt worksheet can be resolved.</returns>
    WorksheetProtectionState? Read();

    /// <summary>
    /// Protects the Gantt worksheet with shape protection, content protection and
    /// interface-only restriction, and no password.
    /// </summary>
    /// <returns>The typed result.</returns>
    /// <remarks>
    /// The host raises when protection is already applied, so a caller that wants to re-apply
    /// <c>UserInterfaceOnly</c> must <see cref="Unprotect"/> first. That policy is the caller's
    /// precisely because "already protected" is not an error and this method must not guess.
    /// </remarks>
    ProtectionWriteOutcome Protect();

    /// <summary>Removes protection from the Gantt worksheet.</summary>
    /// <returns>The typed result.</returns>
    ProtectionWriteOutcome Unprotect();
}
