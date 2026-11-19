namespace GanttCreator.Office;

/// <summary>Why the verified copy stage produced no success (R6.6).</summary>
public enum StagedCopyRefusal
{
    /// <summary>No verified staged group outcome was supplied.</summary>
    NullOutcome = 0,

    /// <summary>No shape copy port was supplied.</summary>
    NullPort = 1,

    /// <summary>No clipboard format reader was supplied.</summary>
    NullReader = 2,

    /// <summary>The verified group carries no name to copy.</summary>
    BlankGroupName = 3,

    /// <summary>
    /// The copy port refused. <see cref="StagingCopyResult.PortRefusal"/> carries the port's
    /// reason, and nothing was placed on the clipboard by this operation.
    /// </summary>
    CopyRefused = 4,

    /// <summary>
    /// The clipboard was busy or unavailable after the copy. The detail carries the reader's
    /// diagnostic token and an actionable sentence (work-item D2).
    /// </summary>
    ClipboardUnavailable = 5,

    /// <summary>
    /// The clipboard does not carry the required Office drawing formats, so no consumer paste
    /// can materialise the group. The detail names the missing format and the observed set
    /// (work-item D1).
    /// </summary>
    ExpectedFormatAbsent = 6,
}

/// <summary>
/// R6.6: copies the verified staged group and reports success only after the clipboard carries
/// the expected Office drawing formats.
/// </summary>
/// <remarks>
/// <para>
/// <b>Success is a verified format, never a non-throwing call (D1).</b> The stage copies
/// through <see cref="IShapeCopyPort"/>, reads the clipboard through
/// <see cref="IClipboardFormatReader"/>, and requires both
/// <see cref="RequiredDrawingFormatNames"/> to be present before returning a successful
/// <see cref="StagingCopyResult"/>. The set is the R6.7 spike's measured evidence (ADR-0066
/// D2); it is never widened to "any format".
/// </para>
/// <para>
/// <b>A busy clipboard is a distinct typed failure (D2).</b> The reader returns a typed
/// failure rather than raising, and the stage maps it to
/// <see cref="StagedCopyRefusal.ClipboardUnavailable"/> with the token and an actionable
/// sentence - not an exception, not a silent pass.
/// </para>
/// <para>
/// <b>Immediate and bounded (D3).</b> One copy call, one read, one comparison - no polling,
/// no sleeps, no <c>Task.Run</c>. The lifecycle invokes this seam before its <c>finally</c>
/// cleanup, so the format is verified while the staging shapes still exist.
/// </para>
/// <para>
/// <b>No COM of its own.</b> Every host interaction goes through the injected ports, so this
/// stage is contract-tested without a host; the tagged live integration test drives the real
/// adapters.
/// </para>
/// </remarks>
public static class StagedShapeCopier
{
    /// <summary>
    /// The Office drawing clipboard formats a shape copy must leave for a consumer paste to
    /// materialise, measured on the R6.7 reference host (ADR-0066 D2). Both must be present.
    /// </summary>
    public static readonly string[] RequiredDrawingFormatNames =
    [
        "Excel 2007 Internal Shape",
        "Art::GVML ClipFormat",
    ];

    /// <summary>Copies the verified group and verifies the required clipboard formats.</summary>
    /// <param name="staged">The R6.4 verified group, or <see langword="null"/>.</param>
    /// <param name="port">The copy port, or <see langword="null"/>.</param>
    /// <param name="reader">The clipboard format reader, or <see langword="null"/>.</param>
    /// <returns>
    /// The enriched copy result. A successful result is returned only after the format
    /// verification passes.
    /// </returns>
    public static StagingCopyResult Copy(
        StagedGroupOutcome? staged,
        IShapeCopyPort? port,
        IClipboardFormatReader? reader)
    {
        if (staged is null)
        {
            return Refused(StagedCopyRefusal.NullOutcome, "No verified staged group was supplied.");
        }

        if (port is null)
        {
            return Refused(StagedCopyRefusal.NullPort, "No shape copy port was supplied.");
        }

        if (reader is null)
        {
            return Refused(StagedCopyRefusal.NullReader, "No clipboard format reader was supplied.");
        }

        if (string.IsNullOrWhiteSpace(staged.GroupName))
        {
            return Refused(StagedCopyRefusal.BlankGroupName, "The verified group carries no name to copy.");
        }

        string groupName = staged.GroupName;

        ShapeWriteOutcome copied = port.CopyStaged(groupName);
        if (!copied.Succeeded)
        {
            return new StagingCopyResult(
                false,
                $"The staged group could not be copied: {copied.Refusal}.",
                StagedCopyRefusal.CopyRefused,
                copied.Refusal);
        }

        ClipboardFormatRead read = reader.Read();
        if (!read.Succeeded || read.Snapshot is not { } snapshot)
        {
            return new StagingCopyResult(
                false,
                $"The clipboard was busy or unavailable ({read.Failure ?? "no snapshot"}). "
                    + "Close the application holding the clipboard and try again.",
                StagedCopyRefusal.ClipboardUnavailable,
                null);
        }

        foreach (string required in RequiredDrawingFormatNames)
        {
            if (!snapshot.Contains(required))
            {
                return new StagingCopyResult(
                    false,
                    $"The clipboard does not carry the expected drawing format '{required}'. "
                        + $"Observed formats: [{string.Join(", ", snapshot.FormatNames)}].",
                    StagedCopyRefusal.ExpectedFormatAbsent,
                    null);
            }
        }

        return new StagingCopyResult(
            true,
            "Copied the staged group and verified "
                + $"{string.Join(", ", RequiredDrawingFormatNames)} "
                + $"at clipboard sequence {snapshot.SequenceNumber}.",
            null,
            null);
    }

    private static StagingCopyResult Refused(StagedCopyRefusal refusal, string detail) =>
        new(false, detail, refusal, null);
}
