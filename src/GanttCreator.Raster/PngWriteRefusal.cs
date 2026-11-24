namespace GanttCreator.Raster;

/// <summary>The closed set of operational failures an atomic PNG write can report.</summary>
/// <remarks>
/// Each refusal maps to exactly one actionable user message and one
/// redaction-ready technical record; <see cref="AtomicPngFileWriter"/> never
/// reports success for any of them.
/// </remarks>
public enum PngWriteRefusal
{
    /// <summary>The target path is not writable (invalid path, missing parent, or access denied).</summary>
    UnwritablePath,

    /// <summary>The target file is open in another process, so the atomic replace was refused.</summary>
    LockedTarget,

    /// <summary>A storage-class write error occurred (for example, a full disk).</summary>
    WriteFailed,

    /// <summary>The written temporary file failed the post-write verification.</summary>
    ValidationFailed
}
