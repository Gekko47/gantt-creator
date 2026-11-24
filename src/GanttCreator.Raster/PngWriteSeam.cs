namespace GanttCreator.Raster;

/// <summary>
/// The file-system seam the atomic writer uses. The default instance performs
/// the real operations; tests subclass it to inject failures without a host.
/// </summary>
/// <remarks>
/// This mirrors the injected-collaborator pattern already used for the pinned
/// font source: a public, overridable seam rather than an ambient machine probe.
/// </remarks>
public class PngWriteSeam
{
    /// <summary>Creates the temporary file in the target directory with an exclusive create.</summary>
    /// <param name="path">The temporary file path.</param>
    /// <returns>The writable stream the writer owns and disposes.</returns>
    public virtual Stream CreateTempStream(string path)
    {
        ArgumentNullException.ThrowIfNull(path);
        return new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None);
    }

    /// <summary>Replaces the target with the flushed temporary file atomically.</summary>
    /// <param name="tempPath">The fully written and verified temporary file.</param>
    /// <param name="targetPath">The user-chosen target path.</param>
    public virtual void ReplaceTarget(string tempPath, string targetPath)
    {
        ArgumentNullException.ThrowIfNull(tempPath);
        ArgumentNullException.ThrowIfNull(targetPath);
        if (File.Exists(targetPath))
        {
            File.Replace(tempPath, targetPath, null);
        }
        else
        {
            File.Move(tempPath, targetPath);
        }
    }
}
