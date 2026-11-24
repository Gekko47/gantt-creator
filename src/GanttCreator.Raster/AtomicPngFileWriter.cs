using System.Globalization;

namespace GanttCreator.Raster;

/// <summary>
/// Writes encoded PNG bytes through a same-directory temporary file, verifies
/// the written bytes, and atomically replaces the target only after the
/// verification passes.
/// </summary>
/// <remarks>
/// <para>
/// The sequence is the work item's D1 decision: create an exclusively-named
/// temporary file beside the target, write and flush it, re-open and verify it
/// from disk, then atomically replace the target. Invalid bytes never reach the
/// target path, and every failure arm removes the temporary file.
/// </para>
/// <para>
/// Operational failures return typed <see cref="PngWriteOutcome"/> values; a
/// contract fault (null or empty arguments) throws. The command boundary owns
/// logging and user-message presentation.
/// </para>
/// </remarks>
public static class AtomicPngFileWriter
{
    private const int _sharingViolation = unchecked((int)0x80070020);
    private const int _lockViolation = unchecked((int)0x80070021);

    /// <summary>Writes <paramref name="png"/> to <paramref name="targetPath"/> atomically and verifies it.</summary>
    /// <param name="png">The encoded bytes and their declared dimensions.</param>
    /// <param name="targetPath">The user-chosen target path.</param>
    /// <param name="seam">The file-system seam; null selects the real operations.</param>
    /// <returns>A success outcome, or one typed refusal with message and record.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="png"/> or <paramref name="targetPath"/> is null.</exception>
    /// <exception cref="ArgumentException">Thrown when the path is empty or whitespace, or the bytes are empty.</exception>
    public static PngWriteOutcome Write(PngExportBytes png, string targetPath, PngWriteSeam? seam = null)
    {
        ArgumentNullException.ThrowIfNull(png);
        ArgumentNullException.ThrowIfNull(targetPath);
        if (string.IsNullOrWhiteSpace(targetPath))
        {
            throw new ArgumentException("The target path must not be empty.", nameof(targetPath));
        }

        if (png.Bytes.IsEmpty)
        {
            throw new ArgumentException("The encoded PNG bytes must not be empty.", nameof(png));
        }

        seam ??= new PngWriteSeam();

        string fullTarget;
        try
        {
            fullTarget = Path.GetFullPath(targetPath);
        }
        catch (ArgumentException ex)
        {
            return UnwritablePath(targetPath, ex);
        }
        catch (NotSupportedException ex)
        {
            return UnwritablePath(targetPath, ex);
        }
        catch (IOException ex)
        {
            return UnwritablePath(targetPath, ex);
        }

        string? directory = Path.GetDirectoryName(fullTarget);
        if (string.IsNullOrEmpty(directory))
        {
            return new PngWriteOutcome(
                false,
                PngWriteRefusal.UnwritablePath,
                "The PNG could not be written because the selected path has no directory.",
                Record("path", targetPath, "no-directory"));
        }

        string tempPath = Path.Combine(
            directory,
            Path.GetFileName(fullTarget) + "." + Path.GetRandomFileName() + ".tmp");
        bool tempCreated = false;
        try
        {
            using (Stream stream = seam.CreateTempStream(tempPath))
            {
                tempCreated = true;
                stream.Write(png.Bytes.Span);
                stream.Flush();
                if (stream is FileStream fileStream)
                {
                    fileStream.Flush(true);
                }
            }

            PngValidationResult validation = PngFileValidator.Validate(tempPath, png.Dimensions);
            if (!validation.IsValid)
            {
                return new PngWriteOutcome(
                    false,
                    PngWriteRefusal.ValidationFailed,
                    "The PNG was written but failed verification, so it was not saved. Try exporting again.",
                    Record("validate", targetPath, validation.Failure));
            }

            seam.ReplaceTarget(tempPath, fullTarget);
            return new PngWriteOutcome(
                true,
                null,
                $"Saved {fullTarget} ({png.Dimensions.PixelWidth} x {png.Dimensions.PixelHeight} px at 300 DPI).",
                string.Create(
                    CultureInfo.InvariantCulture,
                    $"png-write stage=complete path={fullTarget} width={png.Dimensions.PixelWidth} height={png.Dimensions.PixelHeight} bytes={png.Bytes.Length}"));
        }
        catch (UnauthorizedAccessException ex)
        {
            return UnwritablePath(targetPath, ex);
        }
        catch (DirectoryNotFoundException ex)
        {
            return UnwritablePath(targetPath, ex);
        }
        catch (PathTooLongException ex)
        {
            return UnwritablePath(targetPath, ex);
        }
        catch (NotSupportedException ex)
        {
            return UnwritablePath(targetPath, ex);
        }
        catch (ArgumentException ex)
        {
            return UnwritablePath(targetPath, ex);
        }
        catch (IOException ex) when (IsSharingViolation(ex))
        {
            return new PngWriteOutcome(
                false,
                PngWriteRefusal.LockedTarget,
                "The PNG could not be written because the file is open in another program. Close it and try again.",
                Record("replace", targetPath, ex));
        }
        catch (IOException ex)
        {
            return new PngWriteOutcome(
                false,
                PngWriteRefusal.WriteFailed,
                "The PNG could not be written because of a storage error.",
                Record("write", targetPath, ex));
        }
        finally
        {
            if (tempCreated)
            {
                File.Delete(tempPath);
            }
        }
    }

    private static bool IsSharingViolation(IOException exception) =>
        exception.HResult is _sharingViolation or _lockViolation;

    private static PngWriteOutcome UnwritablePath(string targetPath, Exception exception) =>
        new(
            false,
            PngWriteRefusal.UnwritablePath,
            "The PNG could not be written because the selected location is not writable.",
            Record("write", targetPath, exception));

    private static string Record(string stage, string targetPath, string detail) =>
        string.Create(
            CultureInfo.InvariantCulture,
            $"png-write stage={stage} path={targetPath} detail={detail}");

    private static string Record(string stage, string targetPath, Exception exception) =>
        string.Create(
            CultureInfo.InvariantCulture,
            $"png-write stage={stage} path={targetPath} exception={exception.GetType().Name} hresult=0x{exception.HResult:X8}");
}
