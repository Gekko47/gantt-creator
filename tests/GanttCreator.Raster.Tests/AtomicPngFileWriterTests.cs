using GanttCreator.Core;

namespace GanttCreator.Raster.Tests;

/// <summary>
/// Acceptance pins for the R8.6 atomic writer: success paths create or replace
/// the target with verified bytes, and every injected failure leaves no partial
/// target, no temporary residue, and one typed refusal with a distinct message
/// and technical record.
/// </summary>
public class AtomicPngFileWriterTests
{
    private static readonly PixelDimensions _dimensions = new(60, 40);

    [Fact]
    public void A_successful_write_creates_the_target_with_the_exact_bytes()
    {
        using TestTempDirectory directory = new();
        PngExportBytes png = PngTestSupport.Encode(_dimensions);
        string target = directory.File("export.png");

        PngWriteOutcome outcome = AtomicPngFileWriter.Write(png, target);

        Assert.True(outcome.Succeeded);
        Assert.Null(outcome.Refusal);
        Assert.False(string.IsNullOrWhiteSpace(outcome.ActionableMessage));
        Assert.False(string.IsNullOrWhiteSpace(outcome.TechnicalRecord));
        Assert.Equal(png.Bytes.ToArray(), File.ReadAllBytes(target));
        Assert.True(PngFileValidator.Validate(target, _dimensions).IsValid);
        Assert.Empty(Directory.GetFiles(directory.Path, "*.tmp"));
    }

    [Fact]
    public void A_successful_write_replaces_an_existing_target_atomically()
    {
        using TestTempDirectory directory = new();
        PngExportBytes png = PngTestSupport.Encode(_dimensions);
        string target = directory.File("export.png");
        File.WriteAllBytes(target, [9, 9, 9, 9]);

        PngWriteOutcome outcome = AtomicPngFileWriter.Write(png, target);

        Assert.True(outcome.Succeeded);
        Assert.Equal(png.Bytes.ToArray(), File.ReadAllBytes(target));
        Assert.Single(Directory.GetFiles(directory.Path));
        Assert.Empty(Directory.GetFiles(directory.Path, "*.tmp"));
    }

    [Fact]
    public void A_mid_write_failure_leaves_no_target_and_cleans_the_temp()
    {
        using TestTempDirectory directory = new();
        PngExportBytes png = PngTestSupport.Encode(_dimensions);
        string target = directory.File("export.png");

        PngWriteOutcome outcome = AtomicPngFileWriter.Write(png, target, new FailingWriteSeam(10));

        Assert.False(outcome.Succeeded);
        Assert.Equal(PngWriteRefusal.WriteFailed, outcome.Refusal);
        Assert.Contains("storage error", outcome.ActionableMessage, StringComparison.Ordinal);
        Assert.False(string.IsNullOrWhiteSpace(outcome.TechnicalRecord));
        Assert.False(File.Exists(target));
        Assert.Empty(Directory.GetFiles(directory.Path, "*.tmp"));
    }

    [Fact]
    public void A_locked_target_is_refused_as_locked_and_left_unchanged()
    {
        using TestTempDirectory directory = new();
        PngExportBytes png = PngTestSupport.Encode(_dimensions);
        string target = directory.File("locked.png");
        byte[] original = [1, 2, 3, 4];
        File.WriteAllBytes(target, original);

        PngWriteOutcome outcome;
        using (FileStream hold = new(target, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
        {
            outcome = AtomicPngFileWriter.Write(png, target);
        }

        Assert.False(outcome.Succeeded);
        Assert.Equal(PngWriteRefusal.LockedTarget, outcome.Refusal);
        Assert.Contains("open in another program", outcome.ActionableMessage, StringComparison.Ordinal);
        Assert.False(string.IsNullOrWhiteSpace(outcome.TechnicalRecord));
        Assert.Equal(original, File.ReadAllBytes(target));
        Assert.Empty(Directory.GetFiles(directory.Path, "*.tmp"));
    }

    [Fact]
    public void A_read_only_target_is_refused_as_unwritable_and_left_unchanged()
    {
        using TestTempDirectory directory = new();
        PngExportBytes png = PngTestSupport.Encode(_dimensions);
        string target = directory.File("readonly.png");
        byte[] original = [5, 6, 7];
        File.WriteAllBytes(target, original);

        PngWriteOutcome outcome;
        File.SetAttributes(target, FileAttributes.ReadOnly);
        try
        {
            outcome = AtomicPngFileWriter.Write(png, target);
        }
        finally
        {
            File.SetAttributes(target, FileAttributes.Normal);
        }

        Assert.False(outcome.Succeeded);
        Assert.Equal(PngWriteRefusal.UnwritablePath, outcome.Refusal);
        Assert.Contains("not writable", outcome.ActionableMessage, StringComparison.Ordinal);
        Assert.False(string.IsNullOrWhiteSpace(outcome.TechnicalRecord));
        Assert.Equal(original, File.ReadAllBytes(target));
        Assert.Empty(Directory.GetFiles(directory.Path, "*.tmp"));
    }

    [Fact]
    public void A_parent_that_is_a_file_is_refused_as_unwritable()
    {
        using TestTempDirectory directory = new();
        PngExportBytes png = PngTestSupport.Encode(_dimensions);
        string blocker = directory.File("notadir");
        File.WriteAllBytes(blocker, [1]);
        string target = Path.Combine(blocker, "export.png");

        PngWriteOutcome outcome = AtomicPngFileWriter.Write(png, target);

        Assert.False(outcome.Succeeded);
        Assert.Equal(PngWriteRefusal.UnwritablePath, outcome.Refusal);
        Assert.Contains("not writable", outcome.ActionableMessage, StringComparison.Ordinal);
    }

    [Fact]
    public void An_invalid_path_is_refused_as_unwritable()
    {
        using TestTempDirectory directory = new();
        PngExportBytes png = PngTestSupport.Encode(_dimensions);

        PngWriteOutcome outcome = AtomicPngFileWriter.Write(png, "bad\0name.png");

        Assert.False(outcome.Succeeded);
        Assert.Equal(PngWriteRefusal.UnwritablePath, outcome.Refusal);
        Assert.False(string.IsNullOrWhiteSpace(outcome.TechnicalRecord));
        Assert.Empty(Directory.GetFiles(directory.Path));
    }

    [Fact]
    public void Validator_rejecting_bytes_leave_no_target_and_clean_the_temp()
    {
        using TestTempDirectory directory = new();
        PngExportBytes valid = PngTestSupport.Encode(_dimensions);
        PngExportBytes corrupt = new(
            PngTestSupport.SetDensity(valid.Bytes.ToArray(), 3780, 3780, 1),
            valid.Dimensions);
        string target = directory.File("corrupt.png");

        PngWriteOutcome outcome = AtomicPngFileWriter.Write(corrupt, target);

        Assert.False(outcome.Succeeded);
        Assert.Equal(PngWriteRefusal.ValidationFailed, outcome.Refusal);
        Assert.Contains("failed verification", outcome.ActionableMessage, StringComparison.Ordinal);
        Assert.False(string.IsNullOrWhiteSpace(outcome.TechnicalRecord));
        Assert.False(File.Exists(target));
        Assert.Empty(Directory.GetFiles(directory.Path, "*.tmp"));
    }

    [Fact]
    public void Contract_faults_are_rejected()
    {
        PngExportBytes png = PngTestSupport.Encode(_dimensions);
        PngExportBytes empty = new(ReadOnlyMemory<byte>.Empty, _dimensions);

        Assert.Throws<ArgumentNullException>(() => AtomicPngFileWriter.Write(null!, "export.png"));
        Assert.Throws<ArgumentNullException>(() => AtomicPngFileWriter.Write(png, null!));
        Assert.Throws<ArgumentException>(() => AtomicPngFileWriter.Write(png, "   "));
        Assert.Throws<ArgumentException>(() => AtomicPngFileWriter.Write(empty, "export.png"));
    }

    private sealed class FailingWriteSeam : PngWriteSeam
    {
        private readonly int _failAfterBytes;

        public FailingWriteSeam(int failAfterBytes) => _failAfterBytes = failAfterBytes;

        public override Stream CreateTempStream(string path) =>
            new FailingStream(base.CreateTempStream(path), _failAfterBytes);
    }

    private sealed class FailingStream : Stream
    {
        private readonly Stream _inner;
        private readonly int _failAfterBytes;
        private int _written;

        public FailingStream(Stream inner, int failAfterBytes)
        {
            _inner = inner;
            _failAfterBytes = failAfterBytes;
        }

        public override bool CanRead => false;

        public override bool CanSeek => false;

        public override bool CanWrite => true;

        public override long Length => _inner.Length;

        public override long Position
        {
            get => _inner.Position;
            set => throw new NotSupportedException();
        }

        public override void Flush() => _inner.Flush();

        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count)
        {
            ThrowWhenFull(count);
            _inner.Write(buffer, offset, count);
        }

        public override void Write(ReadOnlySpan<byte> buffer)
        {
            ThrowWhenFull(buffer.Length);
            _inner.Write(buffer);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _inner.Dispose();
            }

            base.Dispose(disposing);
        }

        private void ThrowWhenFull(int count)
        {
            if (_written + count > _failAfterBytes)
            {
                throw new IOException("The disk is full.");
            }

            _written += count;
        }
    }
}
