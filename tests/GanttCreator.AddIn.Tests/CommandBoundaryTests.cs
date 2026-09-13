using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using GanttCreator.Core.Logging;
using Moq;

// CA2201: tests construct COMException to exercise the COM-failure catch
// path in CommandErrorTranslator, which cannot be driven with any other
// exception type. Scoped to this file with a written justification.
#pragma warning disable CA2201

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Contract tests for the command error boundary (R1.4): the success path,
/// the one-record/one-dialog failure path, deterministic operation IDs,
/// HRESULT capture, the failure-proof guards, and user-safe translation.
/// None of these require Excel.
/// </summary>
public class CommandBoundaryTests
{
    private const string CommandName = "btnTest";

    /// <summary>The fixed instant used by the deterministic-clock tests.</summary>
    private static readonly DateTimeOffset FixedNow = new(2026, 9, 13, 10, 0, 0, 123, TimeSpan.Zero);

    /// <summary>Clock boundary frozen at a fixed instant for deterministic operation IDs.</summary>
    private sealed class FixedTimeProvider(DateTimeOffset now) : TimeProvider
    {
        public override DateTimeOffset GetUtcNow() => now;
    }

    /// <summary>
    /// Creates a boundary with a fake log whose formatted writes are captured
    /// into <paramref name="written"/>.
    /// </summary>
    private static CommandBoundary CreateBoundary(
        List<string> written,
        TimeProvider? clock = null,
        Action<string>? presenter = null)
    {
        written.Clear();
        var log = new Mock<IRollingLog>();
        log.Setup(l => l.IsFailed).Returns(false);
        log
            .Setup(l => l.Write(It.IsAny<string>(), It.IsAny<object?[]>()))
            .Callback<string, object?[]>(
                (fmt, args) => written.Add(string.Format(CultureInfo.InvariantCulture, fmt, args)));
        var boundary = new CommandBoundary(clock, presenter);
        boundary.SetLog(log.Object);
        return boundary;
    }

    /// <summary>Extracts the operation ID from one <c>CommandError</c> record.</summary>
    private static string ExtractOperationId(string record)
    {
        const string Marker = "op=";
        var start = record.IndexOf(Marker, StringComparison.Ordinal) + Marker.Length;
        var end = record.IndexOf(' ', start);
        return record[start..end];
    }

    [Fact]
    public void Run_executes_command_and_shows_nothing_on_success()
    {
        var written = new List<string>();
        var shown = new List<string>();
        var boundary = CreateBoundary(written, presenter: shown.Add);
        var executed = false;

        boundary.Run(CommandName, () => executed = true);

        Assert.True(executed, "The command must run.");
        Assert.Empty(written);
        Assert.Empty(shown);
    }

    [Fact]
    public void Run_thrown_fake_command_writes_exactly_one_log_record()
    {
        var written = new List<string>();
        var boundary = CreateBoundary(written);

        boundary.Run(CommandName, () => throw new InvalidOperationException("simulated"));

        var record = Assert.Single(written);
        Assert.Contains("CommandError", record, StringComparison.Ordinal);
        Assert.Contains("op=GC-", record, StringComparison.Ordinal);
        Assert.Contains($"command={CommandName}", record, StringComparison.Ordinal);
        Assert.Contains("stage=ribbon-callback", record, StringComparison.Ordinal);
        Assert.Contains("InvalidOperationException", record, StringComparison.Ordinal);
    }

    [Fact]
    public void Run_translates_exception_to_user_safe_message_containing_operation_id()
    {
        var written = new List<string>();
        var shown = new List<string>();
        var boundary = CreateBoundary(written, presenter: shown.Add);

        boundary.Run(CommandName, () => throw new InvalidOperationException("SECRET-DETAIL-42"));

        var record = Assert.Single(written);
        var message = Assert.Single(shown);
        Assert.Contains(ExtractOperationId(record), message, StringComparison.Ordinal);
        Assert.Contains("log", message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("SECRET-DETAIL-42", message, StringComparison.Ordinal);
        Assert.DoesNotContain("InvalidOperationException", message, StringComparison.Ordinal);
        Assert.DoesNotContain("at GanttCreator", message, StringComparison.Ordinal);
    }

    [Fact]
    public void Operation_id_is_deterministic_for_fixed_clock_and_increments_per_call()
    {
        var written = new List<string>();
        var boundary = CreateBoundary(written, clock: new FixedTimeProvider(FixedNow));

        boundary.Run(CommandName, () => throw new InvalidOperationException("a"));
        boundary.Run(CommandName, () => throw new InvalidOperationException("b"));

        var first = ExtractOperationId(written[0]);
        var second = ExtractOperationId(written[1]);
        const string Prefix = "GC-20260913-100000-123-";
        Assert.StartsWith(Prefix, first, StringComparison.Ordinal);
        Assert.StartsWith(Prefix, second, StringComparison.Ordinal);
        Assert.NotEqual(first, second);
        Assert.Equal("01", first[Prefix.Length..]);
        Assert.Equal("02", second[Prefix.Length..]);
    }

    [Fact]
    public void COMException_hresult_is_recorded_in_the_log()
    {
        var written = new List<string>();
        var boundary = CreateBoundary(written);

        boundary.Run(CommandName, () => throw new COMException("com rejected", unchecked((int)0x80004005)));

        Assert.Contains("hresult=0x80004005", Assert.Single(written), StringComparison.Ordinal);
    }

    [Fact]
    public void Boundary_never_throws_when_the_log_write_fails()
    {
        var log = new Mock<IRollingLog>();
        log
            .Setup(l => l.Write(It.IsAny<string>(), It.IsAny<object?[]>()))
            .Throws(new InvalidOperationException("log broken"));
        var shown = new List<string>();
        var boundary = new CommandBoundary(new FixedTimeProvider(FixedNow), shown.Add);
        boundary.SetLog(log.Object);

        boundary.Run(CommandName, () => throw new InvalidOperationException("simulated"));

        Assert.Single(shown);
    }

    [Fact]
    public void Boundary_never_throws_when_the_dialog_fails()
    {
        var written = new List<string>();
        var boundary = CreateBoundary(
            written,
            presenter: _ => throw new InvalidOperationException("dialog broken"));

        boundary.Run(CommandName, () => throw new InvalidOperationException("simulated"));

        Assert.Single(written);
    }

    [Theory]
    [InlineData(FailureCategory.Com, "Excel could not complete")]
    [InlineData(FailureCategory.Permission, "permission")]
    [InlineData(FailureCategory.FileUnavailable, "file")]
    [InlineData(FailureCategory.Cancelled, "cancelled")]
    [InlineData(FailureCategory.Generic, "unexpected error")]
    internal void Translate_maps_known_categories_and_falls_back_generic(
        FailureCategory category, string expectedFragment)
    {
        Exception exception = category switch
        {
            FailureCategory.Com => new COMException("SENTINEL-DETAIL", unchecked((int)0x80004005)),
            FailureCategory.Permission => new UnauthorizedAccessException("SENTINEL-DETAIL"),
            FailureCategory.FileUnavailable => new IOException("SENTINEL-DETAIL"),
            FailureCategory.Cancelled => new OperationCanceledException("SENTINEL-DETAIL"),
            _ => new InvalidOperationException("SENTINEL-DETAIL"),
        };

        var message = CommandErrorTranslator.Translate(exception, "GC-TEST-01");

        Assert.Contains(expectedFragment, message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("GC-TEST-01", message, StringComparison.Ordinal);
        Assert.DoesNotContain("SENTINEL-DETAIL", message, StringComparison.Ordinal);
    }

    /// <summary>The translated failure categories exercised by the translation theory.</summary>
    internal enum FailureCategory
    {
        Com,
        Permission,
        FileUnavailable,
        Cancelled,
        Generic,
    }

    [Fact]
    public void Translate_rejects_null_exception()
        => Assert.Throws<ArgumentNullException>(
            () => CommandErrorTranslator.Translate(null!, "GC-TEST-01"));

    [Fact]
    public void Translate_rejects_null_operation_id()
        => Assert.Throws<ArgumentNullException>(
            () => CommandErrorTranslator.Translate(new InvalidOperationException("x"), null!));

    [Fact]
    public void Run_rejects_null_command_name()
    {
        var written = new List<string>();
        var boundary = CreateBoundary(written);

        Assert.Throws<ArgumentNullException>(() => boundary.Run(null!, () => { }));
    }

    [Fact]
    public void Run_rejects_blank_command_name()
    {
        var written = new List<string>();
        var boundary = CreateBoundary(written);

        Assert.Throws<ArgumentException>(() => boundary.Run("   ", () => { }));
    }

    [Fact]
    public void Run_rejects_null_command()
    {
        var written = new List<string>();
        var boundary = CreateBoundary(written);

        Assert.Throws<ArgumentNullException>(() => boundary.Run(CommandName, null!));
    }

    [Fact]
    public void SetLog_rejects_null()
    {
        var boundary = new CommandBoundary();

        Assert.Throws<ArgumentNullException>(() => boundary.SetLog(null!));
    }
}
