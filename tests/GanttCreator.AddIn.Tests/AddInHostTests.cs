using System.Globalization;
using ExcelDna.Integration;
using GanttCreator.Core.Logging;

// CA2000: The fake logs are deliberately not disposed in the test body —
// the code under test (AddInHost.AutoClose) owns disposal and the tests
// assert that it disposes. The analyzer cannot see this ownership transfer
// through the injected delegates. Scoped to this file; production code
// still enforces CA2000.
#pragma warning disable CA2000

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Capturing fake for <see cref="IRollingLog"/>: records formatted messages
/// with the invariant culture so lifecycle contracts can be asserted
/// without a real file system.
/// </summary>
internal sealed class CapturingLog : IRollingLog
{
    public List<string> Records { get; } = [];

    public bool Disposed { get; private set; }

    public bool IsFailed => false;

    public void Write(string message) => Records.Add(message);

    public void Write(string format, params object?[] args) =>
        Records.Add(string.Format(CultureInfo.InvariantCulture, format, args));

    public void Dispose() => Disposed = true;
}

/// <summary>
/// Entry-point contract tests for <see cref="AddInHost"/>: the Excel-DNA
/// reflection contract, the one-record-per-call lifecycle contract, and the
/// never-throw guarantees around the guarded sources.
/// </summary>
public class AddInHostTests
{
    private static readonly AddInIdentity TestIdentity = new(
        "0.0.0+test",
        "16.0",
        "x64",
        "GanttCreator.AddIn-AddIn64-packed.xll");

    private const string ExpectedOpenRecord =
        "open addin-version=0.0.0+test excel-version=16.0 process=x64 xll=GanttCreator.AddIn-AddIn64-packed.xll";

    [Fact]
    public void AddInHost_is_public_implements_IExcelAddIn_and_has_parameterless_ctor()
    {
        Assert.True(typeof(AddInHost).IsPublic, "Excel-DNA requires a public add-in type.");
        Assert.True(
            typeof(IExcelAddIn).IsAssignableFrom(typeof(AddInHost)),
            "The entry point must implement ExcelDna.Integration.IExcelAddIn.");
        var ctor = typeof(AddInHost).GetConstructor(Type.EmptyTypes);
        Assert.NotNull(ctor);
        Assert.True(ctor!.IsPublic, "Excel-DNA requires a public parameterless constructor.");
    }

    [Fact]
    public void AutoOpen_writes_exactly_one_open_record_with_all_identifiers()
    {
        var log = new CapturingLog();
        var host = new AddInHost(() => TestIdentity, () => log);

        host.AutoOpen();

        Assert.Equal([ExpectedOpenRecord], log.Records);
        Assert.False(log.Disposed);
    }

    [Fact]
    public void AutoClose_writes_exactly_one_close_record_and_disposes_the_log()
    {
        var log = new CapturingLog();
        var host = new AddInHost(() => TestIdentity, () => log);
        host.AutoOpen();

        host.AutoClose();

        Assert.Equal([ExpectedOpenRecord, "close"], log.Records);
        Assert.True(log.Disposed, "AutoClose must dispose the log it opened.");
    }

    [Fact]
    public void AutoClose_without_AutoOpen_writes_nothing_and_does_not_throw()
    {
        var log = new CapturingLog();
        var host = new AddInHost(() => TestIdentity, () => log);

        host.AutoClose();

        Assert.Empty(log.Records);
        Assert.False(log.Disposed, "AutoClose must not dispose a log it never opened.");
    }

    [Fact]
    public void AutoOpen_log_source_failure_never_throws_and_writes_nothing()
    {
        var host = new AddInHost(
            () => TestIdentity,
            () => throw new InvalidOperationException("log source failed"));

        var exception = Record.Exception(host.AutoOpen);

        Assert.Null(exception);
    }

    [Fact]
    public void AutoOpen_identity_source_failure_never_throws_and_writes_nothing()
    {
        var log = new CapturingLog();
        var host = new AddInHost(
            () => throw new InvalidOperationException("identity source failed"),
            () => log);

        var exception = Record.Exception(host.AutoOpen);

        Assert.Null(exception);
        Assert.Empty(log.Records);
    }

    [Fact]
    public void AutoClose_log_write_failure_never_throws_and_still_disposes()
    {
        var log = new ThrowingLog();
        var host = new AddInHost(() => TestIdentity, () => log);
        host.AutoOpen();

        var exception = Record.Exception(host.AutoClose);

        Assert.Null(exception);
        Assert.True(log.Disposed, "AutoClose must dispose the log even when the write failed.");
    }

    [Fact]
    public void Second_AutoClose_is_a_no_op_and_never_throws()
    {
        var log = new CapturingLog();
        var host = new AddInHost(() => TestIdentity, () => log);
        host.AutoOpen();
        host.AutoClose();

        var exception = Record.Exception(host.AutoClose);

        Assert.Null(exception);
        Assert.Equal([ExpectedOpenRecord, "close"], log.Records);
        Assert.True(log.Disposed);
    }

    [Fact]
    public void AddInHost_null_identity_source_throws()
    {
        Assert.Throws<ArgumentNullException>(() => new AddInHost(null!, () => new CapturingLog()));
    }

    [Fact]
    public void AddInHost_null_log_source_throws()
    {
        Assert.Throws<ArgumentNullException>(() => new AddInHost(() => TestIdentity, null!));
    }

    /// <summary>
    /// A log whose writes throw — the never-throw guarantee of the entry
    /// point must hold even when the logging boundary misbehaves.
    /// </summary>
    private sealed class ThrowingLog : IRollingLog
    {
        public bool Disposed { get; private set; }

        public bool IsFailed => false;

        public void Write(string message) => throw new IOException("log write failed");

        public void Write(string format, params object?[] args) => throw new IOException("log write failed");

        public void Dispose() => Disposed = true;
    }
}
