using GanttCreator.Core.Logging;

// CA2000: The fake log handed to AddInLifecycle is deliberately not
// disposed in the test body — AddInLifecycle is a writer, not the owner;
// AddInHost.AutoClose owns disposal and is covered by AddInHostTests.
// Scoped to this file; production code still enforces CA2000.
#pragma warning disable CA2000

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Record-shape contract tests for <see cref="AddInLifecycle"/>.
/// </summary>
public class AddInLifecycleTests
{
    [Fact]
    public void LogOpen_writes_exactly_one_record_with_the_agreed_shape()
    {
        var log = new CapturingLog();
        var lifecycle = new AddInLifecycle(log);

        lifecycle.LogOpen(new AddInIdentity("1.2.3", "16.0", "x64", "packed.xll"));

        Assert.Equal(
            ["open addin-version=1.2.3 excel-version=16.0 process=x64 xll=packed.xll"],
            log.Records);
    }

    [Fact]
    public void LogOpen_normalizes_missing_fields_to_unknown()
    {
        var log = new CapturingLog();
        var lifecycle = new AddInLifecycle(log);

        lifecycle.LogOpen(new AddInIdentity("", "unknown", " ", string.Empty));

        Assert.Equal(
            ["open addin-version=unknown excel-version=unknown process=unknown xll=unknown"],
            log.Records);
    }

    [Fact]
    public void LogClose_writes_exactly_one_close_record()
    {
        var log = new CapturingLog();
        var lifecycle = new AddInLifecycle(log);

        lifecycle.LogClose();

        Assert.Equal(["close"], log.Records);
    }

    [Fact]
    public void LogOpen_null_identity_throws_ArgumentNullException()
    {
        var lifecycle = new AddInLifecycle(new CapturingLog());
        Assert.Throws<ArgumentNullException>(() => lifecycle.LogOpen(null!));
    }

    [Fact]
    public void AddInLifecycle_null_log_throws_ArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>(() => new AddInLifecycle(null!));
    }
}
