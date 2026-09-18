using System.Globalization;
using ExcelDna.Integration;
using GanttCreator.Core.Logging;
using GanttCreator.Office;

// CA2000: The fake logs are deliberately not disposed in the test body —
// the code under test (AddInHost.AutoClose) owns disposal and the tests
// assert that it disposes. The analyzer cannot see this ownership transfer
// through the injected delegates. Scoped to this file; production code
// still enforces CA2000.
#pragma warning disable CA2000

namespace GanttCreator.AddIn.Tests;

// CA1515: xUnit requires collection definitions to be public.
#pragma warning disable CA1515
[CollectionDefinition("diagnostics-service")]
public class DiagnosticsServiceCollectionDefinition
{
}
#pragma warning restore CA1515

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

    public string? LogFilePath => Path.Combine(Path.GetTempPath(), "GanttCreatorTests", "capturing.log");

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
[Collection("diagnostics-service")]
public class AddInHostTests
{
    private static readonly AddInIdentity TestIdentity = new(
        "0.0.0+test",
        "16.0",
        "x64",
        "GanttCreator.AddIn-AddIn64-packed.xll",
        "s7-g0h1i2j3k4l5");

    private const string ExpectedOpenRecord =
        "open addin-version=0.0.0+test excel-version=16.0 process=x64 xll=GanttCreator.AddIn-AddIn64-packed.xll session=s7-g0h1i2j3k4l5";

    private const string ExpectedCloseRecord = "close session=s7-g0h1i2j3k4l5";

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

        Assert.Equal([ExpectedOpenRecord, ExpectedCloseRecord], log.Records);
        Assert.True(log.Disposed, "AutoClose must dispose the log it opened.");
    }

    [Fact]
    public void GenerateSessionToken_survives_the_log_redactor_unchanged()
    {
        // D5 correlation depends on the token reaching disk byte-identical:
        // Core's Redactor masks GUIDs and long hex runs, so assert every
        // generated token passes through Redact untouched.
        var redactor = new GanttCreator.Core.Logging.Redactor();
        for (int i = 0; i < 25; i++)
        {
            var token = AddInHost.GenerateSessionToken();
            Assert.Equal(token, redactor.Redact(token));
        }
    }

    [Fact]
    public void GenerateSessionToken_produces_distinct_nonempty_tokens()
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        for (int i = 0; i < 25; i++)
        {
            var token = AddInHost.GenerateSessionToken();
            Assert.False(string.IsNullOrWhiteSpace(token));
            Assert.True(seen.Add(token), "Session tokens must be distinct per load.");
        }
    }

    [Fact]
    public void AutoClose_carries_the_open_session_token_into_the_close_record()
    {
        // The open/close pair of one load must be correlatable: AutoClose
        // writes the session token captured from the identity AutoOpen logged.
        var log = new CapturingLog();
        var host = new AddInHost(() => TestIdentity, () => log);
        host.AutoOpen();

        host.AutoClose();

        Assert.Equal([ExpectedOpenRecord, ExpectedCloseRecord], log.Records);
    }

    [Fact]
    public void AutoOpen_arms_the_ribbon_state_service_without_changing_the_one_record_contract()
    {
        var log = new CapturingLog();
        var host = new AddInHost(() => TestIdentity, () => log);

        try
        {
            host.AutoOpen();

            var service = RibbonStateService.Instance;
            // Outside a live Excel host the workbook fact is not determinable,
            // so the Diagnostics gate keeps its initial false value; the
            // capturing log exposes a path, so the Open-log gate is enabled.
            Assert.False(service.GetEnabled(RibbonControlIds.Diagnostics));
            Assert.True(service.GetEnabled(RibbonControlIds.OpenLog));
            Assert.Null(service.GetRibbon());

            // The state service writes no log record (work item R1.5): the
            // one-record contract is untouched.
            Assert.Equal([ExpectedOpenRecord], log.Records);
        }
        finally
        {
            RibbonStateService.Reset();
        }
    }

    [Fact]
    public void AutoClose_resets_the_ribbon_state_service()
    {
        var log = new CapturingLog();
        var host = new AddInHost(() => TestIdentity, () => log);
        host.AutoOpen();

        host.AutoClose();

        // Reset restores the initial snapshot: the fresh session service knows
        // no facts and holds no ribbon handle, so nothing from the closed
        // session leaks into the next one.
        Assert.False(RibbonStateService.Instance.GetEnabled(RibbonControlIds.OpenLog));
        Assert.Null(RibbonStateService.Instance.GetRibbon());
    }

    [Fact]
    public void AutoClose_resets_diagnostics_singleton_after_disposing_the_log()
    {
        var log = new CapturingLog();
        var host = new AddInHost(() => TestIdentity, () => log);
        DiagnosticsService.Reset();

        try
        {
            host.AutoOpen();

            var service = DiagnosticsService.Instance;
            Assert.Same(service, DiagnosticsService.Instance);

            host.AutoClose();

            Assert.Null(service.LogFilePath);
            Assert.NotSame(service, DiagnosticsService.Instance);
            Assert.Null(DiagnosticsService.Instance.LogFilePath);
            Assert.True(log.Disposed);
        }
        finally
        {
            DiagnosticsService.Reset();
        }
    }

    [Fact]
    public void AutoOpen_wires_the_log_into_the_diagnostics_service()
    {
        var log = new CapturingLog();
        var host = new AddInHost(() => TestIdentity, () => log);
        DiagnosticsService.Reset();

        try
        {
            host.AutoOpen();

            Assert.Equal(log.LogFilePath, DiagnosticsService.Instance.LogFilePath);
        }
        finally
        {
            DiagnosticsService.Reset();
        }
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
        Assert.Equal([ExpectedOpenRecord, ExpectedCloseRecord], log.Records);
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
    /// A log whose writes throw after the first call — the never-throw
    /// guarantee of the entry point must hold even when the logging
    /// boundary misbehaves. The first write (AutoOpen) succeeds so the
    /// log is retained; subsequent writes (AutoClose) throw.
    /// </summary>
    private sealed class ThrowingLog : IRollingLog
    {
        public bool Disposed { get; private set; }

        public bool IsFailed => false;

        public string? LogFilePath => Path.Combine(Path.GetTempPath(), "GanttCreatorTests", "throwing.log");

        private int _writeCount;

        public void Write(string message)
        {
            if (++_writeCount > 1)
                throw new IOException("log write failed");
        }

        public void Write(string format, params object?[] args)
        {
            if (++_writeCount > 1)
                throw new IOException("log write failed");
        }

        public void Dispose() => Disposed = true;
    }

    /// <summary>
    /// A log whose <see cref="IRollingLog.Dispose"/> throws — verifies
    /// that AutoClose suppresses the disposal exception and always
    /// clears <see cref="AddInHost._log"/>.
    /// </summary>
    private sealed class ThrowingDisposeLog : IRollingLog
    {
        public bool Disposed { get; private set; }

        public bool IsFailed => false;

        public string? LogFilePath => Path.Combine(Path.GetTempPath(), "GanttCreatorTests", "throwing-dispose.log");

        public int DisposeCalls { get; private set; }

        public void Write(string message) { }

        public void Write(string format, params object?[] args) { }

        public void Dispose()
        {
            DisposeCalls++;
            throw new IOException("dispose failed");
        }
    }

    /// <summary>
    /// A log that appends one token per IRollingLog call to a shared event
    /// list so the R1.6 teardown-order contract can assert the exact
    /// sequence: the formatted write is recorded as its first token
    /// (<c>open</c>), the plain write as its message (<c>close</c>), and
    /// disposal as <c>dispose</c>.
    /// </summary>
    private sealed class SequencedLog(List<string> events) : IRollingLog
    {
        public bool Disposed { get; private set; }

        public bool IsFailed => false;

        public string? LogFilePath => null;

        public void Write(string message) => events.Add(message);

        public void Write(string format, params object?[] args) =>
            events.Add(format.Split(' ')[0]);

        public void Dispose()
        {
            Disposed = true;
            events.Add("dispose");
        }
    }

    /// <summary>
    /// An adapter that records its subscription lifecycle into a shared
    /// event list (<c>subscribe</c> on subscribe, <c>detach</c> on dispose)
    /// and counts handler invocations so the dead-subscription contract can
    /// be asserted.
    /// </summary>
    private sealed class RecordingAdapter : IExcelApplicationAdapter
    {
        private readonly List<string> _events;
        private Action? _handler;

        public RecordingAdapter(List<string> events) => _events = events;

        public int HandlerInvocations { get; private set; }

        public bool? HasActiveWorkbook() => null;

        public IDisposable SubscribeWorkbookStateChanged(Action handler)
        {
            ArgumentNullException.ThrowIfNull(handler);
            _events.Add("subscribe");
            _handler = () =>
            {
                HandlerInvocations++;
                handler();
            };
            return new Subscription(this);
        }

        /// <summary>Simulates one workbook-state event from Excel.</summary>
        internal void RaiseWorkbookStateChanged() => _handler?.Invoke();

        private sealed class Subscription(RecordingAdapter owner) : IDisposable
        {
            private bool _disposed;

            public void Dispose()
            {
                if (_disposed)
                {
                    return;
                }

                _disposed = true;
                owner._events.Add("detach");
                owner._handler = null;
            }
        }
    }

    private static T? ReadPrivateField<T>(object obj, string fieldName)
    {
        var field = typeof(AddInHost).GetField(fieldName,
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        return field is null ? default : (T?)field.GetValue(obj);
    }

    [Fact]
    public void AutoClose_dispose_failure_never_throws_and_clears_log()
    {
        var log = new ThrowingDisposeLog();
        var host = new AddInHost(() => TestIdentity, () => log);
        host.AutoOpen();

        var exception = Record.Exception(host.AutoClose);

        Assert.Null(exception);
        Assert.Null(ReadPrivateField<IRollingLog?>(host, "_log"));
        Assert.False(log.Disposed);
        Assert.Equal(1, log.DisposeCalls);
    }

    // ---- Work item R1.6: deterministic shutdown and owned-resource cleanup ----

    [Fact]
    public void AutoClose_detaches_the_workbook_state_subscription_before_writing_the_close_record()
    {
        var events = new List<string>();
        var log = new SequencedLog(events);
        var adapter = new RecordingAdapter(events);
        RibbonStateService.Reset();
        var host = new AddInHost(() => TestIdentity, () => log, () => adapter);

        try
        {
            host.AutoOpen();
            host.AutoClose();

            // The deterministic teardown order (work item R1.6 D1): the owned
            // COM event connection is detached before the close record is
            // written, and the log is disposed last.
            Assert.Equal(
                ["open", "subscribe", "detach", "close", "dispose"],
                events);
        }
        finally
        {
            RibbonStateService.Reset();
        }
    }

    [Fact]
    public void AutoClose_leaves_a_dead_subscription_that_delivers_no_events()
    {
        var events = new List<string>();
        var log = new SequencedLog(events);
        var adapter = new RecordingAdapter(events);
        RibbonStateService.Reset();
        var host = new AddInHost(() => TestIdentity, () => log, () => adapter);

        try
        {
            host.AutoOpen();
            host.AutoClose();

            adapter.RaiseWorkbookStateChanged();

            // After teardown the event connection is dead: no handler runs
            // and the raised event writes nothing anywhere.
            Assert.Equal(0, adapter.HandlerInvocations);
            Assert.Equal(
                ["open", "subscribe", "detach", "close", "dispose"],
                events);
        }
        finally
        {
            RibbonStateService.Reset();
        }
    }

    [Fact]
    public void AutoClose_drops_every_service_log_reference_before_disposing_the_log()
    {
        var log = new CapturingLog();
        DiagnosticsService.Reset();
        CommandBoundary.Reset();
        RibbonStateService.Reset();
        var host = new AddInHost(() => TestIdentity, () => log);

        try
        {
            host.AutoOpen();

            // Capture the boundary that is active during the session: after
            // AutoClose the session is reset, and accessing the singleton
            // again would fabricate a fresh, trivially log-free boundary.
            var boundary = CommandBoundary.Instance;
            Assert.NotNull(boundary.GetLog());
            Assert.Equal(log.LogFilePath, DiagnosticsService.Instance.LogFilePath);

            host.AutoClose();

            // Both singletons dropped their log reference before the log was
            // disposed (work item R1.6 D1): no service observes a disposed log.
            Assert.Null(boundary.GetLog());
            Assert.Null(DiagnosticsService.Instance.LogFilePath);
            Assert.True(log.Disposed);
        }
        finally
        {
            DiagnosticsService.Reset();
            CommandBoundary.Reset();
            RibbonStateService.Reset();
        }
    }

    [Fact]
    public void AutoClose_close_record_failure_does_not_skip_the_singleton_resets()
    {
        // Every teardown step is guarded separately (work item R1.6 D1): a
        // failing close-record write must not skip the log-reference drops,
        // so no session singleton keeps the log that the last step disposes.
        // The throwing log's first write (AutoOpen) succeeds, so the log is
        // retained and only the close record fails.
        var log = new ThrowingLog();
        DiagnosticsService.Reset();
        CommandBoundary.Reset();
        RibbonStateService.Reset();
        var host = new AddInHost(() => TestIdentity, () => log);

        try
        {
            host.AutoOpen();

            // Capture the boundary that is active during the session: a
            // post-reset CommandBoundary.Instance would be a fresh,
            // trivially log-free boundary.
            var boundary = CommandBoundary.Instance;
            Assert.NotNull(boundary.GetLog());

            host.AutoClose();

            Assert.Null(boundary.GetLog());
            Assert.Null(DiagnosticsService.Instance.LogFilePath);
            Assert.True(log.Disposed, "the step after the failed close record must still dispose the log.");
        }
        finally
        {
            DiagnosticsService.Reset();
            CommandBoundary.Reset();
            RibbonStateService.Reset();
        }
    }

    [Fact]
    public void AutoOpen_after_AutoClose_reopens_deterministically()
    {
        // Each load cycle must get its own log: AutoClose disposes the log it
        // opened, so the reloaded session must open a fresh one, not reuse
        // the disposed instance.
        var firstLog = new CapturingLog();
        var secondLog = new CapturingLog();
        var loadCycle = 0;
        DiagnosticsService.Reset();
        CommandBoundary.Reset();
        RibbonStateService.Reset();
        var host = new AddInHost(
            () => TestIdentity,
            () => ++loadCycle == 1 ? firstLog : secondLog);

        try
        {
            host.AutoOpen();
            host.AutoClose();

            Assert.Equal([ExpectedOpenRecord, ExpectedCloseRecord], firstLog.Records);
            Assert.True(firstLog.Disposed);

            // Excel-DNA can load/unload/reload an XLL in one Excel session
            // (work item R1.6 D2): the reload must re-arm everything and
            // produce exactly one open/close pair per load cycle.
            host.AutoOpen();
            Assert.False(secondLog.Disposed, "the second session must run on a fresh, active log");
            Assert.Equal([ExpectedOpenRecord], secondLog.Records);

            host.AutoClose();
            Assert.Equal([ExpectedOpenRecord, ExpectedCloseRecord], secondLog.Records);
            Assert.True(secondLog.Disposed);

            Assert.Null(CommandBoundary.Instance.GetLog());
            Assert.Null(DiagnosticsService.Instance.LogFilePath);
        }
        finally
        {
            DiagnosticsService.Reset();
            CommandBoundary.Reset();
            RibbonStateService.Reset();
        }
    }

    [Fact]
    public void AutoClose_releases_the_real_log_file_handle()
    {
        var dir = Directory.CreateTempSubdirectory();
        try
        {
            var log = new RollingLog(dir.FullName, "r16-handle-release");
            DiagnosticsService.Reset();
            CommandBoundary.Reset();
            RibbonStateService.Reset();
            var host = new AddInHost(() => TestIdentity, () => log);

            try
            {
                host.AutoOpen();
                host.AutoClose();

                // The Windows-observable proof that the owned file handle was
                // released: the active log file is deletable after AutoClose
                // (File.Delete throws IOException while any handle is open).
                var activeLogPath = Path.Combine(dir.FullName, "r16-handle-release.log");
                Assert.True(
                    File.Exists(activeLogPath),
                    "the open record must have created the active log file");
                File.Delete(activeLogPath);
                Assert.False(File.Exists(activeLogPath));
            }
            finally
            {
                DiagnosticsService.Reset();
                CommandBoundary.Reset();
                RibbonStateService.Reset();
            }
        }
        finally
        {
            try
            {
                dir.Delete(recursive: true);
            }
            catch (IOException)
            {
                // Intentionally empty: temp-directory cleanup is best-effort.
            }
        }
    }
}
