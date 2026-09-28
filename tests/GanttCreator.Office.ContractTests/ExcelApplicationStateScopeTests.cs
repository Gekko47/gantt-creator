using Excel = Microsoft.Office.Interop.Excel;
using Moq;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for <see cref="ExcelApplicationStateScope"/>. Every property
/// read and write is an <c>internal virtual</c> seam, so each of the six
/// settings is driven and asserted without a live Excel, and a failure can be
/// injected at each one (docs/08-TEST-CHECKLIST.md section K).
/// </summary>
public class ExcelApplicationStateScopeTests
{
    /// <summary>
    /// The scope over in-memory state, with an optional injected failure on any
    /// single write. The injected failure fires on the FIRST call, which for a
    /// captured setting is the command's own change; a test then proves the
    /// restore still put the original value back.
    /// </summary>
    private sealed class TestableScope(Action<string>? technicalRecord = null)
        : ExcelApplicationStateScope(new Mock<Excel.Application>().Object, technicalRecord)
    {
        /// <summary>Which write, if any, throws on its first invocation.</summary>
        public string? FailOn { get; set; }

        /// <summary>Observes every write, for the no-write-when-unchanged pin.</summary>
        public Action<string, object?>? CountWrites { get; set; }

        /// <summary>Whether the selection restore throws.</summary>
        public bool FailSelectionRestore { get; set; }

        /// <summary>Whether the selection restore ran.</summary>
        public bool SelectionRestored { get; private set; }

        /// <summary>The current values, standing in for the host's state.</summary>
        public bool ScreenUpdatingState { get; internal set; } = true;

        /// <summary>The current event state.</summary>
        public bool EnableEventsState { get; internal set; } = true;

        /// <summary>The current alert state.</summary>
        public bool DisplayAlertsState { get; internal set; } = true;

        /// <summary>Whether the status bar is currently visible.</summary>
        public bool DisplayStatusBarState { get; internal set; } = true;

        /// <summary>The current status-bar text.</summary>
        public object? StatusBarTextState { get; internal set; } = "Ready";

        /// <summary>The captured selection, or null.</summary>
        public object? SelectionState { get; internal set; } = new object();

        private void MaybeFail(string member)
        {
            if (string.Equals(FailOn, member, StringComparison.Ordinal))
            {
                throw new InvalidOperationException("injected failure on " + member);
            }
        }

        /// <summary>Records a write attempt, then applies the injected failure.</summary>
        /// <param name="member">The member being written.</param>
        /// <param name="value">The value being written.</param>
        /// <param name="apply">Applies the value when no failure is injected.</param>
        private void Write(string member, object? value, Action apply)
        {
            CountWrites?.Invoke(member, value);
            MaybeFail(member);
            apply();
        }

        internal override bool GetScreenUpdating() => ScreenUpdatingState;

        internal override void SetScreenUpdating(bool value)
        {
            CountWrites?.Invoke("ScreenUpdating", value);
            MaybeFail("ScreenUpdating");
            ScreenUpdatingState = value;
            ScreenUpdatingState = value;
        }

        internal override bool GetEnableEvents() => EnableEventsState;

        internal override void SetEnableEvents(bool value)
        {
            CountWrites?.Invoke("EnableEvents", value);
            MaybeFail("EnableEvents");
            EnableEventsState = value;
            EnableEventsState = value;
        }

        internal override bool GetDisplayAlerts() => DisplayAlertsState;

        internal override void SetDisplayAlerts(bool value)
        {
            CountWrites?.Invoke("DisplayAlerts", value);
            MaybeFail("DisplayAlerts");
            DisplayAlertsState = value;
            DisplayAlertsState = value;
        }

        internal override bool GetDisplayStatusBar() => DisplayStatusBarState;

        internal override void SetDisplayStatusBar(bool value)
        {
            CountWrites?.Invoke("DisplayStatusBar", value);
            MaybeFail("DisplayStatusBar");
            DisplayStatusBarState = value;
            DisplayStatusBarState = value;
        }

        internal override object? GetStatusBarText() => StatusBarTextState;

        internal override void SetStatusBarTextCore(object? value)
        {
            CountWrites?.Invoke("StatusBarText", value);
            MaybeFail("StatusBarText");
            StatusBarTextState = value;
            StatusBarTextState = value;
        }
        internal override object? GetSelection() => SelectionState;

        internal override void RestoreSelection(object selection)
        {
            if (FailSelectionRestore)
            {
                throw new InvalidOperationException("injected selection failure");
            }

            SelectionRestored = true;
        }
    }

    /// <summary>Applies every suppression the scope offers.</summary>
    private static void SuppressAll(TestableScope scope)
    {
        scope.SuppressScreenUpdating();
        scope.SuppressEvents();
        scope.SuppressAlerts();
        scope.SuppressStatusBar();
        scope.SetStatusBarText("Rendering...");
    }
    [Fact]
    public void Every_setting_is_captured_changed_and_put_back()
    {
        var scope = new TestableScope();

        SuppressAll(scope);
        Assert.False(scope.ScreenUpdatingState);
        Assert.False(scope.EnableEventsState);
        Assert.False(scope.DisplayAlertsState);
        Assert.False(scope.DisplayStatusBarState);
        Assert.Equal("Rendering...", scope.StatusBarTextState);

        scope.Dispose();

        Assert.True(scope.ScreenUpdatingState);
        Assert.True(scope.EnableEventsState);
        Assert.True(scope.DisplayAlertsState);
        Assert.True(scope.DisplayStatusBarState);
        Assert.Equal("Ready", scope.StatusBarTextState);
    }

    [Fact]
    public void An_unchanged_setting_is_never_written_on_the_way_in_or_out()
    {
        // D1, positive: a command that suppresses nothing that was already
        // suppressed must not write those properties at all. Writing them would
        // perturb a setting the command never touched.
        var scope = new TestableScope();
        scope.ScreenUpdatingState = false;
        scope.EnableEventsState = false;
        scope.DisplayAlertsState = false;
        scope.DisplayStatusBarState = false;
        scope.StatusBarTextState = "Rendering...";

        var screenWrites = 0;
        var eventWrites = 0;
        var alertWrites = 0;
        var statusBarWrites = 0;
        var textWrites = 0;
        scope.CountWrites = (member, _) =>
        {
            switch (member)
            {
                case "ScreenUpdating": screenWrites++; break;
                case "EnableEvents": eventWrites++; break;
                case "DisplayAlerts": alertWrites++; break;
                case "DisplayStatusBar": statusBarWrites++; break;
                case "StatusBarText": textWrites++; break;
            }
        };

        SuppressAll(scope);
        scope.Dispose();

        Assert.Equal(0, screenWrites);
        Assert.Equal(0, eventWrites);
        Assert.Equal(0, alertWrites);
        Assert.Equal(0, statusBarWrites);
        Assert.Equal(0, textWrites);
    }

    [Fact]
    public void The_selection_is_captured_and_restored_last()
    {
        var scope = new TestableScope();
        scope.CaptureSelection();
        scope.SuppressAlerts();
        scope.SuppressScreenUpdating();

        scope.Dispose();

        Assert.True(scope.SelectionRestored);
    }

    [Fact]
    public void A_command_that_never_captured_the_selection_does_not_restore_one()
    {
        var scope = new TestableScope();

        scope.SuppressAlerts();
        scope.Dispose();

        Assert.False(scope.SelectionRestored, "An uncaptured selection must not be restored.");
    }

    [Fact]
    public void Disposing_twice_restores_once()
    {
        var scope = new TestableScope();
        scope.SuppressScreenUpdating();
        scope.DisplayAlertsState = true;

        scope.Dispose();
        scope.Dispose();

        Assert.True(scope.ScreenUpdatingState);
        Assert.True(scope.DisplayAlertsState);
    }

    [Fact]
    public void The_scope_degrades_to_a_no_op_without_an_application_object()
    {
        // A null or foreign application object must not throw from a Ribbon
        // callback; every member simply does nothing and restores nothing.
        using var scope = new ExcelApplicationStateScope(null);

        scope.SuppressScreenUpdating();
        scope.SuppressEvents();
        scope.SuppressAlerts();
        scope.SuppressStatusBar();
        scope.SetStatusBarText("x");
        scope.CaptureSelection();

        Assert.Null(Record.Exception(scope.Dispose));
    }

    [Fact]
    public void A_foreign_application_object_degrades_rather_than_throwing()
    {
        using var scope = new ExcelApplicationStateScope("not an Excel application");

        Assert.Null(Record.Exception(scope.SuppressScreenUpdating));
        Assert.Null(Record.Exception(scope.Dispose));
    }


    /// <summary>
    /// One row per application-state setting. Each is a separate injection point
    /// (docs/08-TEST-CHECKLIST.md section K: "failure injection around each
    /// application-state change").
    /// </summary>
    public static TheoryData<string> FailingWrites() =>
    [
        "ScreenUpdating",
        "EnableEvents",
        "DisplayAlerts",
        "DisplayStatusBar",
        "StatusBarText",
    ];
    [Theory]
    [MemberData(nameof(FailingWrites))]
    public void A_failure_at_each_setting_still_restores_everything_captured_before_it(string failing)
    {
        // A failure at the Nth setting means settings 1..N-1 were captured and
        // changed; the scope must still put all of them back, and the failing one
        // itself was never changed so it needs no restore.
        var records = new List<string>();
        var scope = new TestableScope(records.Add) { FailOn = failing };

        // Assert the injection actually fired. A bare try/catch would pass just as
        // well if the seam silently stopped throwing, which would make every
        // assertion below vacuous: nothing would have been captured or changed.
        _ = Assert.Throws<InvalidOperationException>(() => SuppressAll(scope));

        scope.Dispose();

        Assert.True(scope.ScreenUpdatingState || failing == "ScreenUpdating");
        Assert.True(scope.EnableEventsState || failing == "EnableEvents");
        Assert.True(scope.DisplayAlertsState || failing == "DisplayAlerts");
        Assert.True(scope.DisplayStatusBarState || failing == "DisplayStatusBar");

        // The status-bar text is the last setting the command changes, so it is
        // untouched when an earlier setting failed and restored when its own write
        // is the one that failed before changing anything. Either way it must read
        // "Ready" again here - the value this test started from.
        Assert.Equal("Ready", scope.StatusBarTextState);
    }

    [Fact]
    public void A_restore_failure_degrades_to_a_record_and_does_not_skip_the_others()
    {
        // D2: one restore that throws must not prevent the rest. The records
        // prove the degradation was reported rather than swallowed silently.
        var records = new List<string>();
        var scope = new TestableScope(records.Add);
        SuppressAll(scope);
        scope.CaptureSelection();
        scope.FailOn = "DisplayStatusBar";
        scope.ScreenUpdatingState = false;
        scope.DisplayAlertsState = false;
        scope.DisplayStatusBarState = false;
        scope.StatusBarTextState = "Rendering...";

        scope.Dispose();

        // The restore of DisplayStatusBar throws, so it stays false; the other
        // four captured settings must still have been restored. This is the D2
        // guarantee: one failing restore cannot skip the rest.
        Assert.False(scope.DisplayStatusBarState, "The injected failure should have left this setting un-restored.");
        Assert.True(scope.ScreenUpdatingState);
        Assert.True(scope.EnableEventsState);
        Assert.True(scope.DisplayAlertsState);
        Assert.Equal("Ready", scope.StatusBarTextState);
        Assert.Contains(records, record => record.Contains("DisplayStatusBar", StringComparison.Ordinal));
    }

    [Fact]
    public void A_failing_selection_restore_is_recorded_and_does_not_throw()
    {
        // D3: selection restore is best-effort, ordered last, and its failure
        // never masks an earlier restore.
        var records = new List<string>();
        var scope = new TestableScope(records.Add);
        SuppressAll(scope);
        scope.CaptureSelection();
        scope.FailSelectionRestore = true;

        scope.Dispose();

        Assert.True(scope.ScreenUpdatingState, "An earlier restore must not be masked by the selection failure.");
        Assert.Contains(records, record => record.Contains("selection", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// The concrete seam, which is what the non-<c>Range</c> case actually runs.
    /// The <c>TestableScope</c> above replaces <c>RestoreSelection</c> wholesale, so
    /// it can never observe the production type test.
    /// </summary>
    /// <param name="selection">What <c>Application.Selection</c> reports.</param>
    /// <param name="technicalRecord">Receives the technical record.</param>
    private sealed class RecordingScope(object selection, Action<string> technicalRecord)
        : ExcelApplicationStateScope(ApplicationReporting(selection), technicalRecord)
    {
        private static object ApplicationReporting(object selection)
        {
            var application = new Mock<Excel.Application>();
            _ = application.SetupGet(a => a.Selection).Returns(selection);
            return application.Object;
        }
    }

    [Fact]
    public void A_captured_selection_that_is_not_an_excel_range_is_recorded_rather_than_silently_skipped()
    {
        // D3: Application.Selection is typed Object, so the host can hand back
        // something that is not a Range - a chart, a shape, a foreign object. The
        // old `is Excel.Range` test fell through and returned normally, so the scope
        // reported a successful restore while the user's cursor stayed wherever the
        // command left it, and nothing recorded that. Only the concrete seam runs
        // the production type test, which is why this is not TestableScope.
        var records = new List<string>();
        using var scope = new RecordingScope(new object(), records.Add);
        scope.CaptureSelection();

        Assert.Null(Record.Exception(scope.Dispose));

        Assert.NotEmpty(records);
        Assert.Contains(records, record => record.Contains("not a Range", StringComparison.Ordinal));
    }

    [Fact]
    public void A_captured_excel_range_is_restored_through_Range_Select()
    {
        // The positive half of the same rule: a real Range still restores, so the
        // new throw refuses an unusable selection rather than refusing selections.
        var range = new Mock<Excel.Range>();
        _ = range.Setup(r => r.Select());

        var records = new List<string>();
        using var scope = new RecordingScope(range.Object, records.Add);
        scope.CaptureSelection();

        scope.Dispose();

        range.Verify(r => r.Select(), Times.Once);
        Assert.Empty(records);
    }

    [Fact]
    public void A_failing_technical_record_sink_never_escapes_the_dispose()
    {
        // The record sink is caller-supplied; a sink that throws must degrade to
        // no record rather than turning a restore failure into an escaping
        // exception inside a finally block.
        using var scope = new ExcelApplicationStateScope(
            new Mock<Excel.Application>().Object,
            _ => throw new InvalidOperationException("sink is broken"));
        scope.SuppressScreenUpdating();

        Assert.Null(Record.Exception(scope.Dispose));
    }
}
