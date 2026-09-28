using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;
using GanttCreator.Office;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// Live-Excel gate for the R4.2 application-state scope. Tagged
/// <c>[Trait("Category","OfficeIntegration")]</c> so <c>verify-quick.ps1</c> and
/// <c>verify.ps1</c> exclude it; run via <c>pwsh ./scripts/verify-office.ps1</c>.
/// </summary>
/// <remarks>
/// <para>
/// The roadmap marks R4.2 Office gate <b>Required</b>: "force a live command
/// failure and confirm Excel state is restored". This test is the automated half.
/// The command it needs does not exist until R4.9, so it drives the scope
/// directly against a real Excel application: it records the live values, changes
/// the five settings (ADR-0020; calculation mode is deliberately absent) and the
/// selection, throws <em>inside</em> the scope exactly as a failing
/// command would, and asserts every one is back.
/// </para>
/// <para>
/// This is the first Phase-4 Office evidence row, so the host Windows/Office
/// build is recorded in the evidence output (roadmap Phase 4 note). The contract
/// assertions mirror <c>ContractTests.ExcelApplicationStateScopeTests</c> over
/// Moq, so drift between the faked property surface and the real one is caught.
/// </para>
/// </remarks>
public class ApplicationStateScopeIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    private static string XllPath => OfficeFixtureTests.ResolvePackedXllPath();

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task A_failure_inside_the_scope_restores_every_application_setting()
    {
        var fixture = new OfficeFixture();
        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            Assert.True(
                fixture.RegisterXll(XllPath),
                $"Application.RegisterXLL returned false for '{XllPath}'.");

            // A workbook must exist before Application.Selection is non-null: with
            // no open workbook the host reports no selection at all. The fixture's
            // teardown closes it, so the test does not own its disposal.
            Excel.Workbook workbook = fixture.CreateWorkbook();
            Excel.Application application = fixture.Excel;
            _output.WriteLine("Windows build: " + Environment.OSVersion.Version.Build);
            _output.WriteLine("Excel version: " + application.Version);

            // Record the live values BEFORE the scope touches anything, so the
            // assertion compares against the host own starting state rather than
            // against a value this test chose.
            bool screenBefore = application.ScreenUpdating;
            bool eventsBefore = application.EnableEvents;
            bool alertsBefore = application.DisplayAlerts;
            bool statusBarVisibleBefore = application.DisplayStatusBar;
            object? statusTextBefore = application.StatusBar;

            _output.WriteLine("before: " + Describe(application));
            Assert.NotNull(application.Selection);
            Assert.NotNull(workbook);

            var records = new List<string>();
            var thrown = new InvalidOperationException("injected mid-render failure");

            try
            {
                using var scope = new ExcelApplicationStateScope(application, records.Add);
                scope.SuppressScreenUpdating();
                scope.SuppressEvents();
                scope.SuppressAlerts();
                scope.SuppressStatusBar();
                scope.SetStatusBarText("Rendering...");
                scope.CaptureSelection();

                // The settings really did change inside the scope: prove it before
                // the throw, or the restore assertions below would be vacuous.
                Assert.False(application.ScreenUpdating);
                Assert.False(application.EnableEvents);
                Assert.False(application.DisplayAlerts);
                Assert.False(application.DisplayStatusBar);

                throw thrown;
            }
            catch (InvalidOperationException caught) when (ReferenceEquals(caught, thrown))
            {
                // Expected: the simulated command failure. The scope Dispose has
                // already run on the way out of the using block.
            }

            _output.WriteLine("after: " + Describe(application));

            Assert.Equal(screenBefore, application.ScreenUpdating);
            Assert.Equal(eventsBefore, application.EnableEvents);
            Assert.Equal(alertsBefore, application.DisplayAlerts);
            Assert.Equal(statusBarVisibleBefore, application.DisplayStatusBar);

            // StatusBar is Object-typed, and Excel does not round-trip its
            // representation: an empty status bar reads as the bool False, but
            // after a write and a restore it reads back as the string "FALSE".
            // The value is equivalent, so the comparison is case-insensitive
            // rather than the test failing on a representation change the scope
            // did not cause. This is recorded in the work item's Notes.
            Assert.Equal(
                Convert.ToString(statusTextBefore, CultureInfo.InvariantCulture),
                Convert.ToString(application.StatusBar, CultureInfo.InvariantCulture),
                ignoreCase: true);

            // The selection restore is best-effort by contract, so a record here
            // is reported rather than failed. An empty list is the expected result
            // on the reference host and is itself worth recording.
            _output.WriteLine("restore records: " + string.Join(" | ", records));
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    /// <summary>Renders the live application state as one evidence line.</summary>
    /// <param name="application">The live Excel application.</param>
    /// <returns>The rendered state.</returns>
    private static string Describe(Excel.Application application) =>
        "ScreenUpdating=" + application.ScreenUpdating
        + " EnableEvents=" + application.EnableEvents
        + " DisplayAlerts=" + application.DisplayAlerts
        + " DisplayStatusBar=" + application.DisplayStatusBar
        + " StatusBar=" + Convert.ToString(application.StatusBar, CultureInfo.InvariantCulture);
}
