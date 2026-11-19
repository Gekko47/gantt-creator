using System.Globalization;
using System.Runtime.InteropServices;
using GanttCreator.AddIn;
using GanttCreator.Core.Scene;
using Xunit.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// R6.8 Office gate: the production Copy Editable command runs the whole pipeline on a real
/// host — read, configuration, validate, measure, export request, selection, bounds, then the
/// R6.5 lifecycle through the R6.6 verified copy — and leaves zero residue with the view and
/// selection
/// restored.
/// </summary>
/// <remarks>
/// <para>
/// <b>The command is driven, not replicated.</b> The body calls
/// <see cref="CopyEditableCommand.Run"/> with the real adapters over the fixture's Excel
/// object, so the one-click contract is exercised end to end. `RunForExcel` is deliberately
/// not called: it resolves `ExcelDnaUtil.Application`, which only exists inside the Excel-DNA
/// host.
/// </para>
/// <para>
/// <b>Materialisation is the primary measurement.</b> After the command's cleanup the body
/// pastes into a test-owned cell, identifies the pasted range by the shape-count delta,
/// asserts positive bounds and the staged child count, ungroups to prove editability, and
/// deletes the pasted shapes — the R6.6 discipline, reused rather than restated.
/// </para>
/// <para>
/// <b>The busy-clipboard body records what the host does (plan C3).</b> Holding the clipboard
/// open makes `Shape.Copy` itself fail, so the typed refusal arrives at the copy port rather
/// than the reader's post-copy read. The assertions are therefore channel-independent — a
/// refusal, never a success; one actionable line; zero residue — and the observed refusal
/// member is recorded for the evidence table.
/// </para>
/// </remarks>
public sealed partial class CopyEditableIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    /// <summary>
    /// C8 wiring spike: the real command type and its pure helpers load in this host before
    /// any live body relies on them, so a dependency or internals-visibility failure surfaces
    /// here as a fast host-free fact rather than as a confusing live failure.
    /// </summary>
    [Fact]
    public void The_command_is_reachable_from_the_integration_harness()
    {
        ExportSelection selection = CopyEditableCommand.SelectExportSelection(
            new Dictionary<string, string>
            {
                [CopyEditableCommand.ExportIncludeDataPanelSettingKey] = "FALSE",
            });

        Assert.Equal(ExportSelection.ChartOnly, selection);
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task One_click_copies_the_verified_group_and_materialises_after_cleanup()
    {
        OfficeFixture fixture = new();

        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            using OfficeFixture.ComScope scope = new();
            (
                Excel.Worksheet sheet,
                Excel.Sheets sheets,
                _,
                Excel.Shape sentinel,
                Excel.Range witness,
                object? witnessBefore,
                int sheetCountBefore) = StagingLifecycleIntegrationTests.Prepare(fixture, scope);

            ExcelShapeWriter stagingWriter = new(
                fixture.Excel,
                ownershipScheme: ShapeOwnershipScheme.Staging);
            ExcelShapeWriter liveWriter = new(fixture.Excel);
            Assert.True(liveWriter.Create(StagingLifecycleIntegrationTests.LiveProbe("live:probe-1")).Succeeded);

            (int scrollRow, int scrollColumn, double zoom, string selection) =
                StagingLifecycleIntegrationTests.CaptureView(fixture, scope);

            List<string> successes = [];
            List<string> failures = [];
            List<string> records = [];

            // The lifecycle disposes the scope inside its finally; `using` is the
            // analyzer-satisfying double dispose, which ExcelApplicationStateScope guards.
            using ExcelApplicationStateScope state = new(fixture.Excel);
            CopyEditableOutcome outcome = CopyEditableCommand.Run(
                LiveRequest(fixture, state, records.Add),
                successes.Add,
                failures.Add);

            Assert.True(
                outcome.Succeeded,
                string.Create(
                    CultureInfo.InvariantCulture,
                    $"The one-click copy refused: {outcome.Refusal}; message={outcome.Message}"));
            Assert.Equal([outcome.Message], successes);
            Assert.Empty(failures);
            // The sink is shared with the writers, so it carries their routine diagnostics;
            // what it must not carry is a dropped-row record on a clean table.
            Assert.DoesNotContain(
                records,
                record => record.Contains("rows-dropped", StringComparison.Ordinal));
            Assert.True(outcome.StagedShapeCount > 0, "The copied composition staged no shapes.");
            Assert.Contains("shape", outcome.Message, StringComparison.Ordinal);

            // Independent probe: the measured drawing pair survived the cleanup.
            ClipboardFormatProbe probe = new();
            ClipboardReadResult read = probe.Read();
            Assert.True(read.Succeeded, read.Failure);
            ClipboardSnapshot snapshot = Assert.IsType<ClipboardSnapshot>(read.Snapshot);
            foreach (string required in StagedShapeCopier.RequiredDrawingFormatNames)
            {
                Assert.True(
                    snapshot.Contains(required),
                    string.Create(
                        CultureInfo.InvariantCulture,
                        $"'{required}' is absent after cleanup; formats=[{string.Join(", ", snapshot.Names)}]"));
            }

            // The lifecycle's finally restored the view and selection; assert before the
            // paste moves the selection to its destination.
            StagingLifecycleIntegrationTests.AssertViewRestored(
                fixture,
                scope,
                sheets,
                sentinel,
                witness,
                witnessBefore,
                sheetCountBefore,
                scrollRow,
                scrollColumn,
                zoom,
                selection);

            // Materialisation (ADR-0066 D2): a real consumer paste must produce the group.
            ClipboardCopyIntegrationTests.ExcelPasteOutcome paste =
                ClipboardCopyIntegrationTests.PasteIntoExcelAndVerify(
                    scope,
                    sheet,
                    outcome.StagedShapeCount);

            // Zero residue and the live chart intact.
            Assert.Empty(stagingWriter.ListOwned());
            Assert.Equal(0, StagingLifecycleIntegrationTests.CountStaged(scope, sheet));
            Assert.Equal(["live:probe-1"], liveWriter.ListOwned());

            _output.WriteLine(string.Create(
                CultureInfo.InvariantCulture,
                $"Excel {fixture.Excel.Version} build {fixture.Excel.Build};"
                    + $" staged={outcome.StagedShapeCount}"
                    + $" message='{outcome.Message}'"
                    + $" pasteShapes={paste.ShapeCount} groupItems={paste.GroupItemCount}"
                    + $" width={paste.WidthPt} height={paste.HeightPt} children={paste.ChildCount}"));

            StagingLifecycleIntegrationTests.CleanupLive(liveWriter, sentinel);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task A_busy_clipboard_is_a_typed_refusal_with_no_residue()
    {
        OfficeFixture fixture = new();

        try
        {
            await fixture.InitializeAsync().ConfigureAwait(true);
            using OfficeFixture.ComScope scope = new();
            (
                Excel.Worksheet sheet,
                Excel.Sheets sheets,
                _,
                Excel.Shape sentinel,
                Excel.Range witness,
                object? witnessBefore,
                int sheetCountBefore) = StagingLifecycleIntegrationTests.Prepare(fixture, scope);

            ExcelShapeWriter stagingWriter = new(
                fixture.Excel,
                ownershipScheme: ShapeOwnershipScheme.Staging);

            Assert.True(HoldClipboard(), "The test process could not hold the clipboard open.");
            try
            {
                List<string> successes = [];
                List<string> failures = [];
                List<string> records = [];

                using ExcelApplicationStateScope state = new(fixture.Excel);
                CopyEditableOutcome outcome = CopyEditableCommand.Run(
                    LiveRequest(fixture, state, records.Add),
                    successes.Add,
                    failures.Add);

                // Channel-independent assertions: the host's path may refuse at the copy
                // port (the expected measured member) or at the reader, and either is a
                // refusal. What must never happen is a reported success.
                Assert.False(outcome.Succeeded);
                Assert.Empty(successes);
                string message = Assert.Single(failures);
                Assert.False(string.IsNullOrWhiteSpace(message));
                Assert.True(outcome.Refusal is not null);

                // The lifecycle's finally ran: the staged shapes are gone and the user's
                // content is untouched.
                Assert.Empty(stagingWriter.ListOwned());
                Assert.Equal(0, StagingLifecycleIntegrationTests.CountStaged(scope, sheet));
                Assert.Equal(sheetCountBefore, sheets.Count);
                Assert.Equal(witnessBefore, witness.Value2);
                Assert.Equal("User note", sentinel.Name);

                _output.WriteLine(string.Create(
                    CultureInfo.InvariantCulture,
                    $"Excel {fixture.Excel.Version} build {fixture.Excel.Build};"
                        + $" busy-clipboard refusal={outcome.Refusal}"
                        + $" message='{message}'"));
            }
            finally
            {
                _ = NativeMethods.CloseClipboard();
            }
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    /// <summary>
    /// Builds the live request over the fixture's Excel object. This mirrors
    /// <see cref="CopyEditableCommand.RunForExcel"/> without the Excel-DNA host pieces
    /// (application resolution, protection window, tab restore), which the harness replaces
    /// with its own equivalents.
    /// </summary>
    /// <param name="fixture">The live Office fixture.</param>
    /// <param name="scope">The application-state scope the lifecycle owns and disposes.</param>
    /// <param name="trace">The technical-record sink.</param>
    /// <returns>The request for one run.</returns>
    private static CopyEditableRequest LiveRequest(
        OfficeFixture fixture,
        ExcelApplicationStateScope scope,
        Action<string> trace)
    {
        IWorksheetProtectionGuard guard = new ExcelWorksheetProtectionGuard(fixture.Excel);
        return new CopyEditableRequest(
            new ExcelGanttTableReader(fixture.Excel),
            new ExcelConfigCatalogueReader(fixture.Excel),
            new ExcelPanelGridMeasurement(fixture.Excel),
            EditableExportRequest.Create,
            ExportBoundsResolver.TryResolve,
            StagingLifecycle.Run,
            scope,
            // Mirrors the production staging writer, including the batch template reuse
            // that keeps the editable copy's create block off the per-shape style path.
            new ExcelShapeWriter(
                fixture.Excel,
                guard,
                trace,
                ShapeOwnershipScheme.Staging,
                reuseTemplates: true),
            new ExcelShapeGrouper(fixture.Excel, guard),
            new ExcelShapeCopier(fixture.Excel, guard),
            new WindowsClipboardFormatReader(),
            trace);
    }

    /// <summary>Holds the clipboard open for this thread's window station.</summary>
    /// <returns><see langword="true"/> when the clipboard was held.</returns>
    private static bool HoldClipboard() => NativeMethods.OpenClipboard(IntPtr.Zero);

    /// <summary>The Win32 clipboard open/close pair the busy-clipboard body runs with.</summary>
    private static partial class NativeMethods
    {
        [LibraryImport("user32.dll", SetLastError = true)]
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static partial bool OpenClipboard(IntPtr newOwner);

        [LibraryImport("user32.dll", SetLastError = true)]
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static partial bool CloseClipboard();
    }
}
