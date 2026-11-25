using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using GanttCreator.Core;
using GanttCreator.Core.Export;
using GanttCreator.Core.Scene;
using Xunit.Abstractions;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// R7.7 TransferSoak: 25 sequential transfers over the landed production R6 + R7 chain,
/// measuring stability (D1 counters), asserting per-iteration cleanup (D2), and proving no
/// orphaned Office process survives the run (D5's non-destructive ownership model).
/// </summary>
/// <remarks>
/// <para>
/// <b>Category split (D4).</b> Tagged <c>TransferSoak</c> only. It is therefore excluded
/// from the routine <c>Category=OfficeIntegration</c> gate and from the copy after a
/// failing run; the routine <i>negated</i> filters in <c>test-non-office.ps1</c> and
/// <c>verify.ps1</c> exclude it as well (R7.7 Phase 1), and it runs only through
/// <c>scripts/verify-transfer-soak.ps1</c> with that script's measured deadline.
/// </para>
/// <para>
/// <b>Two arms.</b> Iterations 1–15 attach to one long-lived fixture-owned deck carrying an
/// unsaved marker — the R7.6 body-1 simulation of a user's unsaved session (a real user
/// session is never mutated by an automated gate) — and cycle 4:3, 16:9 and a custom
/// 1200×400 slide. Iterations 16–25 dispose and restart the host with no presentation, so
/// the transfer creates the deck (the R7.6 body-2 arm); that is the process-lifecycle
/// churn where an orphan would surface.
/// </para>
/// <para>
/// <b>Pre-condition.</b> The body refuses to start while any <c>POWERPNT</c> process runs.
/// With a user host present the fixture's <c>new Application()</c> attaches to it
/// (<c>_startedByUs</c> stays false) and the target resolver may read the user's active
/// deck as the transfer target. The R7.2 gate's <c>WaitForNoPowerPointAsync</c> poll is
/// reused, including between Arm B iterations so a fresh start cannot attach to a dying
/// instance.
/// </para>
/// <para>
/// <b>Race versus orphan.</b> The known Excel shutdown race (R5.18n D4) is a fixture-owned
/// process the fixture itself kills after <c>Quit()</c> returns, recorded in the
/// forced-kill log and bounded by the per-teardown ceiling; it is not an orphan. The orphan
/// verdict here is the end-state live-PID subtraction against the pre-run snapshot — a
/// process the race cannot produce.
/// </para>
/// <para>
/// <b>4:3 residue.</b> The staged composition already overflows some text children at 4:3
/// (the open R7.4 owner decision). This body asserts the placement-side rule only and
/// records the residue; it does not re-litigate the composition decision.
/// </para>
/// <para>
/// <b>Tolerances.</b> The handle and working-set tolerances below are provisional bounds
/// pending the first reference-host measured run, which scripts/verify-transfer-soak.ps1
/// archives. The measured series is recorded with the result; the constants are then
/// pinned from it with headroom, exactly as the routine gate's 780 s deadline was.
/// </para>
/// </remarks>
public sealed class PowerPointTransferSoakIntegrationTests(ITestOutputHelper output)
{
    private const int _iterations = 25;
    private const int _attachIterations = 15;

    // Measured on the reference host 2026-11-22 (the first 25-transfer run): the attach-arm
    // handle series read 1973, 1968, 1968, 1956, 1960, 1962, 1959, 1961, 1961, 1961, 1961,
    // 1959, 1949, 1949, 1936 — min 1936, max 1973, delta 37, with the series DECLINING after
    // the iteration-1 warm-up (no monotonic growth, i.e. no leak). The tolerance is the
    // measured spread plus headroom, exactly as the routine gate's 780 s deadline was set.
    private const int _handleTrendTolerance = 50;

    // Measured in the same run: first-third median 307 MB against last-third 315 MB, delta
    // 8 MB; the recorded tolerance is a coarse bound so a single OS trim cannot trip it.
    private const long _workingSetTrendToleranceBytes = 64L * 1024 * 1024;
    private const string _markerText = "R7.7 soak unsaved marker";

    private static readonly (double WidthPt, double HeightPt)[] _slideSizes =
        [(720, 540), (960, 540), (1200, 400)];

    private readonly ITestOutputHelper _output = output;

    [Trait("Category", "TransferSoak")]
    [Fact]
    public async Task Twenty_five_transfers_stay_clean_with_no_orphans_and_a_flat_handle_trend()
    {
        // Pre-condition: refuse to start while a user PowerPoint is running (F1).
        Dictionary<int, long> preExisting = await WaitForNoPowerPointAsync().ConfigureAwait(true);
        Assert.True(
            preExisting.Count == 0,
            $"A POWERPNT process is already running (PID(s): {string.Join(", ", preExisting.Keys)}). The soak mutates only fixture-owned decks, and with a user host present the resolver can target the user's active deck; close PowerPoint before running it.");

        HashSet<int> excelBefore = LiveProcessIds("EXCEL");
        HashSet<int> powerPointBefore = LiveProcessIds("POWERPNT");

        OfficeFixture excelFixture = new();
        PowerPointFixture powerPointFixture = new();
        List<int> attachHandles = [];
        List<long> attachWorkingSets = [];
        List<double> timings = [];
        HashSet<int> startedPowerPointPids = [];
        Dictionary<int, long> identityTicks = [];
        string? deckPath = null;
        try
        {
            await excelFixture.InitializeAsync().ConfigureAwait(true);

            using OfficeFixture.ComScope scope = new();
            ExportBoundsOutcome resolution =
                StagingLifecycleIntegrationTests.Prepare(excelFixture, scope).Resolution;

            // Arm A — attach (iterations 1–15) on one long-lived fixture-owned host.
            powerPointFixture.Initialize(withWindow: true, slideCount: 1);
            startedPowerPointPids.UnionWith(powerPointFixture.OwnedProcessIds);
            deckPath = powerPointFixture.SavePrimaryPresentationToTemp();
            DateTime savedAt = File.GetLastWriteTimeUtc(deckPath);
            powerPointFixture.AddUnsavedMarker(_markerText);
            Assert.True(powerPointFixture.PrimaryPresentationHasUnsavedChanges());

            for (int iteration = 1; iteration <= _attachIterations; iteration++)
            {
                (double widthPt, double heightPt) = _slideSizes[(iteration - 1) % _slideSizes.Length];
                powerPointFixture.SetSlideSize(widthPt, heightPt);
                (double readWidthPt, double readHeightPt) = powerPointFixture.ReadSlideSize();
                Assert.Equal(widthPt, readWidthPt, 3);
                Assert.Equal(heightPt, readHeightPt, 3);

                int slideShapesBefore = powerPointFixture.PrimarySlideShapeCount();
                timings.Add(TransferOnce(excelFixture, resolution));

                Assert.True(
                    powerPointFixture.PrimarySlideHasText(_markerText),
                    $"iteration {iteration}: the unsaved marker was lost — the transfer touched a different deck.");
                Assert.True(
                    powerPointFixture.PrimaryPresentationHasUnsavedChanges(),
                    $"iteration {iteration}: the deck stopped being dirty — something saved the user's session.");
                Assert.Equal(savedAt, File.GetLastWriteTimeUtc(deckPath));
                Assert.Equal(slideShapesBefore, powerPointFixture.PrimarySlideShapeCount());

                SoakProcessCounters[] samples = SampleOwned(powerPointFixture.OwnedProcessIds, identityTicks);
                foreach (SoakProcessCounters sample in samples)
                {
                    if (!sample.Alive)
                    {
                        continue;
                    }

                    attachHandles.Add(sample.HandleCount);
                    attachWorkingSets.Add(sample.WorkingSetBytes);
                }

                SoakProcessCounters[] excelSamples = SampleOwned(excelFixture.OwnedProcessIds, identityTicks);
                _output.WriteLine(string.Create(
                    CultureInfo.InvariantCulture,
                    $"soak iter={iteration} arm=attach slide={widthPt}x{heightPt} totalMs={timings[^1]:F1} slideShapes={slideShapesBefore} residue=0 markerAlive=true deckDirty=true handles={JoinAlive(samples, static sample => sample.HandleCount)} excelHandles={JoinAlive(excelSamples, static sample => sample.HandleCount)} workingSetMb={JoinAlive(samples, static sample => sample.WorkingSetBytes / (1024 * 1024))}"));
            }

            // Arm B — fresh create (iterations 16–25): restart the host with no deck, so the
            // transfer creates it (the R7.6 body-2 arm and the lifecycle churn).
            for (int iteration = _attachIterations + 1; iteration <= _iterations; iteration++)
            {
                await powerPointFixture.DisposeAsync().ConfigureAwait(true);
                powerPointFixture = new PowerPointFixture();

                Dictionary<int, long> residualHost = await WaitForNoPowerPointAsync().ConfigureAwait(true);
                Assert.True(
                    residualHost.Count == 0,
                    $"iteration {iteration}: the previous host had not exited before the fresh start (PID(s): {string.Join(", ", residualHost.Keys)}).");

                powerPointFixture.InitializeWithoutPresentation();
                startedPowerPointPids.UnionWith(powerPointFixture.OwnedProcessIds);

                int editableChildren = TransferCreateOnce(
                    excelFixture,
                    resolution,
                    powerPointFixture,
                    out double totalMs);
                timings.Add(totalMs);
                Assert.True(
                    editableChildren > 0,
                    $"iteration {iteration}: the created deck's paste did not ungroup to editable children.");

                SoakProcessCounters[] samples = SampleOwned(powerPointFixture.OwnedProcessIds, identityTicks);
                SoakProcessCounters[] excelSamples = SampleOwned(excelFixture.OwnedProcessIds, identityTicks);
                _output.WriteLine(string.Create(
                    CultureInfo.InvariantCulture,
                    $"soak iter={iteration} arm=fresh-create totalMs={totalMs:F1} children={editableChildren} residue=0 handles={JoinAlive(samples, static sample => sample.HandleCount)} excelHandles={JoinAlive(excelSamples, static sample => sample.HandleCount)} workingSetMb={JoinAlive(samples, static sample => sample.WorkingSetBytes / (1024 * 1024))}"));
            }
        }
        finally
        {
            await powerPointFixture.DisposeAsync().ConfigureAwait(true);
            await excelFixture.DisposeAsync().ConfigureAwait(true);
            TryDelete(deckPath);
        }

        // End-state orphan verdict: no Office process that was not already running.
        HashSet<int> residual = await WaitForOfficeProcessesToExitAsync(excelBefore, powerPointBefore)
            .ConfigureAwait(true);
        Assert.True(
            residual.Count == 0,
            $"Unattributed Office process(es) survived the soak: {string.Join(", ", residual)}. Ownership attribution comes from the owned-PID manifest and the sweep in scripts/verify-transfer-soak.ps1.");

        Assert.Equal(_iterations, timings.Count);

        SoakTrendOutcome handleTrend = SoakTrend.Handles(attachHandles, _handleTrendTolerance);
        SoakTrendOutcome workingSetTrend = SoakTrend.WorkingSet(attachWorkingSets, _workingSetTrendToleranceBytes);

        // The manifest assertion is authoritative only when the gate script armed the
        // owned-PID manifest; a raw filter run records manifestArmed=false, and the script
        // fails closed on that (F2/R12).
        string? manifestPath = Environment.GetEnvironmentVariable("GANTTCREATOR_OWNED_PIDS_PATH");
        bool manifestArmed = !string.IsNullOrEmpty(manifestPath) && File.Exists(manifestPath);
        if (manifestArmed)
        {
            HashSet<int> recorded = ReadManifestProcessIds(manifestPath!);
            foreach (int processId in startedPowerPointPids)
            {
                Assert.Contains(processId, recorded);
            }
        }

        _output.WriteLine(string.Create(
            CultureInfo.InvariantCulture,
            $"soak summary iterations={_iterations} attach={_attachIterations} medianMs={Median(timings):F1} maxMs={timings.Max():F1} handlesMin={handleTrend.Baseline} handlesMax={handleTrend.End} handlesDelta={handleTrend.Delta} handlesTolerance={_handleTrendTolerance} workingSetFirstThirdMb={workingSetTrend.Baseline / (1024 * 1024)} workingSetLastThirdMb={workingSetTrend.End / (1024 * 1024)} workingSetToleranceMb={_workingSetTrendToleranceBytes / (1024 * 1024)} startedPowerPointPids={startedPowerPointPids.Count} manifestArmed={manifestArmed} orphans=0"));

        Assert.True(
            handleTrend.WithinTolerance,
            $"Handle-count trend drifted: min={handleTrend.Baseline} max={handleTrend.End} delta={handleTrend.Delta} > tolerance={_handleTrendTolerance} over {attachHandles.Count} attach-arm samples (the first {SoakTrend.HandleWarmUpSamples} are excluded as host warm-up).");
        Assert.True(
            workingSetTrend.WithinTolerance,
            $"Working-set trend drifted: first-third={workingSetTrend.Baseline} last-third={workingSetTrend.End} delta={workingSetTrend.Delta} > tolerance={_workingSetTrendToleranceBytes}.");
    }

    private double TransferOnce(OfficeFixture excelFixture, ExportBoundsOutcome resolution)
    {
        Stopwatch stopwatch = Stopwatch.StartNew();
        StagedGroupOutcome group = StageAndCopy(excelFixture, resolution);

        using PowerPointTransferAdapter adapter = new(new PowerPointProxyOwnership(), _output.WriteLine);
        PowerPointTargetDecision decision = new PowerPointTargetResolver(adapter).Resolve();
        Assert.Equal(PowerPointTargetAction.ProceedWithActiveSlide, decision.Action);

        int groupId = PasteAndPlace(adapter, group, resolution);
        PowerPointGroupDeleteResult deleted = adapter.DeletePastedGroup(groupId);
        Assert.Equal(PowerPointGroupDeleteOutcome.Deleted, deleted.Outcome);

        stopwatch.Stop();
        return stopwatch.Elapsed.TotalMilliseconds;
    }

    private int TransferCreateOnce(
        OfficeFixture excelFixture,
        ExportBoundsOutcome resolution,
        PowerPointFixture fixture,
        out double totalMs)
    {
        Stopwatch stopwatch = Stopwatch.StartNew();
        StagedGroupOutcome group = StageAndCopy(excelFixture, resolution);

        using PowerPointTransferAdapter adapter = new(new PowerPointProxyOwnership(), _output.WriteLine);
        PowerPointTargetDecision decision = new PowerPointTargetResolver(adapter).Resolve();
        Assert.Equal(PowerPointTargetAction.CreatePresentation, decision.Action);

        PowerPointPresentationCreation creation = adapter.CreatePresentationWithBlankSlide();
        Assert.True(creation.Created, $"Creation refused: hresult=0x{creation.HResult:X8}");

        int groupId = PasteAndPlace(adapter, group, resolution);
        int editableChildren = fixture.UngroupShapeById(groupId);
        fixture.DeletePastedShapes();

        stopwatch.Stop();
        totalMs = stopwatch.Elapsed.TotalMilliseconds;
        return editableChildren;
    }

    private static StagedGroupOutcome StageAndCopy(OfficeFixture excelFixture, ExportBoundsOutcome resolution)
    {
        ExcelShapeWriter stagingWriter = new(
            excelFixture.Excel,
            ownershipScheme: ShapeOwnershipScheme.Staging);

        using ExcelApplicationStateScope state = new(excelFixture.Excel);
        StagingLifecycleOutcome lifecycle = StagingLifecycle.Run(new(
            state,
            stagingWriter,
            new ExcelShapeGrouper(excelFixture.Excel),
            resolution,
            Copy: group => StagedShapeCopier.Copy(
                group,
                new ExcelShapeCopier(excelFixture.Excel),
                new WindowsClipboardFormatReader())));

        Assert.True(lifecycle.Succeeded, $"Lifecycle refused: {lifecycle.Refusal}");
        Assert.Empty(lifecycle.Residue);
        Assert.NotNull(lifecycle.Copy);
        Assert.True(
            lifecycle.Copy.Succeeded,
            $"Copy refused: {lifecycle.Copy.Refusal} {lifecycle.Copy.PortRefusal} {lifecycle.Copy.Detail}");
        return Assert.IsType<StagedGroupOutcome>(lifecycle.Group);
    }

    private static int PasteAndPlace(
        PowerPointTransferAdapter adapter,
        StagedGroupOutcome group,
        ExportBoundsOutcome resolution)
    {
        RectD bounds = Assert.IsType<RectD>(group.Bounds);
        double insetPt = ChartPaddingPt.Uniform(
            GanttCatalogues.MetricDefault("ChartOuterPaddingPt")).MaxPt;

        PowerPointPasteStageResult paste = new PowerPointPasteStage(adapter).Paste(
            new PowerPointPasteExpectation(group.ChildOrder.Count, bounds.Width, bounds.Height));
        Assert.True(paste.Succeeded, $"Paste refused: {paste.Refusal} {paste.Validation?.Refusal}");
        int groupId = paste.GroupId
            ?? throw new InvalidOperationException("The paste stage reported no group identity.");

        PowerPointPlacementResult placed = new PowerPointPlacementStage(adapter).Place(
            groupId,
            resolution.Bounds!.Bounds,
            insetPt);
        Assert.True(
            placed.Succeeded,
            $"Placement refused: {placed.Refusal} fit={placed.FitRefusal} hresult=0x{placed.HResult:X8}");
        Assert.True(
            Assert.IsType<PowerPointPlacementDeltas>(placed.Deltas).MaxPt
                <= HostOwnedGeometry.GeometryMatchTolerancePt,
            "The placement read-back drifted.");
        return groupId;
    }

    private static async Task<Dictionary<int, long>> WaitForNoPowerPointAsync()
    {
        DateTime deadline = DateTime.UtcNow + TimeSpan.FromSeconds(15);
        Dictionary<int, long> snapshot = PowerPointFixture.SnapshotProcesses();
        while (snapshot.Count > 0 && DateTime.UtcNow < deadline)
        {
            await Task.Delay(TimeSpan.FromMilliseconds(100)).ConfigureAwait(false);
            snapshot = PowerPointFixture.SnapshotProcesses();
        }

        return snapshot;
    }

    private static async Task<HashSet<int>> WaitForOfficeProcessesToExitAsync(
        HashSet<int> excelBefore,
        HashSet<int> powerPointBefore)
    {
        DateTime deadline = DateTime.UtcNow + TimeSpan.FromSeconds(15);
        HashSet<int> residual = ResidualOfficeProcesses(excelBefore, powerPointBefore);
        while (residual.Count > 0 && DateTime.UtcNow < deadline)
        {
            await Task.Delay(TimeSpan.FromMilliseconds(100)).ConfigureAwait(false);
            residual = ResidualOfficeProcesses(excelBefore, powerPointBefore);
        }

        return residual;
    }

    private static HashSet<int> ResidualOfficeProcesses(HashSet<int> excelBefore, HashSet<int> powerPointBefore)
    {
        HashSet<int> residual = [];
        foreach (int processId in LiveProcessIds("EXCEL"))
        {
            if (!excelBefore.Contains(processId))
            {
                _ = residual.Add(processId);
            }
        }

        foreach (int processId in LiveProcessIds("POWERPNT"))
        {
            if (!powerPointBefore.Contains(processId))
            {
                _ = residual.Add(processId);
            }
        }

        return residual;
    }

    private static HashSet<int> LiveProcessIds(string processName)
    {
        HashSet<int> ids = [];
        foreach (Process process in Process.GetProcessesByName(processName))
        {
            using (process)
            {
                _ = ids.Add(process.Id);
            }
        }

        return ids;
    }

    private static SoakProcessCounters[] SampleOwned(
        IReadOnlyList<int> ownedProcessIds,
        Dictionary<int, long> identityTicks)
    {
        SoakProcessCounters[] samples = new SoakProcessCounters[ownedProcessIds.Count];
        for (int index = 0; index < ownedProcessIds.Count; index++)
        {
            int processId = ownedProcessIds[index];
            SoakProcessCounters sample = SoakProcessSampler.Sample(processId);
            samples[index] = sample;
            if (!sample.Alive)
            {
                continue;
            }

            if (identityTicks.TryGetValue(processId, out long ticks) && ticks != sample.StartTimeUtcTicks)
            {
                throw new InvalidOperationException(
                    $"Owned PID {processId} changed identity between samples (recorded ticks {ticks}, live {sample.StartTimeUtcTicks}); the PID was recycled and the counters cannot be attributed.");
            }

            identityTicks[processId] = sample.StartTimeUtcTicks;
        }

        return samples;
    }

    private static HashSet<int> ReadManifestProcessIds(string path)
    {
        HashSet<int> ids = [];
        foreach (string line in File.ReadLines(path))
        {
            string trimmed = line.Trim();
            if (trimmed.Length == 0)
            {
                continue;
            }

            try
            {
                using JsonDocument document = JsonDocument.Parse(trimmed);
                if (document.RootElement.TryGetProperty("ProcessId", out JsonElement processId)
                    && processId.ValueKind == JsonValueKind.Number)
                {
                    _ = ids.Add(processId.GetInt32());
                }
            }
            catch (JsonException)
            {
                // One unreadable line must not discard the rest of the manifest — the same
                // rule the shell reader follows (verify-office.ps1 Read-OwnedPidManifest).
            }
        }

        return ids;
    }

    private static string JoinAlive<T>(
        IReadOnlyList<SoakProcessCounters> samples,
        Func<SoakProcessCounters, T> selector) =>
        string.Join('/', samples.Where(static sample => sample.Alive).Select(selector));

    private static double Median(List<double> values)
    {
        if (values.Count == 0)
        {
            return 0;
        }

        double[] sorted = [.. values.OrderBy(value => value)];
        int middle = sorted.Length / 2;
        return sorted.Length % 2 == 1
            ? sorted[middle]
            : (sorted[middle - 1] + sorted[middle]) / 2;
    }

    private static void TryDelete(string? path)
    {
        if (path is null)
        {
            return;
        }

        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (IOException)
        {
            // The file lives in the OS temp folder; a failed delete must not fail the body.
        }
        catch (UnauthorizedAccessException)
        {
            // Same.
        }
    }
}
