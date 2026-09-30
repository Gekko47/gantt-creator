using System.Diagnostics;
using System.Globalization;
using GanttCreator.Core.Scene;
using Xunit.Abstractions;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// The R3.12 representative 1,000-event benchmark.
/// </summary>
/// <remarks>
/// <para>
/// This harness deliberately adds no benchmarking package (D3). It uses
/// <see cref="Stopwatch"/> directly, always runs, and **records** the elapsed time
/// without asserting a wall-clock budget. A CI runner or a developer laptop is not
/// the reference machine, so failing on a timing threshold would produce a red gate
/// that says nothing about the product. The 250 ms judgement is recorded evidence
/// in the work item's Notes, with the hardware it was measured on, per the
/// architecture's "record the reference hardware" rule.
/// </para>
/// <para>
/// What the test <em>does</em> assert is correctness at scale: the build must
/// succeed, the validator must find nothing, and the output must be deterministic.
/// A regression that made the pipeline slower <em>and</em> wrong fails here; one
/// that made it slower alone is reported, not failed.
/// </para>
/// </remarks>
public sealed class SceneBenchmarkTests
{
    /// <summary>The synthetic event count the architecture budget is stated against.</summary>
    private const int EventCount = 1_000;

    /// <summary>The architecture's stated Core budget for validate plus build.</summary>
    private const double BudgetMs = 250;

    private readonly ITestOutputHelper _output;

    /// <summary>Initialises the benchmark with xUnit's output sink.</summary>
    /// <param name="output">The test output helper used to record measurements.</param>
    public SceneBenchmarkTests(ITestOutputHelper output) => _output = output;

    [Fact]
    public void A_thousand_event_scene_builds_cleanly_and_within_the_recorded_budget()
    {
        List<GanttEvent> events = SyntheticEvents(EventCount);

        // A warm-up pass so the measurement is not dominated by JIT of the builder
        // path on first use. The refresh budget is "after warm-up" and is read the
        // same way here.
        _ = SceneBuilder.TryBuild(Request(events));

        var stopwatch = Stopwatch.StartNew();
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(Request(events));
        // The budget is stated for "validate + build", so validation is inside the
        // timed region. Measuring only the build would report a number the
        // architecture never claimed, and would understate a validation regression.
        SceneValidationReport? report = null;
        if (outcome.Result is { } built)
        {
            report = SceneValidator.Validate(built.Scene);
        }

        stopwatch.Stop();

        Assert.True(outcome.Succeeded, "Scene build refused at scale: " + outcome.Refusal);
        Assert.NotNull(report);
        Assert.True(report!.IsClean, "Findings at scale: " + string.Join("; ", report.Findings));

        double elapsedMs = stopwatch.Elapsed.TotalMilliseconds;
        _output.WriteLine(string.Create(
            CultureInfo.InvariantCulture,
            $"R3.12 benchmark: {EventCount} events, validate+build {elapsedMs:F1} ms, {outcome.Result!.Scene.Primitives.Count} primitives, budget {BudgetMs} ms on the reference machine."));

        // The judgement is recorded, not asserted, so a slower machine does not fail
        // the gate. The verdict line is printed for whoever reads the test log.
        _output.WriteLine(elapsedMs < BudgetMs
            ? "R3.12 benchmark verdict: under the 250 ms Core budget on this machine."
            : "R3.12 benchmark verdict: OVER the 250 ms Core budget on this machine; not the reference machine, so this is reported rather than failed.");
    }

    [Fact]
    public void A_thousand_event_scene_is_deterministic()
    {
        List<GanttEvent> events = SyntheticEvents(EventCount);

        string baseline = SceneSnapshot.Serialize(SceneBuilder.TryBuild(Request(events)).Result!.Scene);
        string shuffled = SceneSnapshot.Serialize(
            SceneBuilder.TryBuild(Request([.. events.OrderByDescending(@event => @event.Id.Value)])).Result!.Scene);

        Assert.Equal(baseline, shuffled);
    }

    [Fact]
    public void The_benchmark_geometry_and_dates_are_deterministic()
    {
        // The row *dates and types* are deterministic, but the row identities are
        // not: GanttRowId.New() mints a fresh id per call by design, and a
        // benchmark that asserted stable ids would be asserting the opposite of the
        // contract. The determinism that matters is the scene's, proven above.
        List<GanttEvent> first = SyntheticEvents(32);
        List<GanttEvent> second = SyntheticEvents(32);

        Assert.Equal(
            first.Select(@event => (@event.RowNumber, @event.Type, @event.Start, @event.Finish)),
            second.Select(@event => (@event.RowNumber, @event.Type, @event.Start, @event.Finish)));
    }

    /// <summary>
    /// Builds the deterministic synthetic event set. A fixed arithmetic sequence is
    /// used rather than <see cref="Random"/> so the input is identical on every
    /// machine and every run; a seeded RNG would also work but ties the benchmark to
    /// the runtime's algorithm staying fixed.
    /// </summary>
    /// <param name="count">The number of events to generate.</param>
    /// <returns>The synthetic events in a deterministic order.</returns>
    private static List<GanttEvent> SyntheticEvents(int count)
    {
        // The plot spans the whole of 2026 so a thousand events of varying length all
        // land inside it, and each event gets its own lane so the benchmark exercises
        // real lane geometry rather than one very tall lane.
        var plotStart = new DateOnly(2026, 1, 5);
        var plotFinish = new DateOnly(2026, 12, 31);
        int totalDays = plotFinish.DayNumber - plotStart.DayNumber + 1;

        List<GanttEvent> events = new(count);
        for (var i = 0; i < count; i++)
        {
            var offset = (i * 7) % Math.Max(1, totalDays - 30);
            var start = plotStart.AddDays(offset);
            var length = 1 + ((i * 3) % 20);
            var finish = start.AddDays(Math.Min(length, totalDays - offset - 1));

            events.Add(new GanttEvent(
                i + 1,
                GanttRowId.New(),
                GanttRowId.New(),
                0,
                GanttEntityType.AsPlannedActivity,
                $"Synthetic activity {i}",
                start,
                finish,
                null,
                "AsPlannedActivity",
                null,
                null,
                null,
                true,
                null));
        }

        return events;
    }

    private static SceneBuildRequest Request(List<GanttEvent> events) =>
        new()
        {
            Events = events,
            Registry = ReferenceSceneBuilder.StyleRegistry,
            Grid = PanelCellGrid.TryCreate(
                [new PanelColumn("Id", 80), new PanelColumn("Description", 180)],
                [.. Enumerable.Repeat(10.0, events.Count)],
                10,
                ["Id", "Description"]).Grid,
            // A thousand lanes need a plot tall enough to hold them, so the benchmark
            // geometry is scaled rather than reusing the small fixture geometry.
            PlotBounds = new RectD(260, 110, 3_000, 30_000),
            Metrics = new FakeTextMetrics(static _ => 4.0, 10.0),
            LaneMetrics = new LaneLayoutMetrics(18, 3, 3, 2, 18, 9),
            FrameTheme = ReferenceSceneBuilder.FrameTheme,
            PlotStart = new DateOnly(2026, 1, 5),
            PlotFinish = new DateOnly(2026, 12, 31),
            Scale = GanttTimeScale.Month,
            PeriodLabelFormat = GanttPeriodLabelFormat.MMM,
            GridLinePt = 0.5,
            MajorBoundaryPt = 1,
            MilestoneSizePt = 8,
            CriticalLinePt = 1,
            TitleBandHeightPt = 14,
            YearBandHeightPt = 16,
            PeriodBandHeightPt = 20,
            // Label planning is part of the build the 250 ms budget is stated
            // against, so the benchmark must actually reach it: a null LabelStyle
            // makes BuildLabels return before planning anything, and the measured
            // figure would then cover a scene with no labels at all -- a thousand
            // descriptions and two thousand date labels unplanned. Zero/zero metrics
            // would be a second way to plan nothing, so the gap and height are the
            // same non-zero values the golden fixture uses.
            LabelGapPt = 2,
            LabelHeightPt = 8,
            LabelStyle = new SceneStyle("DefaultText", fillColour: ColourHex.Parse("#000000")),
        };
}
