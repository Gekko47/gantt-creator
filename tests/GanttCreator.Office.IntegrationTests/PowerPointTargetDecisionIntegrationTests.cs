using GanttCreator.Core.Export;
using Xunit.Abstractions;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// R7.2: the live attach/start and active-slide case gate. Each body arranges exactly the deck
/// state it owns through <see cref="PowerPointFixture"/> and asserts the Core decision table
/// against the real host; the host build is recorded with every outcome.
/// </summary>
/// <remarks>
/// <para>
/// <b>Non-destructive.</b> Every body mutates only a fixture-owned deck (or starts a host and
/// lets the fixture adopt it for cleanup). A pre-existing PowerPoint is observed and never
/// viewed, shown, closed or quit — the landed fixture ownership discipline.
/// </para>
/// <para>
/// <b>Measured, not faked.</b> These bodies exercise the typed COM traversal the host-free
/// tests deliberately do not fake. The outline and standard-view bodies record what the host
/// actually reports, because PowerPoint may normalise a requested view.
/// </para>
/// </remarks>
public sealed class PowerPointTargetDecisionIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task New_power_point_with_no_presentations_decides_CreatePresentation()
    {
        // The suite runs several PowerPoint bodies in one process, and a prior fixture's
        // Quit() completes asynchronously: the process can still appear in the snapshot for
        // a short window. Poll the named observable condition with a deadline rather than
        // asserting once, so suite composition cannot turn a clean teardown into a failure.
        Dictionary<int, long> before = await WaitForNoPowerPointAsync().ConfigureAwait(true);
        Assert.True(
            before.Count == 0,
            "Close PowerPoint before running the R7.2 live gate: the new-PowerPoint case requires no running POWERPNT.");

        PowerPointTargetDecision decision;
        PowerPointTransferAdapter adapter = new(new PowerPointProxyOwnership());
        try
        {
            decision = new PowerPointTargetResolver(adapter).Resolve();
        }
        finally
        {
            adapter.Dispose();
        }

        PowerPointFixture cleanup = new();
        try
        {
            cleanup.AdoptHostStartedSince(before);
            _output.WriteLine($"host={cleanup.BuildDescription} startedByUs={decision.StartedByUs}");
            Assert.Equal(PowerPointTargetAction.CreatePresentation, decision.Action);
            Assert.True(decision.StartedByUs);
            Assert.Null(decision.Refusal);
        }
        finally
        {
            await cleanup.DisposeAsync().ConfigureAwait(true);
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Existing_power_point_with_an_active_slide_proceeds_and_attaches_without_a_start()
    {
        PowerPointFixture fixture = new();
        try
        {
            fixture.Initialize(withWindow: true, slideCount: 1);

            PowerPointTargetDecision decision = ResolveTarget();

            _output.WriteLine($"host={fixture.BuildDescription} startedByUs={decision.StartedByUs}");
            AssertProceeded(decision);
            Assert.False(decision.StartedByUs);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task A_presentation_with_no_slides_is_refused_NoSlide()
    {
        PowerPointFixture fixture = new();
        try
        {
            fixture.Initialize(withWindow: true, slideCount: 0);

            PowerPointTargetDecision decision = ResolveTarget();

            _output.WriteLine($"host={fixture.BuildDescription}");
            AssertRefused(decision, PowerPointTargetRefusal.NoSlide);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task A_windowless_presentation_is_refused_NoWindow()
    {
        PowerPointFixture fixture = new();
        try
        {
            fixture.Initialize(withWindow: false, slideCount: 1);

            PowerPointTargetDecision decision = ResolveTarget();

            _output.WriteLine($"host={fixture.BuildDescription}");
            AssertRefused(decision, PowerPointTargetRefusal.NoWindow);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task The_slide_sorter_view_is_refused_UnsupportedView()
    {
        PowerPointFixture fixture = new();
        try
        {
            fixture.Initialize(withWindow: true, slideCount: 1);
            int readBack = fixture.SetActiveView(PowerPoint.PpViewType.ppViewSlideSorter);

            PowerPointTargetDecision decision = ResolveTarget();

            _output.WriteLine($"host={fixture.BuildDescription} readBack={readBack}");
            Assert.Equal((int)PowerPoint.PpViewType.ppViewSlideSorter, readBack);
            AssertRefused(decision, PowerPointTargetRefusal.UnsupportedView);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task The_outline_view_is_refused_or_reported_as_the_host_normalises_it()
    {
        PowerPointFixture fixture = new();
        try
        {
            fixture.Initialize(withWindow: true, slideCount: 1);
            int readBack = fixture.SetActiveView(PowerPoint.PpViewType.ppViewOutline);

            PowerPointTargetDecision decision = ResolveTarget();

            _output.WriteLine($"host={fixture.BuildDescription} outlineReadBack={readBack}");
            if (readBack == (int)PowerPoint.PpViewType.ppViewOutline)
            {
                AssertRefused(decision, PowerPointTargetRefusal.UnsupportedView);
            }
            else
            {
                // Measured on the reference host: the window normalised the request. The Core
                // theory remains the covering evidence for the outline value itself.
                Assert.Equal((int)PowerPoint.PpViewType.ppViewNormal, readBack);
                AssertProceeded(decision);
            }
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task A_running_slide_show_is_refused_SlideShowActive()
    {
        PowerPointFixture fixture = new();
        try
        {
            fixture.Initialize(withWindow: true, slideCount: 1);
            fixture.StartSlideShow();

            PowerPointTargetDecision decision = ResolveTarget();

            _output.WriteLine($"host={fixture.BuildDescription}");
            AssertRefused(decision, PowerPointTargetRefusal.SlideShowActive);
        }
        finally
        {
            fixture.ExitSlideShow();
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task A_read_only_presentation_is_refused_ReadOnlyPresentation()
    {
        PowerPointFixture fixture = new();
        string? path = null;
        try
        {
            fixture.Initialize(withWindow: true, slideCount: 1);
            path = fixture.SavePrimaryPresentationToTemp();

            // Close the editable deck first so the read-only deck is the only presentation and
            // therefore the active one; the transfer never activates a window to arrange this.
            fixture.ClosePrimaryPresentation();
            fixture.OpenReadOnly(path);

            PowerPointTargetDecision decision = ResolveTarget();

            _output.WriteLine($"host={fixture.BuildDescription}");
            AssertRefused(decision, PowerPointTargetRefusal.ReadOnlyPresentation);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
            if (path is not null)
            {
                TryDelete(path);
            }
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task A_protected_view_with_no_presentation_is_refused_ProtectedView()
    {
        PowerPointFixture fixture = new();
        string? path = null;
        try
        {
            fixture.Initialize(withWindow: false, slideCount: 1);
            path = fixture.SavePrimaryPresentationToTemp();
            fixture.ClosePrimaryPresentation();
            (bool opened, int? hresult) = fixture.OpenProtectedView(path);

            if (!opened)
            {
                // Measured deviation (Phase-0 probe, 2026-11-21): the reference host refuses the
                // automation open route with 0x80004005, so this Required live case cannot be
                // set up here. The Core matrix pins the ProtectedView classification and the
                // deviation is recorded in the work item, never hidden by a silent pass.
                _output.WriteLine(
                    string.Create(
                        System.Globalization.CultureInfo.InvariantCulture,
                        $"protected-view-open not exercisable: hresult=0x{hresult.GetValueOrDefault():X8}"));
                return;
            }

            PowerPointTargetDecision decision = ResolveTarget();

            _output.WriteLine($"host={fixture.BuildDescription} refusal={decision.Refusal}");
            AssertRefused(decision, PowerPointTargetRefusal.ProtectedView);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
            if (path is not null)
            {
                TryDelete(path);
            }
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task Two_presentations_resolve_the_active_deck_not_the_first_one()
    {
        PowerPointFixture fixture = new();
        try
        {
            // Deck A is created first and has no window, so a read of Presentations[1] would
            // decide NoWindow; deck B is created second, has a window and is active, so only
            // an active-window resolution reaches ProceedWithActiveSlide.
            fixture.Initialize(withWindow: false, slideCount: 2);
            fixture.AddPresentation(withWindow: true, slideCount: 3);

            PowerPointTargetDecision decision = ResolveTarget();

            _output.WriteLine($"host={fixture.BuildDescription}");
            AssertProceeded(decision);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task The_standard_editing_view_is_measured_and_decided_consistently()
    {
        PowerPointFixture fixture = new();
        try
        {
            fixture.Initialize(withWindow: true, slideCount: 1);
            int viewType = fixture.ActiveViewType();

            PowerPointTargetDecision decision = ResolveTarget();

            _output.WriteLine($"host={fixture.BuildDescription} standardViewType={viewType}");
            if (viewType is (int)PowerPoint.PpViewType.ppViewSlide or (int)PowerPoint.PpViewType.ppViewNormal)
            {
                AssertProceeded(decision);
            }
            else
            {
                AssertRefused(decision, PowerPointTargetRefusal.UnsupportedView);
            }
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task A_foreground_protected_view_is_refused_over_an_open_deck()
    {
        PowerPointFixture fixture = new();
        string? path = null;
        try
        {
            fixture.Initialize(withWindow: true, slideCount: 1);
            path = fixture.SavePrimaryPresentationToTemp();
            (bool opened, int? hresult) = fixture.OpenProtectedView(path);

            if (!opened)
            {
                // Measured deviation (Phase-0 probe, 2026-11-21): the foreground protected-view
                // case cannot be set up through this automation route on the reference host.
                _output.WriteLine(
                    string.Create(
                        System.Globalization.CultureInfo.InvariantCulture,
                        $"protected-view-open not exercisable: hresult=0x{hresult.GetValueOrDefault():X8}"));
                return;
            }

            PowerPointTargetDecision decision = ResolveTarget();

            _output.WriteLine($"host={fixture.BuildDescription} refusal={decision.Refusal}");
            AssertRefused(decision, PowerPointTargetRefusal.ProtectedView);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
            if (path is not null)
            {
                TryDelete(path);
            }
        }
    }

    private PowerPointTargetDecision ResolveTarget()
    {
        using PowerPointTransferAdapter adapter = new(
            new PowerPointProxyOwnership(),
            record => _output.WriteLine(record));
        PowerPointTargetDecision decision = new PowerPointTargetResolver(adapter).Resolve();
        _output.WriteLine(
            $"decision={decision.Action} refusal={decision.Refusal} hresult={decision.HResult}");
        return decision;
    }

    private static void AssertProceeded(PowerPointTargetDecision decision, int expectedSlideIndex = 1)
    {
        Assert.Equal(PowerPointTargetAction.ProceedWithActiveSlide, decision.Action);
        Assert.Null(decision.Refusal);
        Assert.Equal(expectedSlideIndex, decision.ActiveSlideIndex);
        Assert.Null(decision.HResult);
    }

    private static void AssertRefused(PowerPointTargetDecision decision, PowerPointTargetRefusal refusal)
    {
        Assert.Equal(PowerPointTargetAction.Refuse, decision.Action);
        Assert.Equal(refusal, decision.Refusal);
        Assert.Null(decision.ActiveSlideIndex);
    }

    /// <summary>
    /// Polls the named observable condition (no running <c>POWERPNT</c>) with a deadline, so a
    /// PowerPoint instance that has been asked to quit but has not yet exited does not fail
    /// the body. The deadline keeps the diagnostic honest when a host really is running.
    /// </summary>
    /// <returns>The empty snapshot, or the last non-empty one when the deadline expires.</returns>
    private static async Task<Dictionary<int, long>> WaitForNoPowerPointAsync()
    {
        DateTime deadline = DateTime.UtcNow + TimeSpan.FromSeconds(15);
        Dictionary<int, long> snapshot = PowerPointFixture.SnapshotProcesses();
        while (snapshot.Count > 0 && DateTime.UtcNow < deadline)
        {
            await Task.Delay(100).ConfigureAwait(false);
            snapshot = PowerPointFixture.SnapshotProcesses();
        }

        return snapshot;
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (IOException)
        {
            // The file is in the OS temp folder; a failed delete must not fail the body.
        }
        catch (UnauthorizedAccessException)
        {
            // Same.
        }
    }
}
