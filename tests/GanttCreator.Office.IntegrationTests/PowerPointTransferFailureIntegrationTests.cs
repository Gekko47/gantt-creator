using GanttCreator.Core.Export;
using Xunit.Abstractions;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// R7.5 Office gate: each safe supported failure the reference host can arrange is driven
/// through the production adapter and target resolver, and the R7.5 mapper is asserted to
/// report the matching taxonomy member.
/// </summary>
/// <remarks>
/// <para>
/// <b>Non-destructive.</b> Every body mutates only a fixture-owned deck. A pre-existing
/// PowerPoint is observed and never viewed, shown, closed or quit — the landed fixture
/// ownership discipline.
/// </para>
/// <para>
/// <b>The busy-host transient measurement is deferred, not faked.</b> The R7.5 Step 0 requires
/// the transient set to be measured on the reference host through a busy PowerPoint. No safe
/// modal-hold mechanism is established for the reference host without risking a hung gate, so
/// the work item's own stop condition applies ("A failure mode cannot be reproduced safely ->
/// record it; the human decides whether it is supported-tested or documented-only"). The record
/// is in the work item; the retry mechanism is proven host-free by
/// <c>PowerPointRetryLoopTests</c>, and the policy ships the documented candidates pending that
/// measurement.
/// </para>
/// </remarks>
public sealed class PowerPointTransferFailureIntegrationTests(ITestOutputHelper output)
{
    private readonly ITestOutputHelper _output = output;

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task A_read_only_presentation_maps_to_PresentationReadOnly()
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

            PowerPointTransferFailure failure = ResolveFailure(fixture);

            Assert.Equal(PowerPointTransferFailureKind.PresentationReadOnly, failure.Kind);
            Assert.False(string.IsNullOrWhiteSpace(failure.Message));
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
            TryDelete(path);
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task The_slide_sorter_view_maps_to_UnsupportedView()
    {
        PowerPointFixture fixture = new();
        try
        {
            fixture.Initialize(withWindow: true, slideCount: 1);
            int readBack = fixture.SetActiveView(PowerPoint.PpViewType.ppViewSlideSorter);

            PowerPointTransferFailure failure = ResolveFailure(fixture);

            Assert.Equal((int)PowerPoint.PpViewType.ppViewSlideSorter, readBack);
            Assert.Equal(PowerPointTransferFailureKind.UnsupportedView, failure.Kind);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task A_presentation_with_no_slides_maps_to_NoSlide()
    {
        PowerPointFixture fixture = new();
        try
        {
            fixture.Initialize(withWindow: true, slideCount: 0);

            PowerPointTransferFailure failure = ResolveFailure(fixture);

            Assert.Equal(PowerPointTransferFailureKind.NoSlide, failure.Kind);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    [Trait("Category", "OfficeIntegration")]
    [Fact]
    public async Task A_windowless_presentation_maps_to_NoWindow()
    {
        PowerPointFixture fixture = new();
        try
        {
            fixture.Initialize(withWindow: false, slideCount: 1);

            PowerPointTransferFailure failure = ResolveFailure(fixture);

            Assert.Equal(PowerPointTransferFailureKind.NoWindow, failure.Kind);
        }
        finally
        {
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }

    private PowerPointTransferFailure ResolveFailure(PowerPointFixture fixture)
    {
        using PowerPointTransferAdapter adapter = new(new PowerPointProxyOwnership(), _output.WriteLine);
        PowerPointTargetDecision decision = new PowerPointTargetResolver(adapter).Resolve();
        _output.WriteLine(
            $"host={fixture.BuildDescription} decision={decision.Action} refusal={decision.Refusal}");

        return Assert.IsType<PowerPointTransferFailure>(
            PowerPointTransferFailureMapper.MapTargetDecision(decision));
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
            // The file is in the OS temp folder; a failed delete must not fail the body.
        }
        catch (UnauthorizedAccessException)
        {
            // Same.
        }
    }
}
