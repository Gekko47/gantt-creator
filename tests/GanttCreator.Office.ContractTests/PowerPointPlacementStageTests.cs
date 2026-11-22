using GanttCreator.Core;
using GanttCreator.Core.Export;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R7.4: the placement stage's host-read mapping over a fake port. The fit runs in Core, the
/// placement is consumed once, the port's typed observation is mapped to the stage refusal, and
/// the read-back deltas are verified against the repository tolerance — all without a host.
/// </summary>
/// <remarks>
/// The adapter's COM traversal is what the live gate exercises — this file deliberately owns no
/// host, following <see cref="PowerPointPasteStageTests"/>. The exception path is covered
/// host-free through the adapter's typed refused observation.
/// </remarks>
public sealed class PowerPointPlacementStageTests
{
    /// <summary>A 4:3 slide width in points.</summary>
    private const double _slideWidthPt = 720;

    /// <summary>A 4:3 slide height in points.</summary>
    private const double _slideHeightPt = 540;

    /// <summary>The placeholder policy inset used by these tests (PD-6 makes it a parameter).</summary>
    private const double _insetPt = 4;

    /// <summary>The forwarded pasted group identifier (R7.3).</summary>
    private const int _groupId = 42;

    /// <summary>A generic COM failure code (E_FAIL).</summary>
    private const int _genericFailureHResult = unchecked((int)0x80004005);

    /// <summary>The resolved zero-origin export bounds used by the fit.</summary>
    private static readonly RectD _groupBounds = new(0, 0, 800, 200);

    [Fact]
    public void A_healthy_placement_is_fitted_applied_and_verified()
    {
        FakePlacementPort port = new(HealthySize(), ObservedFromPlacement);
        List<string> records = [];
        PowerPointPlacementStage stage = new(port, records.Add);

        PowerPointPlacementResult result = stage.Place(_groupId, _groupBounds, _insetPt);

        Assert.True(result.Succeeded, $"Refused: {result.Refusal} {result.FitRefusal}");
        Assert.Null(result.Refusal);
        Assert.Null(result.FitRefusal);
        Assert.NotNull(result.Placement);
        Assert.NotNull(result.Deltas);
        Assert.True(result.Deltas.MaxPt <= HostOwnedGeometry.GeometryMatchTolerancePt);
        Assert.Single(port.Calls);
        Assert.Equal(_groupId, port.Calls[0].GroupId);
        Assert.Empty(records);
    }

    [Fact]
    public void An_invalid_slide_size_is_a_fit_refusal_and_the_port_is_never_called()
    {
        FakePlacementPort port = new(new PowerPointSlideSizeFacts(true, null, 0, _slideHeightPt), ObservedFromPlacement);
        List<string> records = [];
        PowerPointPlacementStage stage = new(port, records.Add);

        PowerPointPlacementResult result = stage.Place(_groupId, _groupBounds, _insetPt);

        Assert.False(result.Succeeded);
        Assert.Equal(SlideFitRefusal.InvalidSlideSize, result.FitRefusal);
        Assert.Null(result.Refusal);
        Assert.Empty(port.Calls);
        Assert.Single(records);
    }

    [Fact]
    public void An_unavailable_slide_size_is_refused_with_the_hresult()
    {
        FakePlacementPort port = new(
            new PowerPointSlideSizeFacts(false, _genericFailureHResult, 0, 0),
            ObservedFromPlacement);
        List<string> records = [];
        PowerPointPlacementStage stage = new(port, records.Add);

        PowerPointPlacementResult result = stage.Place(_groupId, _groupBounds, _insetPt);

        Assert.False(result.Succeeded);
        Assert.Equal(PowerPointPlacementRefusal.SlideSizeUnavailable, result.Refusal);
        Assert.Equal(_genericFailureHResult, result.HResult);
        Assert.Empty(port.Calls);
        Assert.Single(records);
    }

    [Fact]
    public void A_second_application_of_the_same_placement_is_refused()
    {
        FakePlacementPort port = new(HealthySize(), ObservedFromPlacement);
        List<string> records = [];
        PowerPointPlacementStage stage = new(port, records.Add);
        SlidePlacement placement = Assert.IsType<SlidePlacement>(
            SlideFit.Fit(_groupBounds, _slideWidthPt, _slideHeightPt, _insetPt).Placement);

        PowerPointPlacementResult first = stage.Apply(_groupId, placement);
        PowerPointPlacementResult second = stage.Apply(_groupId, placement);

        Assert.True(first.Succeeded);
        Assert.False(second.Succeeded);
        Assert.Equal(PowerPointPlacementRefusal.AlreadyApplied, second.Refusal);
        Assert.Single(port.Calls);
        Assert.Single(records);
    }

    [Fact]
    public void A_missing_group_is_refused()
    {
        FakePlacementPort port = new(
            HealthySize(),
            _ => new PowerPointPlacementObservation(PowerPointPlacementApplyOutcome.GroupMissing, null, 0, 0, 0, 0));
        List<string> records = [];
        PowerPointPlacementStage stage = new(port, records.Add);

        PowerPointPlacementResult result = stage.Place(_groupId, _groupBounds, _insetPt);

        Assert.False(result.Succeeded);
        Assert.Equal(PowerPointPlacementRefusal.GroupMissing, result.Refusal);
        Assert.Null(result.HResult);
        Assert.Single(records);
    }

    [Fact]
    public void A_host_rejection_carries_the_hresult()
    {
        FakePlacementPort port = new(
            HealthySize(),
            _ => new PowerPointPlacementObservation(
                PowerPointPlacementApplyOutcome.HostRejected,
                _genericFailureHResult,
                0,
                0,
                0,
                0));
        List<string> records = [];
        PowerPointPlacementStage stage = new(port, records.Add);

        PowerPointPlacementResult result = stage.Place(_groupId, _groupBounds, _insetPt);

        Assert.False(result.Succeeded);
        Assert.Equal(PowerPointPlacementRefusal.HostRejected, result.Refusal);
        Assert.Equal(_genericFailureHResult, result.HResult);
        Assert.Single(records);
        Assert.Contains("refused", records[0], StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void A_missing_slide_is_refused()
    {
        FakePlacementPort port = new(
            HealthySize(),
            _ => new PowerPointPlacementObservation(PowerPointPlacementApplyOutcome.NoSlide, null, 0, 0, 0, 0));
        List<string> records = [];
        PowerPointPlacementStage stage = new(port, records.Add);

        PowerPointPlacementResult result = stage.Place(_groupId, _groupBounds, _insetPt);

        Assert.False(result.Succeeded);
        Assert.Equal(PowerPointPlacementRefusal.NoSlide, result.Refusal);
        Assert.Null(result.HResult);
        Assert.Single(records);
    }

    [Fact]
    public void A_read_back_mismatch_is_refused_and_the_deltas_are_recorded()
    {
        FakePlacementPort port = new(
            HealthySize(),
            placement => new PowerPointPlacementObservation(
                PowerPointPlacementApplyOutcome.Applied,
                null,
                placement.Left,
                placement.Top,
                placement.Width + 5,
                placement.Height));
        List<string> records = [];
        PowerPointPlacementStage stage = new(port, records.Add);

        PowerPointPlacementResult result = stage.Place(_groupId, _groupBounds, _insetPt);

        Assert.False(result.Succeeded);
        Assert.Equal(PowerPointPlacementRefusal.ReadBackMismatch, result.Refusal);
        PowerPointPlacementDeltas deltas = Assert.IsType<PowerPointPlacementDeltas>(result.Deltas);
        Assert.True(Math.Abs(deltas.WidthPt - 5) <= HostOwnedGeometry.GeometryMatchTolerancePt);
        Assert.Single(records);
    }

    [Fact]
    public void A_null_port_is_refused() =>
        _ = Assert.Throws<ArgumentNullException>(() => new PowerPointPlacementStage(null!));

    [Fact]
    public void A_null_placement_is_refused()
    {
        PowerPointPlacementStage stage = new(new FakePlacementPort(HealthySize(), ObservedFromPlacement));

        ArgumentNullException exception =
            Assert.Throws<ArgumentNullException>(() => stage.Apply(_groupId, null!));

        Assert.Equal("placement", exception.ParamName);
    }

    private static PowerPointSlideSizeFacts HealthySize() =>
        new(true, null, _slideWidthPt, _slideHeightPt);

    private static PowerPointPlacementObservation ObservedFromPlacement(SlidePlacement placement) =>
        new(
            PowerPointPlacementApplyOutcome.Applied,
            null,
            placement.Left,
            placement.Top,
            placement.Width,
            placement.Height);

    private sealed class FakePlacementPort(
        PowerPointSlideSizeFacts size,
        Func<SlidePlacement, PowerPointPlacementObservation> apply)
        : IPowerPointPlacementPort
    {
        public List<(int GroupId, SlidePlacement Placement)> Calls { get; } = [];

        public PowerPointSlideSizeFacts ReadSlideSize() => size;

        public PowerPointPlacementObservation ApplyPlacement(int groupId, SlidePlacement placement)
        {
            Calls.Add((groupId, placement));
            return apply(placement);
        }
    }
}
