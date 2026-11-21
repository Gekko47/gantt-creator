namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R7.1/R7.3: the typed adapter's lifecycle is pinned without a host — construction, the
/// unattached fact and paste reads, disposal and the disposed guards. The COM traversal
/// itself is what the live gate exercises; it is deliberately not faked here.
/// </summary>
public sealed class PowerPointTransferAdapterTests
{
    [Fact]
    public void A_null_ownership_is_refused() =>
        _ = Assert.Throws<ArgumentNullException>(() => new PowerPointTransferAdapter(null!));

    [Fact]
    public void Dispose_before_attach_releases_nothing_and_is_idempotent()
    {
        RecordingOwnership ownership = new();
        PowerPointTransferAdapter adapter = new(ownership);

        adapter.Dispose();
        adapter.Dispose();

        Assert.Empty(ownership.Released);
    }

    [Fact]
    public void Attaching_after_dispose_is_refused()
    {
        PowerPointTransferAdapter adapter = new(new PowerPointProxyOwnership());
        adapter.Dispose();

        _ = Assert.Throws<ObjectDisposedException>(() =>
        {
            _ = adapter.AttachOrStart();
        });
    }

    [Fact]
    public void Reading_facts_after_dispose_is_refused()
    {
        PowerPointTransferAdapter adapter = new(new PowerPointProxyOwnership());
        adapter.Dispose();

        _ = Assert.Throws<ObjectDisposedException>(() =>
        {
            _ = adapter.ReadPresentationFacts();
        });
    }

    [Fact]
    public void Reading_facts_before_attach_reports_no_presentation_without_a_host()
    {
        using PowerPointTransferAdapter adapter = new(new PowerPointProxyOwnership());

        PowerPointPresentationFacts facts = adapter.ReadPresentationFacts();

        Assert.True(facts.ReadSucceeded);
        Assert.False(facts.HasPresentation);
        Assert.Null(facts.HResult);
    }

    [Fact]
    public void Reading_the_target_observation_before_attach_reports_an_empty_typed_state()
    {
        using PowerPointTransferAdapter adapter = new(new PowerPointProxyOwnership());

        PowerPointTargetObservation observation = adapter.ReadTargetObservation();

        Assert.True(observation.ReadSucceeded);
        Assert.Null(observation.HResult);
        Assert.Equal(0, observation.PresentationCount);
        Assert.Equal(0, observation.ProtectedViewWindowCount);
        Assert.False(observation.ActiveProtectedViewWindowPresent);
        Assert.False(observation.ActivePresentationReadOnly);
        Assert.False(observation.ActiveWindowPresent);
        Assert.Null(observation.ActiveViewType);
        Assert.False(observation.SlideShowActive);
        Assert.Equal(0, observation.SlideCount);
        Assert.Null(observation.ActiveSlideIndex);
    }

    [Fact]
    public void Reading_the_target_observation_after_dispose_is_refused()
    {
        PowerPointTransferAdapter adapter = new(new PowerPointProxyOwnership());
        adapter.Dispose();

        _ = Assert.Throws<ObjectDisposedException>(() =>
        {
            _ = adapter.ReadTargetObservation();
        });
    }

    [Fact]
    public void Pasting_before_attach_reports_a_refused_observation_without_a_host()
    {
        using PowerPointTransferAdapter adapter = new(new PowerPointProxyOwnership());

        PowerPointPasteObservation observation = adapter.PasteSpecialShape();

        Assert.False(observation.PasteSucceeded);
        Assert.Null(observation.HResult);
        Assert.Equal(0, observation.Count);
        Assert.Equal(0, observation.GroupCount);
        Assert.Null(observation.GroupShapeType);
        Assert.Null(observation.GroupName);
        Assert.Null(observation.GroupId);
    }

    [Fact]
    public void Pasting_after_dispose_is_refused()
    {
        PowerPointTransferAdapter adapter = new(new PowerPointProxyOwnership());
        adapter.Dispose();

        _ = Assert.Throws<ObjectDisposedException>(() =>
        {
            _ = adapter.PasteSpecialShape();
        });
    }

    private sealed class RecordingOwnership : PowerPointProxyOwnership
    {
        internal List<object> Released { get; } = [];

        internal override int ReleaseCore(object proxy)
        {
            Released.Add(proxy);
            return 0;
        }
    }
}
