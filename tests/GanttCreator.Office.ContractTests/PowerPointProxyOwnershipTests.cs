namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R7.1 D2: every proxy the add-in owns is released exactly once, in reverse acquisition
/// order, through the tested ownership helper, and nothing it does not own is touched.
/// </summary>
public sealed class PowerPointProxyOwnershipTests
{
    [Fact]
    public void The_proxies_are_released_in_reverse_acquisition_order()
    {
        RecordingOwnership ownership = new();
        object first = new();
        object second = new();
        object third = new();

        _ = ownership.Acquire(first);
        _ = ownership.Acquire(second);
        _ = ownership.Acquire(third);

        ownership.ReleaseAll();

        Assert.Equal(new[] { third, second, first }, ownership.Released);
    }

    [Fact]
    public void Each_proxy_is_released_exactly_once()
    {
        RecordingOwnership ownership = new();
        _ = ownership.Acquire(new object());
        _ = ownership.Acquire(new object());

        ownership.ReleaseAll();
        ownership.ReleaseAll();

        Assert.Equal(2, ownership.Released.Count);
        Assert.Equal(0, ownership.Count);
    }

    [Fact]
    public void A_proxy_that_was_never_acquired_is_never_released()
    {
        RecordingOwnership ownership = new();
        object owned = new();
        object unowned = new();
        _ = ownership.Acquire(owned);

        ownership.ReleaseAll();

        _ = Assert.Single(ownership.Released);
        Assert.DoesNotContain(unowned, ownership.Released, ReferenceEqualityComparer.Instance);
        Assert.Contains(owned, ownership.Released, ReferenceEqualityComparer.Instance);
    }

    [Fact]
    public void A_refused_release_is_recorded_once_and_does_not_escape()
    {
        List<string> records = [];
        RefusingOwnership ownership = new(records.Add);
        _ = ownership.Acquire(new object());

        ownership.ReleaseAll();

        string record = Assert.Single(records);
        Assert.Contains("powerpoint release refused", record, StringComparison.Ordinal);
        Assert.Contains("0x80070057", record, StringComparison.Ordinal);
        Assert.Equal(0, ownership.Count);
    }

    [Fact]
    public void A_non_com_proxy_is_ignored_by_the_default_release()
    {
        PowerPointProxyOwnership ownership = new();
        _ = ownership.Acquire(new object());

        // A plain object is not a COM proxy; the default release must not call the
        // marshaller and must not throw. This is the case a contract test relies on.
        ownership.ReleaseAll();

        Assert.Equal(0, ownership.Count);
    }

    [Fact]
    public void Acquiring_null_is_refused()
    {
        PowerPointProxyOwnership ownership = new();

        _ = Assert.Throws<ArgumentNullException>(() => ownership.Acquire<object>(null!));
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

    private sealed class RefusingOwnership : PowerPointProxyOwnership
    {
        internal RefusingOwnership(Action<string>? technicalRecord)
            : base(technicalRecord)
        {
        }

        // The tracker's documented refusal path catches ArgumentException or COMException.
        // ArgumentException is used here because COMException is a reserved runtime type.
        internal override int ReleaseCore(object proxy) =>
            throw new ArgumentException("the host refused the release");
    }
}
