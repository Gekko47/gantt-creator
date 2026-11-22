namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R7.1 D2: <c>StartedByUs</c> is measured from the <c>POWERPNT</c> process set, never
/// assumed. The derivation and host resolution are primitive helpers so both arms are
/// proven without a host.
/// </summary>
public sealed class PowerPointProcessSnapshotTests
{
    [Fact]
    public void A_call_that_created_the_host_is_owned() =>
        Assert.True(PowerPointProcessSnapshot.DeriveStartedByUs([], [42]));

    [Fact]
    public void A_preexisting_host_is_not_owned() =>
        Assert.False(PowerPointProcessSnapshot.DeriveStartedByUs([42], [42]));

    [Fact]
    public void A_second_process_does_not_make_an_existing_host_owned() =>
        Assert.False(PowerPointProcessSnapshot.DeriveStartedByUs([42], [42, 43]));

    [Fact]
    public void No_process_on_either_side_is_not_owned() =>
        Assert.False(PowerPointProcessSnapshot.DeriveStartedByUs([], []));

    [Fact]
    public void The_process_created_by_the_call_is_the_host() =>
        Assert.Equal(43, PowerPointProcessSnapshot.ResolveHostProcessId([42], [42, 43]));

    [Fact]
    public void An_existing_process_is_the_host_when_none_was_created() =>
        Assert.Equal(42, PowerPointProcessSnapshot.ResolveHostProcessId([42], [42]));

    [Fact]
    public void No_process_resolves_to_zero() =>
        Assert.Equal(0, PowerPointProcessSnapshot.ResolveHostProcessId([], []));

    [Fact]
    public void An_unknown_process_has_no_build_description() =>
        Assert.Null(PowerPointProcessSnapshot.DescribeBuild(0));
}
