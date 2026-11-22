namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R7.5 D2: the measured transient filter, the explicit attempt cap and the diagnostics budget.
/// Each member has a positive test; the documented candidates are the work item's recorded
/// codes, and the live busy-host body is the measurement that confirms or corrects them.
/// </summary>
public sealed class PowerPointTransientRetryPolicyTests
{
    private const int _unrecognised = unchecked((int)0x80004005);

    [Fact]
    public void The_documented_candidates_are_the_three_rpc_rejections()
    {
        PowerPointTransientRetryPolicy policy = PowerPointTransientRetryPolicy.DocumentedCandidates;

        Assert.True(policy.IsTransient(PowerPointTransientRetryPolicy.RpcCallRejected));
        Assert.True(policy.IsTransient(PowerPointTransientRetryPolicy.RpcServerCallRetryLater));
        Assert.True(policy.IsTransient(PowerPointTransientRetryPolicy.RpcServerCallRejected));
        Assert.Equal(PowerPointTransientRetryPolicy.DefaultMaxAttempts, policy.MaxAttempts);
    }

    [Fact]
    public void An_unrecognised_hresult_is_not_transient() =>
        Assert.False(PowerPointTransientRetryPolicy.DocumentedCandidates.IsTransient(_unrecognised));

    [Fact]
    public void A_missing_member_hresult_is_not_transient() =>
        Assert.False(
            PowerPointTransientRetryPolicy.DocumentedCandidates.IsTransient(
                PowerPointTransferFailureMapper.DispMemberNotFound));

    [Fact]
    public void A_null_set_is_refused() =>
        _ = Assert.Throws<ArgumentNullException>(
            () => new PowerPointTransientRetryPolicy(null!, PowerPointTransientRetryPolicy.DefaultMaxAttempts));

    [Fact]
    public void A_zero_attempt_cap_is_refused()
    {
        ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(
            () => new PowerPointTransientRetryPolicy([PowerPointTransientRetryPolicy.RpcCallRejected], 0));

        Assert.Equal("maxAttempts", exception.ParamName);
    }

    [Fact]
    public void Describe_filter_orders_the_codes_and_reports_none_when_empty()
    {
        PowerPointTransientRetryPolicy ordered = new(
            [PowerPointTransientRetryPolicy.RpcServerCallRejected, PowerPointTransientRetryPolicy.RpcCallRejected],
            2);

        Assert.Equal("0x80010001,0x8001010B", ordered.DescribeFilter());
        Assert.Equal("none", new PowerPointTransientRetryPolicy([], 1).DescribeFilter());
    }
}
