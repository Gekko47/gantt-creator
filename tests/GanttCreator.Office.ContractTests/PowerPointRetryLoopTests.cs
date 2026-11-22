namespace GanttCreator.Office.ContractTests;

/// <summary>
/// R7.5 D2/PD-3: the bounded retry loop — the transient filter, the attempt cap and the
/// mutation boundary. Every branch has a positive test, including the non-transient case (D2's
/// "only measured transient rejections retry") and the post-paste boundary (PD-3).
/// </summary>
public sealed class PowerPointRetryLoopTests
{
    private const int _transient = unchecked((int)0x80010001);
    private const int _unrecognised = unchecked((int)0x80004005);

    [Fact]
    public void A_transient_failure_retries_until_it_succeeds()
    {
        int calls = 0;

        string result = PowerPointRetryLoop.Execute(
            Policy(3),
            retryAfterMutation: true,
            attempt =>
            {
                calls++;
                return attempt < 3
                    ? new PowerPointRetryAttempt<string>("refused", _transient, MutationCommitted: false)
                    : new PowerPointRetryAttempt<string>("ok", null, MutationCommitted: false);
            },
            technicalRecord: null,
            "probe");

        Assert.Equal("ok", result);
        Assert.Equal(3, calls);
    }

    [Fact]
    public void A_non_transient_failure_is_never_retried()
    {
        int calls = 0;

        string result = PowerPointRetryLoop.Execute(
            Policy(3),
            retryAfterMutation: true,
            _ =>
            {
                calls++;
                return new PowerPointRetryAttempt<string>("refused", _unrecognised, MutationCommitted: false);
            },
            technicalRecord: null,
            "probe");

        Assert.Equal("refused", result);
        Assert.Equal(1, calls);
    }

    [Fact]
    public void The_attempt_cap_is_respected()
    {
        int calls = 0;

        string result = PowerPointRetryLoop.Execute(
            Policy(3),
            retryAfterMutation: true,
            _ =>
            {
                calls++;
                return new PowerPointRetryAttempt<string>("refused", _transient, MutationCommitted: false);
            },
            technicalRecord: null,
            "probe");

        Assert.Equal("refused", result);
        Assert.Equal(3, calls);
    }

    [Fact]
    public void A_post_mutation_transient_failure_is_not_retried_for_a_mutation_bounded_operation()
    {
        int calls = 0;

        string result = PowerPointRetryLoop.Execute(
            Policy(3),
            retryAfterMutation: false,
            _ =>
            {
                calls++;
                return new PowerPointRetryAttempt<string>("refused", _transient, MutationCommitted: true);
            },
            technicalRecord: null,
            "paste");

        Assert.Equal("refused", result);
        Assert.Equal(1, calls);
    }

    [Fact]
    public void A_post_mutation_transient_failure_is_retried_for_an_idempotent_operation()
    {
        int calls = 0;

        string result = PowerPointRetryLoop.Execute(
            Policy(3),
            retryAfterMutation: true,
            attempt =>
            {
                calls++;
                return attempt < 2
                    ? new PowerPointRetryAttempt<string>("refused", _transient, MutationCommitted: true)
                    : new PowerPointRetryAttempt<string>("ok", null, MutationCommitted: true);
            },
            technicalRecord: null,
            "placement");

        Assert.Equal("ok", result);
        Assert.Equal(2, calls);
    }

    [Fact]
    public void An_exhausted_retry_writes_the_filter_attempts_and_elapsed_budget()
    {
        List<string> records = [];

        _ = PowerPointRetryLoop.Execute(
            Policy(2),
            retryAfterMutation: false,
            _ => new PowerPointRetryAttempt<string>("refused", _transient, MutationCommitted: false),
            records.Add,
            "paste");

        string stop = Assert.Single(records, record => record.Contains("retry=stop", StringComparison.Ordinal));
        Assert.Contains("attempts=2", stop, StringComparison.Ordinal);
        Assert.Contains("maxattempts=2", stop, StringComparison.Ordinal);
        Assert.Contains("filter=0x80010001", stop, StringComparison.Ordinal);
        Assert.Contains("elapsedms=", stop, StringComparison.Ordinal);
    }

    [Fact]
    public void A_retry_continue_record_is_written_for_each_retry()
    {
        List<string> records = [];

        _ = PowerPointRetryLoop.Execute(
            Policy(3),
            retryAfterMutation: true,
            attempt => attempt < 3
                ? new PowerPointRetryAttempt<string>("refused", _transient, MutationCommitted: false)
                : new PowerPointRetryAttempt<string>("ok", null, MutationCommitted: false),
            records.Add,
            "probe");

        Assert.Equal(
            2,
            records.Count(record => record.Contains("retry=continue", StringComparison.Ordinal)));
    }

    [Fact]
    public void A_null_policy_is_refused() =>
        _ = Assert.Throws<ArgumentNullException>(
            () => PowerPointRetryLoop.Execute(
                null!,
                retryAfterMutation: true,
                _ => new PowerPointRetryAttempt<string>("ok", null, MutationCommitted: false),
                technicalRecord: null,
                "probe"));

    [Fact]
    public void A_null_attempt_is_refused() =>
        _ = Assert.Throws<ArgumentNullException>(
            () => PowerPointRetryLoop.Execute<string>(
                Policy(1),
                retryAfterMutation: true,
                null!,
                technicalRecord: null,
                "probe"));

    [Fact]
    public void A_blank_stage_is_refused() =>
        _ = Assert.Throws<ArgumentException>(
            () => PowerPointRetryLoop.Execute(
                Policy(1),
                retryAfterMutation: true,
                _ => new PowerPointRetryAttempt<string>("ok", null, MutationCommitted: false),
                technicalRecord: null,
                " "));

    private static PowerPointTransientRetryPolicy Policy(int maxAttempts) => new([_transient], maxAttempts);
}
