namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// Host-free facts for <see cref="PowerPointFixture"/>. They carry no
/// <c>OfficeIntegration</c> trait, so the ordinary test step runs them: the paste
/// validator and the process-identity guard are the two places a wrong answer would
/// be reported as a successful transfer or would kill a user-owned instance.
/// </summary>
public class PowerPointFixtureTests
{
    [Fact]
    public void Paste_outcome_validation_rejects_zero_shapes()
    {
        InvalidOperationException error = Assert.Throws<InvalidOperationException>(
            () => PowerPointFixture.ValidatePasteOutcome(shapeCount: 0, widthPt: 100, heightPt: 100));

        Assert.Contains("no shapes", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Paste_outcome_validation_rejects_non_positive_bounds()
    {
        _ = Assert.Throws<InvalidOperationException>(
            () => PowerPointFixture.ValidatePasteOutcome(shapeCount: 1, widthPt: 0, heightPt: 100));
        _ = Assert.Throws<InvalidOperationException>(
            () => PowerPointFixture.ValidatePasteOutcome(shapeCount: 1, widthPt: 100, heightPt: -1));
        _ = Assert.Throws<InvalidOperationException>(
            () => PowerPointFixture.ValidatePasteOutcome(shapeCount: 1, widthPt: double.NaN, heightPt: 100));
    }

    [Fact]
    public void Paste_outcome_validation_accepts_a_positive_range()
    {
        PowerPointFixture.ValidatePasteOutcome(shapeCount: 1, widthPt: 100, heightPt: 50);
    }

    [Fact]
    public void IsOwnedProcessIdentity_rejects_an_unknown_process()
    {
        Assert.False(PowerPointFixture.IsOwnedProcessIdentity(int.MaxValue, expectedStartTicks: 0));
    }

    [Fact]
    public void IsOwnedProcessIdentity_rejects_a_recycled_process_id_with_a_different_start_time()
    {
        int currentProcessId = Environment.ProcessId;

        Assert.False(
            PowerPointFixture.IsOwnedProcessIdentity(currentProcessId, expectedStartTicks: 1));
    }

    [Fact]
    public async Task BuildDescription_is_unknown_before_any_powerpoint_process_is_owned()
    {
        PowerPointFixture fixture = new();

        try
        {
            Assert.Equal("unknown", fixture.BuildDescription, StringComparer.Ordinal);
        }
        finally
        {
            // Disposing an uninitialised fixture is a no-op; the call exists so the
            // ownership rules accept a fixture that is constructed in a test.
            await fixture.DisposeAsync().ConfigureAwait(true);
        }
    }
}
