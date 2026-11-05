namespace GanttCreator.Office.ContractTests;

/// <summary>
/// The unlock set is what suspends the automatic refresh while a manual edit is in flight
/// (ADR-0055 clause 6), so its emptiness is load-bearing rather than bookkeeping.
/// </summary>
public class ShapeUnlockStateTests
{
    [Fact]
    public void A_fresh_state_holds_nothing()
    {
        var state = new ShapeUnlockState();

        Assert.False(state.IsAnyUnlocked);
        Assert.Equal(0, state.Count);
        Assert.False(state.Contains("row-1:bar"));
    }

    [Fact]
    public void Adding_an_identifier_records_it_once()
    {
        var state = new ShapeUnlockState();

        Assert.True(state.Add("row-1:bar"));
        Assert.False(state.Add("row-1:bar"));

        Assert.True(state.IsAnyUnlocked);
        Assert.Equal(1, state.Count);
        Assert.True(state.Contains("row-1:bar"));
    }

    /// <summary>
    /// Clearing is what a Lock command or an explicit Refresh does before it re-locks: the return value
    /// tells the caller whether it had anything to undo, which is what lets the command skip work.
    /// </summary>
    [Fact]
    public void Clearing_forgets_everything_and_reports_whether_it_did()
    {
        var state = new ShapeUnlockState();
        _ = state.Add("row-1:bar");
        _ = state.Add("row-2:milestone");

        Assert.True(state.Clear());
        Assert.False(state.IsAnyUnlocked);
        Assert.False(state.Clear());
    }

    [Fact]
    public void Identifiers_are_compared_ordinal()
    {
        var state = new ShapeUnlockState();
        _ = state.Add("G-0123456789abcdef0123456789abcdef:bar");

        Assert.False(state.Contains("g-0123456789abcdef0123456789abcdef:bar"));
    }
}
