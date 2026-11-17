namespace GanttCreator.Core.Tests;

/// <summary>
/// Tests for <see cref="ChangeTracker"/>, the pending-change batch a scheduled
/// refresh reads before it decides which rows to re-read. The class had no
/// direct test: the contract tests only called <c>Record</c>, so
/// <c>ChangedRows</c>, <c>RowChanged</c>, and <c>Reset</c> were unexercised and
/// a regression in any of them would have surfaced as a wrong-row refresh in
/// the live host rather than as a failing test.
/// </summary>
public sealed class ChangeTrackerTests
{
    [Fact]
    public void A_new_tracker_reports_no_changes()
    {
        ChangeTracker tracker = new();

        Assert.False(tracker.HasChanges);
        Assert.Empty(tracker.Changes);
        Assert.Empty(tracker.ChangedRows());
    }

    [Fact]
    public void A_recorded_change_keeps_its_row_column_and_before_after_values()
    {
        ChangeTracker tracker = new();

        tracker.Record(3, 4, "old", "new");

        CellChange change = Assert.Single(tracker.Changes);
        Assert.Equal(3, change.RowNumber);
        Assert.Equal(4, change.Column);
        Assert.Equal("old", change.OldValue);
        Assert.Equal("new", change.NewValue);
        Assert.True(tracker.HasChanges);
    }

    [Fact]
    public void Changed_rows_are_distinct_and_in_ascending_order()
    {
        // The refresh reads rows, not edits: two edits on one row must yield one
        // row, and the order must be deterministic so the read range is stable.
        ChangeTracker tracker = new();
        tracker.Record(5, 1, null, "a");
        tracker.Record(2, 1, null, "b");
        tracker.Record(5, 2, null, "c");

        Assert.Equal([2, 5], tracker.ChangedRows());
    }

    [Fact]
    public void A_blank_before_and_after_value_is_recorded_rather_than_skipped()
    {
        // A cleared cell is a chart-affecting edit: old value present, new value
        // null. Dropping it would leave the chart showing the stale bar.
        ChangeTracker tracker = new();

        tracker.Record(1, 6, "As-Planned Activity", null);

        CellChange change = Assert.Single(tracker.Changes);
        Assert.Equal("As-Planned Activity", change.OldValue);
        Assert.Null(change.NewValue);
    }

    [Fact]
    public void Reset_clears_every_recorded_change()
    {
        ChangeTracker tracker = new();
        tracker.Record(1, 1, null, "a");
        tracker.Record(2, 1, null, "b");

        tracker.Reset();

        Assert.False(tracker.HasChanges);
        Assert.Empty(tracker.Changes);
        Assert.Empty(tracker.ChangedRows());
        Assert.False(tracker.RowChanged(1));
    }

    [Fact]
    public void RowChanged_reports_only_the_rows_that_were_recorded()
    {
        ChangeTracker tracker = new();
        tracker.Record(7, 3, null, "x");

        Assert.True(tracker.RowChanged(7));
        Assert.False(tracker.RowChanged(8));
    }
}
