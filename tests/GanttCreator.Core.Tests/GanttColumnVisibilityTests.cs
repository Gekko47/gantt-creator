namespace GanttCreator.Core.Tests;

/// <summary>
/// Contracts for <see cref="GanttColumnVisibility"/> (R5.18h): which columns the dropdown
/// may hide, what the settings key may contain, and what must be refused.
/// </summary>
public sealed class GanttColumnVisibilityTests
{
    private static HashSet<string> Set(params string[] names) => new(names, StringComparer.Ordinal);

    [Fact]
    public void Every_toggleable_column_is_a_visible_schema_column()
    {
        // A name that is not a visible column would offer the user a toggle that hides
        // nothing they can see, or worse, one of the engine columns the add-in owns.
        foreach (string name in GanttColumnVisibility.ToggleableColumnNames)
        {
            Assert.True(
                GanttTableSchema.Default.TryGetColumn(name, out GanttTableColumn? column),
                $"{name} is not a schema column.");
            Assert.NotEqual(GanttColumnAccess.EngineHidden, column!.Access);
        }
    }

    [Fact]
    public void A_blank_stored_value_means_nothing_hidden()
    {
        Assert.Empty(GanttColumnVisibility.Parse(null).Hidden!);
        Assert.Empty(GanttColumnVisibility.Parse(string.Empty).Hidden!);
        Assert.Empty(GanttColumnVisibility.Parse("   ").Hidden!);
    }

    [Fact]
    public void A_stored_list_round_trips_and_is_written_in_schema_order()
    {
        ColumnVisibilityOutcome parsed = GanttColumnVisibility.Parse("Start, Type");

        Assert.True(parsed.Succeeded);
        Assert.Equal(Set("Type", "Start"), parsed.Hidden!);

        // Schema order, not the order the user happened to hide them in, so two workbooks
        // holding the same state store identical text.
        Assert.Equal("Type,Start", GanttColumnVisibility.Format(parsed.Hidden!));
    }

    [Fact]
    public void An_unknown_name_is_refused_rather_than_ignored()
    {
        ColumnVisibilityOutcome outcome = GanttColumnVisibility.Parse("Type,Id");

        Assert.Equal(ColumnVisibilityRefusal.UnknownColumn, outcome.Refusal);
        Assert.Null(outcome.Hidden);

        // The refusal must not carry a usable state: `Succeeded` is the conjunction of
        // "a set came back" and "nothing was refused", and a caller that trusted a
        // disjunction would apply a set the parser had already rejected.
        Assert.False(outcome.Succeeded);
    }

    [Fact]
    public void A_malformed_list_is_refused_rather_than_read_as_nothing_hidden()
    {
        // "Type,,Start" came from something other than Format. Reading it as "only Type"
        // would show a column the user hid and look like a fresh workbook.
        Assert.Equal(
            ColumnVisibilityRefusal.MalformedList,
            GanttColumnVisibility.Parse("Type,,Start").Refusal);
    }

    [Fact]
    public void Toggling_a_visible_column_hides_it_and_toggling_again_shows_it()
    {
        ColumnVisibilityOutcome hidden = GanttColumnVisibility.Toggle(Set(), "Start");
        Assert.True(hidden.Succeeded);
        Assert.Equal(Set("Start"), hidden.Hidden!);

        ColumnVisibilityOutcome shown = GanttColumnVisibility.Toggle(hidden.Hidden!, "Start");
        Assert.True(shown.Succeeded);
        Assert.Empty(shown.Hidden!);
    }

    [Fact]
    public void Hiding_the_last_visible_authoring_column_is_refused()
    {
        // Type, Description and Start are already hidden; hiding Finish would leave no
        // authoring column at all, so the table could not be typed into by hand.
        // `Duration` does not count: the schema classifies it as read-only.
        ColumnVisibilityOutcome outcome = GanttColumnVisibility.Toggle(
            Set("Type", "Description", "Start"),
            "Finish");

        Assert.False(outcome.Succeeded);
        Assert.Null(outcome.Hidden);
        Assert.Equal(ColumnVisibilityRefusal.LastVisibleColumn, outcome.Refusal);
    }

    [Fact]
    public void A_column_the_dropdown_does_not_offer_cannot_be_toggled()
    {
        Assert.Equal(
            ColumnVisibilityRefusal.UnknownColumn,
            GanttColumnVisibility.Toggle(Set(), "Id").Refusal);
        Assert.Equal(
            ColumnVisibilityRefusal.UnknownColumn,
            GanttColumnVisibility.Toggle(Set(), null).Refusal);
    }

    [Fact]
    public void A_null_set_is_refused_rather_than_treated_as_nothing_hidden() =>
        _ = Assert.Throws<ArgumentNullException>(() => GanttColumnVisibility.Toggle(null!, "Type"));

    [Fact]
    public void A_null_format_input_is_refused() =>
        _ = Assert.Throws<ArgumentNullException>(() => GanttColumnVisibility.Format(null!));

    [Fact]
    public void Only_an_engine_column_has_a_schema_owned_hidden_flag()
    {
        // The Refresh and Repair writers ask this before forcing a column's visibility:
        // an engine column is the schema's, every user-visible one is the dropdown's.
        Assert.True(GanttColumnVisibility.SchemaOwnsVisibility(GanttTableSchema.Default.Columns
            .Single(column => column.Name == "Id")));
        Assert.True(GanttColumnVisibility.SchemaOwnsVisibility(GanttTableSchema.Default.Columns
            .Single(column => column.Name == "StyleKey")));

        foreach (string name in GanttColumnVisibility.ToggleableColumnNames)
        {
            Assert.False(
                GanttColumnVisibility.SchemaOwnsVisibility(GanttTableSchema.Default.Columns
                    .Single(column => column.Name == name)),
                $"{name} is user-visible, so its hidden flag is the dropdown's choice.");
        }

        Assert.False(GanttColumnVisibility.SchemaOwnsVisibility(null));
    }
}
