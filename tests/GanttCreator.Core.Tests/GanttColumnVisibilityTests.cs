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
    public void Hiding_every_authoring_column_is_allowed_and_one_can_be_shown_again()
    {
        // ADR-0065: zero visible columns is a supported state. The old last-column refusal
        // is gone, so hiding the sixth column succeeds, and the same toggle shows one again.
        ColumnVisibilityOutcome all = GanttColumnVisibility.Toggle(
            Set("Type", "Activity ID", "Description", "Start", "Finish"),
            "Duration");

        Assert.True(all.Succeeded);
        Assert.Equal(
            Set("Type", "Activity ID", "Description", "Start", "Finish", "Duration"),
            all.Hidden!);

        ColumnVisibilityOutcome shown = GanttColumnVisibility.Toggle(all.Hidden!, "Type");
        Assert.True(shown.Succeeded);
        Assert.DoesNotContain("Type", shown.Hidden!);
        Assert.Equal(5, shown.Hidden!.Count);
    }

    [Fact]
    public void Toggling_all_hides_every_offered_column_and_toggling_again_shows_them()
    {
        ColumnVisibilityOutcome hidden = GanttColumnVisibility.ToggleAll(Set());

        Assert.True(hidden.Succeeded);
        Assert.Equal(Set([.. GanttColumnVisibility.ToggleableColumnNames]), hidden.Hidden!);

        ColumnVisibilityOutcome shown = GanttColumnVisibility.ToggleAll(hidden.Hidden!);

        Assert.True(shown.Succeeded);
        Assert.Empty(shown.Hidden!);
    }

    [Fact]
    public void A_partial_state_shows_all_rather_than_completing_the_hide()
    {
        // ADR-0065 D3: any hidden column means show all. The primary is a forgiving reset;
        // precise hiding belongs to the menu rows.
        ColumnVisibilityOutcome outcome = GanttColumnVisibility.ToggleAll(Set("Start"));

        Assert.True(outcome.Succeeded);
        Assert.Empty(outcome.Hidden!);
    }

    [Fact]
    public void The_full_hidden_set_round_trips_in_schema_order()
    {
        IReadOnlySet<string> all = Set([.. GanttColumnVisibility.ToggleableColumnNames]);

        string text = GanttColumnVisibility.Format(all);

        Assert.Equal(string.Join(',', GanttColumnVisibility.ToggleableColumnNames), text);
        Assert.Equal(all, GanttColumnVisibility.Parse(text).Hidden!);
    }

    [Fact]
    public void Visible_columns_are_the_schema_visible_set_minus_the_hidden_names()
    {
        IReadOnlyList<string> expected =
            [.. GanttTableSchema.Default.Columns
                .Where(column => !column.IsHidden)
                .Select(column => column.Name)];

        Assert.Equal(expected, GanttColumnVisibility.VisibleColumns(Set()));
        Assert.Equal(
            [.. expected.Where(name => !string.Equals(name, "Start", StringComparison.Ordinal))],
            GanttColumnVisibility.VisibleColumns(Set("Start")));
        Assert.Empty(GanttColumnVisibility.VisibleColumns(Set([.. expected])));
    }

    [Fact]
    public void A_null_toggle_all_or_visible_columns_input_is_refused()
    {
        _ = Assert.Throws<ArgumentNullException>(() => GanttColumnVisibility.ToggleAll(null!));
        _ = Assert.Throws<ArgumentNullException>(() => GanttColumnVisibility.VisibleColumns(null!));
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

    [Fact]
    public void Every_ribbon_item_id_is_a_legal_st_UniqueID_and_resolves_back()
    {
        // RibbonX types `id` as ST_UniqueID, which rejects a space: the live host refused
        // `colActivity ID` on load on 2026-11-17. The display name keeps its space (it is
        // the header), the id is the name with spaces removed, and the resolver is the
        // inverse the callbacks use.
        Assert.Equal("ActivityID", GanttColumnVisibility.RibbonItemId("Activity ID"));
        Assert.Equal("Start", GanttColumnVisibility.RibbonItemId("Start"));

        foreach (string name in GanttColumnVisibility.ToggleableColumnNames)
        {
            string id = GanttColumnVisibility.RibbonItemId(name);
            Assert.DoesNotContain(id, char.IsWhiteSpace);
            Assert.True(GanttColumnVisibility.TryResolveRibbonItemId(id, out string? resolved));
            Assert.Equal(name, resolved);
        }

        Assert.False(GanttColumnVisibility.TryResolveRibbonItemId("Activity ID", out _));
        Assert.False(GanttColumnVisibility.TryResolveRibbonItemId("Id", out _));
        Assert.False(GanttColumnVisibility.TryResolveRibbonItemId(null, out _));
    }

    [Fact]
    public void The_derived_ribbon_item_ids_are_unique()
    {
        // Two display names that differ only by a space would collide once slugged, and
        // one of the two rows would control the wrong column.
        HashSet<string> ids = new(StringComparer.Ordinal);
        foreach (string name in GanttColumnVisibility.ToggleableColumnNames)
        {
            Assert.True(ids.Add(GanttColumnVisibility.RibbonItemId(name)), $"{name} slug collides.");
        }
    }
}
