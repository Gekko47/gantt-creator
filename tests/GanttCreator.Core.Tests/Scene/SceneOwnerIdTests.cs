using GanttCreator.Core;
using GanttCreator.Core.Scene;
using Xunit;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// Contract tests for the shared scene owner (ADR-0017). Ownership must be
/// truthful and canonically encoded: the scene snapshot is compared
/// byte-for-byte in R3.12, so two spellings of one row set would be a
/// golden-file failure discovered much later.
/// </summary>
public class SceneOwnerIdTests
{
    private static GanttRowId Row(char hex) => GanttRowId.Parse($"G-{new string(hex, 32)}");

    [Fact]
    public void A_row_owner_reports_exactly_one_owned_row()
    {
        var id = Row('a');
        var owner = SceneOwnerId.ForRow(id);

        Assert.Equal(SceneOwnerKind.Row, owner.Kind);
        Assert.Equal([id], owner.OwnedRows);
    }

    [Fact]
    public void The_chart_owner_reports_no_owned_rows()
    {
        Assert.Equal(SceneOwnerKind.Chart, SceneOwnerId.Chart.Kind);
        Assert.Empty(SceneOwnerId.Chart.OwnedRows);
    }

    [Fact]
    public void A_shared_owner_sorts_and_deduplicates_its_rows()
    {
        var a = Row('a');
        var b = Row('b');
        var c = Row('c');

        // Deliberately unsorted and with a repeat: the canonical form must not
        // depend on the caller's order, or two callers would build two owners.
        var owner = SceneOwnerId.ForRows([c, a, b, a]);

        Assert.Equal(SceneOwnerKind.Rows, owner.Kind);
        Assert.Equal([a, b, c], owner.OwnedRows);
        Assert.Equal($"{a.Value}|{b.Value}|{c.Value}", owner.Value);
    }

    [Fact]
    public void A_shared_owner_parses_back_to_the_same_value_and_rows()
    {
        var owner = SceneOwnerId.ForRows([Row('b'), Row('a')]);

        Assert.True(SceneOwnerId.TryParse("rows", owner.Value, out SceneOwnerId? parsed));
        Assert.Equal(owner, parsed);
        Assert.Equal(owner.OwnedRows, parsed!.OwnedRows);
    }

    [Fact]
    public void A_shared_primitive_identifier_contains_every_owning_row()
    {
        var a = Row('a');
        var b = Row('b');
        var owner = SceneOwnerId.ForRows([b, a]);

        // ADR-0017 D4: membership-sensitive, so a membership change is visible
        // as an identifier change rather than an in-place mutation.
        var id = ScenePrimitive.CreateId(owner, "delineator");

        Assert.Equal($"{a.Value}|{b.Value}:delineator", id);
        Assert.NotEqual(
            ScenePrimitive.CreateId(SceneOwnerId.ForRow(a), "delineator"),
            id);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void A_shared_owner_requires_at_least_two_distinct_rows(int distinctCount)
    {
        // Positive test for the D2 guard: one spelling per ownership, so a
        // single row can never be written as a shared owner. The rows are built
        // from distinctCount directly: the zero case must pass an *empty*
        // collection, and clamping to one row would make it a duplicate of the
        // one case instead of exercising the empty input.
        var id = Row('a');
        var rows = Enumerable.Range(0, distinctCount).Select(_ => id).ToList();

        Assert.Throws<ArgumentOutOfRangeException>(() => SceneOwnerId.ForRows(rows));
    }

    [Fact]
    public void A_shared_owner_rejects_a_null_sequence()
    {
        Assert.Throws<ArgumentNullException>(
            () => SceneOwnerId.ForRows(null!));
    }

    [Theory]
    [InlineData("not-an-id|G-0000000000000000000000000000000a")] // malformed part
    [InlineData("G-0000000000000000000000000000000a")]            // only one part
    [InlineData("")]                                               // empty
    public void A_shared_owner_refuses_a_non_canonical_value(string value)
    {
        Assert.False(SceneOwnerId.TryParse("rows", value, out _));
    }

    [Fact]
    public void A_shared_owner_refuses_an_unsorted_or_repeated_value()
    {
        // D3: refused, not normalised, so a hand-edited snapshot cannot carry an
        // ordering the builder can never produce.
        var a = Row('a');
        var b = Row('b');

        Assert.False(SceneOwnerId.TryParse("rows", $"{b.Value}|{a.Value}", out _));
        Assert.False(SceneOwnerId.TryParse("rows", $"{a.Value}|{a.Value}", out _));
    }

    [Theory]
    [InlineData("Rows")]
    [InlineData("ROWS")]
    [InlineData("rowset")]
    [InlineData("")]
    [InlineData(null)]
    public void A_shared_owner_requires_the_exact_kind_text(string? kind)
    {
        var value = $"{Row('a').Value}|{Row('b').Value}";

        // The exact lowercase "rows" is accepted; anything else is refused
        // rather than case-folded, matching the "row" and "chart" kinds.
        Assert.True(SceneOwnerId.TryParse("rows", value, out _));
        Assert.False(SceneOwnerId.TryParse(kind, value, out _));
    }
}