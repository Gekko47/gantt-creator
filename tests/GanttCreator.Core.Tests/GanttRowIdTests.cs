using System.Globalization;

namespace GanttCreator.Core.Tests;

public class GanttRowIdTests
{
    private sealed record TestRow(GanttRowId Id, string Probe);

    [Fact]
    public void New_generates_a_wellformed_identifier()
    {
        var id = GanttRowId.New();

        Assert.Equal(GanttRowId.TextLength, id.Value.Length);
        Assert.StartsWith(GanttRowId.Prefix, id.Value, StringComparison.Ordinal);
        Assert.True(GanttRowId.TryParse(id.Value, out var parsed));
        Assert.Equal(id, parsed);
    }

    [Fact]
    public void New_generates_unique_identifiers()
    {
        const int count = 1000;
        var values = new HashSet<string>(StringComparer.Ordinal);
        for (int i = 0; i < count; i++)
        {
            Assert.True(values.Add(GanttRowId.New().Value), "Duplicate generated identifier.");
        }
    }

    [Fact]
    public void Generated_identifiers_are_excel_text_safe()
    {
        for (int i = 0; i < 100; i++)
        {
            string value = GanttRowId.New().Value;
            Assert.Equal('G', value[0]);
            Assert.Equal(GanttRowId.TextLength, value.Length);
        }
    }

    [Fact]
    public void Roundtrip_preserves_the_exact_value()
    {
        for (int i = 0; i < 100; i++)
        {
            var original = GanttRowId.New();

            Assert.True(GanttRowId.TryParse(original.Value, out var parsed));
            Assert.Equal(original.Value, parsed!.Value);
            Assert.Equal(original, parsed);
            Assert.Equal(original.Value, parsed.ToString());
        }
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("G-")]
    [InlineData("G-123")]
    [InlineData("0123456789abcdef0123456789abcdef")]
    [InlineData("X-0123456789abcdef0123456789abcdef")]
    [InlineData("G-0123456789abcdef0123456789abcde")]
    [InlineData("G-0123456789abcdef0123456789abcdef0")]
    [InlineData("G-0123456789abcdef0123456789abcdeg")]
    [InlineData("G-0123456789ABCDEF0123456789ABCDEF")]
    [InlineData("G-0123456789abcdef 123456789abcdef")]
    public void TryParse_rejects_malformed_input_without_throwing(string? text)
    {
        Assert.False(GanttRowId.TryParse(text, out var id));
        Assert.Null(id);
    }

    [Fact]
    public void TryParse_ignores_surrounding_whitespace_but_not_case()
    {
        string canonical = GanttRowId.New().Value;

        Assert.True(GanttRowId.TryParse("  " + canonical + "\t", out var parsed));
        Assert.Equal(canonical, parsed!.Value);
        Assert.False(GanttRowId.TryParse(canonical.ToUpperInvariant(), out _));
    }

    [Fact]
    public void Parse_throws_for_malformed_input()
    {
        Assert.Throws<ArgumentNullException>(() => GanttRowId.Parse(null));
        Assert.Throws<FormatException>(() => GanttRowId.Parse(string.Empty));
        Assert.Throws<FormatException>(() => GanttRowId.Parse("not-an-id"));
    }

    [Fact]
    public void Equality_is_by_value_ordinal()
    {
        string value = GanttRowId.New().Value;
        var left = GanttRowId.Parse(value);
        var right = GanttRowId.Parse(value);

        Assert.Equal(left, right);
        Assert.True(left == right);
        Assert.False(left != right);
        Assert.Equal(left.GetHashCode(), right.GetHashCode());
        Assert.NotEqual(left, GanttRowId.New());
        Assert.False(left.Equals(null));
        Assert.False(left.Equals("a string"));
        Assert.False(left == null);
        Assert.True((GanttRowId?)null == (GanttRowId?)null);
    }

    [Fact]
    public void Identifiers_survive_insert_sort_move_and_delete()
    {
        var rows = new List<TestRow>
        {
            new(GanttRowId.New(), "bravo"),
            new(GanttRowId.New(), "alpha"),
            new(GanttRowId.New(), "charlie"),
        };

        rows.Insert(1, new TestRow(GanttRowId.New(), "inserted-middle"));
        rows.Insert(0, new TestRow(GanttRowId.New(), "inserted-top"));
        rows.Add(new TestRow(GanttRowId.New(), "appended-bottom"));

        var pairing = rows.ToDictionary(row => row.Probe, row => row.Id, StringComparer.Ordinal);

        rows.Sort((a, b) => string.Compare(a.Probe, b.Probe, StringComparison.Ordinal));

        var head = rows[0];
        rows.RemoveAt(0);
        rows.Add(head);
        var tail = rows[^1];
        rows.RemoveAt(rows.Count - 1);
        rows.Insert(0, tail);

        rows.RemoveAt(0);
        rows.RemoveAt(rows.Count - 1);
        rows.RemoveAt(1);

        Assert.NotEmpty(rows);
        foreach (var row in rows)
        {
            Assert.Equal(pairing[row.Probe], row.Id);
        }

        Assert.Equal(rows.Count, rows.Select(row => row.Id).Distinct().Count());
    }

    [Fact]
    public void Identifier_surfaces_are_culture_invariant_under_tr_tr()
    {
        var originalCulture = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("tr-TR");

            var id = GanttRowId.New();
            Assert.True(GanttRowId.TryParse(id.Value, out var parsed));
            Assert.Equal(id, parsed);
            Assert.False(GanttRowId.TryParse(id.Value.ToUpperInvariant(), out _));
        }
        finally
        {
            CultureInfo.CurrentCulture = originalCulture;
        }
    }
}
