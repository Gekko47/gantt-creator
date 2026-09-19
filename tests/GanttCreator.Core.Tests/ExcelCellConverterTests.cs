using System.Globalization;

namespace GanttCreator.Core.Tests;

public class ExcelCellConverterTests
{
    [Theory]
    [InlineData(1.0, 1899, 12, 31)]
    [InlineData(61.0, 1900, 3, 1)]
    [InlineData(44927.0, 2023, 1, 1)]
    [InlineData(45351.0, 2024, 2, 29)]
    [InlineData(45291.0, 2023, 12, 31)]
    [InlineData(45292.0, 2024, 1, 1)]
    public void TryConvertDate_converts_known_serials(double serial, int year, int month, int day)
    {
        Assert.True(ExcelCellConverter.TryConvertDate(serial, out var date));
        Assert.Equal(new DateOnly(year, month, day), date);
    }

    [Fact]
    public void TryConvertDate_pins_the_1900_leap_bug_phantom()
    {
        // Serial 60 is the inherited 1900 leap-bug phantom (1900-02-29 never
        // existed). DateTime.FromOADate maps it to 1900-02-29; the converter
        // preserves that mapping and must not "fix" it.
        Assert.True(ExcelCellConverter.TryConvertDate(60.0, out var date));
        Assert.Equal(DateOnly.FromDateTime(DateTime.FromOADate(60.0)), date);
    }

    [Fact]
    public void TryConvertDate_converts_a_datetime_payload()
    {
        var payload = new DateTime(2026, 9, 19, 15, 30, 0, DateTimeKind.Unspecified);

        Assert.True(ExcelCellConverter.TryConvertDate(payload, out var date));
        Assert.Equal(new DateOnly(2026, 9, 19), date);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("2026-09-19")]
    [InlineData("19/09/2026")]
    [InlineData(true)]
    [InlineData(false)]
    public void TryConvertDate_rejects_non_cell_value_dates(object? value)
    {
        Assert.False(ExcelCellConverter.TryConvertDate(value, out var date));
        Assert.Null(date);
    }

    [Theory]
    [InlineData(0.0)]
    [InlineData(-1.0)]
    [InlineData(2958466.0)]
    [InlineData(10000000.0)]
    public void TryConvertDate_rejects_out_of_range_serials(double serial)
    {
        Assert.False(ExcelCellConverter.TryConvertDate(serial, out var date));
        Assert.Null(date);
    }

    [Fact]
    public void TryConvertDate_rejects_non_finite_serials()
    {
        Assert.False(ExcelCellConverter.TryConvertDate(double.NaN, out var nan));
        Assert.Null(nan);
        Assert.False(ExcelCellConverter.TryConvertDate(double.PositiveInfinity, out var pos));
        Assert.Null(pos);
        Assert.False(ExcelCellConverter.TryConvertDate(double.NegativeInfinity, out var neg));
        Assert.Null(neg);
    }

    [Fact]
    public void TryConvertDate_maps_excel_error_int_to_null()
    {
        Assert.False(ExcelCellConverter.TryConvertDate(-2146826281, out var date));
        Assert.Null(date);
    }

    [Theory]
    [InlineData(0.0, 0)]
    [InlineData(1.0, 1)]
    [InlineData(3.0, 3)]
    public void TryConvertStackIndex_converts_whole_doubles(double value, int expected)
    {
        Assert.True(ExcelCellConverter.TryConvertStackIndex(value, out var index));
        Assert.Equal(expected, index);
    }

    [Theory]
    [InlineData("3", 3)]
    [InlineData("  2  ", 2)]
    [InlineData("0", 0)]
    public void TryConvertStackIndex_converts_invariant_strings(string text, int expected)
    {
        Assert.True(ExcelCellConverter.TryConvertStackIndex(text, out var index));
        Assert.Equal(expected, index);
    }

    [Theory]
    [InlineData(1.5)]
    [InlineData(-1.0)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    public void TryConvertStackIndex_rejects_bad_doubles(double value)
    {
        Assert.False(ExcelCellConverter.TryConvertStackIndex(value, out var index));
        Assert.Null(index);
    }

    [Theory]
    [InlineData("-1")]
    [InlineData("1.5")]
    [InlineData("abc")]
    [InlineData("")]
    [InlineData("   ")]
    public void TryConvertStackIndex_rejects_bad_strings(string? text)
    {
        Assert.False(ExcelCellConverter.TryConvertStackIndex(text, out var index));
        Assert.Null(index);
    }

    [Fact]
    public void TryConvertStackIndex_rejects_null_bool_error_and_datetime()
    {
        Assert.False(ExcelCellConverter.TryConvertStackIndex(null, out var missing));
        Assert.Null(missing);
        Assert.False(ExcelCellConverter.TryConvertStackIndex(true, out var boolean));
        Assert.Null(boolean);
        Assert.False(ExcelCellConverter.TryConvertStackIndex(-2146826281, out var error));
        Assert.Null(error);
        Assert.False(ExcelCellConverter.TryConvertStackIndex(new DateTime(2026, 1, 1), out var dateTime));
        Assert.Null(dateTime);
    }

    [Theory]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(1.0, true)]
    [InlineData(0.0, false)]
    public void TryConvertVisible_converts_bools_and_ones_and_zeroes(object value, bool expected)
    {
        Assert.True(ExcelCellConverter.TryConvertVisible(value, out var visible));
        Assert.Equal(expected, visible);
    }

    [Theory]
    [InlineData("TRUE", true)]
    [InlineData("true", true)]
    [InlineData("FALSE", false)]
    [InlineData("false", false)]
    [InlineData("1", true)]
    [InlineData("0", false)]
    [InlineData("  TRUE  ", true)]
    public void TryConvertVisible_converts_invariant_strings(string text, bool expected)
    {
        Assert.True(ExcelCellConverter.TryConvertVisible(text, out var visible));
        Assert.Equal(expected, visible);
    }

    [Theory]
    [InlineData(0.5)]
    [InlineData(2.0)]
    [InlineData("yes")]
    [InlineData("maybe")]
    [InlineData("")]
    [InlineData("   ")]
    public void TryConvertVisible_rejects_bad_values(object? value)
    {
        Assert.False(ExcelCellConverter.TryConvertVisible(value, out var visible));
        Assert.Null(visible);
    }

    [Fact]
    public void TryConvertVisible_rejects_null_error_and_datetime()
    {
        Assert.False(ExcelCellConverter.TryConvertVisible(null, out var missing));
        Assert.Null(missing);
        Assert.False(ExcelCellConverter.TryConvertVisible(-2146826281, out var error));
        Assert.Null(error);
        Assert.False(ExcelCellConverter.TryConvertVisible(new DateTime(2026, 1, 1), out var dateTime));
        Assert.Null(dateTime);
    }

    [Theory]
    [InlineData("  As-Planned Activity  ", "As-Planned Activity")]
    [InlineData("Gantt", "Gantt")]
    [InlineData("", null)]
    [InlineData("   ", null)]
    [InlineData(null, null)]
    public void ToText_trims_and_maps_blank_to_null(string? input, string? expected)
    {
        Assert.Equal(expected, ExcelCellConverter.ToText(input));
    }

    [Fact]
    public void ToText_maps_excel_error_int_to_null()
    {
        Assert.Null(ExcelCellConverter.ToText(-2146826281));
    }

    [Fact]
    public void ToText_formats_bool_double_and_datetime_invariantly()
    {
        Assert.Equal("TRUE", ExcelCellConverter.ToText(true));
        Assert.Equal("FALSE", ExcelCellConverter.ToText(false));
        Assert.Equal("1.5", ExcelCellConverter.ToText(1.5));
        Assert.Equal(
            new DateTime(2026, 9, 19).ToString("o", CultureInfo.InvariantCulture),
            ExcelCellConverter.ToText(new DateTime(2026, 9, 19)));
    }

    [Theory]
    [InlineData("tr-TR")]
    [InlineData("de-DE")]
    public void Converters_are_culture_invariant(string cultureName)
    {
        var original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = new CultureInfo(cultureName);

            Assert.True(ExcelCellConverter.TryConvertDate(44927.0, out var date));
            Assert.Equal(new DateOnly(2023, 1, 1), date);
            Assert.True(ExcelCellConverter.TryConvertStackIndex("3", out var index));
            Assert.Equal(3, index);
            Assert.True(ExcelCellConverter.TryConvertVisible("TRUE", out var visible));
            Assert.True(visible);
            Assert.Equal("As-Planned Activity", ExcelCellConverter.ToText("  As-Planned Activity  "));
            Assert.Equal("1.5", ExcelCellConverter.ToText(1.5));
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }
}
