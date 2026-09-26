using System.Globalization;
using GanttCreator.Core;
using Xunit;

namespace GanttCreator.Core.Tests;

/// <summary>
/// Contract tests for the approved event-date display format (work item
/// R2.7d, ADR-0016). The format is a schema contract: a silently wrong value
/// would change every rendered date with no visible error, so the strict-parser
/// refusals and the culture-invariance guarantee are each pinned, and the
/// refusals carry the positive tests the AGENTS.md validator rule requires.
/// </summary>
public class GanttDateDisplayFormatTests
{
    [Fact]
    public void The_stored_default_is_the_approved_member_name()
    {
        var setting = GanttCatalogues.Settings.Single(item => item.Key == "DateDisplayFormat");

        Assert.Equal(nameof(GanttDateDisplayFormat.DdMMyyyy), setting.DefaultValue);
        Assert.True(
            GanttChartSettings.TryParseDateDisplayFormat(setting.DefaultValue, out GanttDateDisplayFormat parsed));
        Assert.Equal(GanttDateDisplayFormat.DdMMyyyy, parsed);
    }

    [Fact]
    public void The_date_display_format_key_is_last_in_contract_order()
    {
        // The key order is part of the schema contract, so a key inserted in any
        // other position must fail this rather than silently renumber the table.
        Assert.Equal("DateDisplayFormat", GanttCatalogues.Settings[^1].Key);
    }

    [Theory]
    [InlineData("dd/mm/yyyy")]
    [InlineData("DD/MM/YYYY")]
    [InlineData("DMyyyy")]
    [InlineData("Date")]
    [InlineData("")]
    [InlineData(" ")]
    public void Date_settings_reject_unknown_or_unformatted_text(string text)
    {
        // Positive test for the strict parser: the set is closed (ADR-0016 D3),
        // so a stored value that is not the exact member name is refused rather
        // than coerced. A raw pattern string must never be accepted, or the
        // closed set would become an open one.
        Assert.False(GanttChartSettings.TryParseDateDisplayFormat(text, out _));
    }

    [Fact]
    public void Date_settings_reject_null()
    {
        Assert.False(GanttChartSettings.TryParseDateDisplayFormat(null, out _));
    }

    [Theory]
    [InlineData(2026, 9, 5, "05/09/2026")]
    [InlineData(2026, 12, 31, "31/12/2026")]
    [InlineData(2026, 1, 1, "01/01/2026")]
    [InlineData(2024, 2, 29, "29/02/2024")]
    public void Dates_format_as_the_approved_pattern(int year, int month, int day, string expected)
    {
        var date = new DateOnly(year, month, day);

        Assert.Equal(expected, GanttDateFormatting.Format(date, GanttDateDisplayFormat.DdMMyyyy));
        Assert.Equal(expected, GanttDateFormatting.FormatDdMMyyyy(date));
    }

    [Fact]
    public void Day_first_order_is_distinguishable_from_month_first()
    {
        // The single most consequential property of dd/mm/yyyy: 05/09 is 5
        // September, not 9 May. A month-first renderer would pass every other
        // test in this file and silently transpose unambiguous dates.
        var date = new DateOnly(2026, 9, 5);

        Assert.Equal("05/09/2026", GanttDateFormatting.Format(date, GanttDateDisplayFormat.DdMMyyyy));
        Assert.NotEqual("09/05/2026", GanttDateFormatting.Format(date, GanttDateDisplayFormat.DdMMyyyy));
    }

    [Fact]
    public void Formatting_is_identical_under_a_day_first_and_a_month_first_culture()
    {
        // ADR-0016 D4: the host's regional settings must never change scene
        // text, or the R3.12 golden snapshot would differ per machine.
        var date = new DateOnly(2026, 9, 5);
        var original = CultureInfo.CurrentCulture;
        try
        {
            var dayFirst = new CultureInfo("en-GB");
            var monthFirst = new DateTimeFormatInfo
            {
                ShortDatePattern = "MM/dd/yyyy",
                DateSeparator = "/",
            };
            var us = (CultureInfo)CultureInfo.InvariantCulture.Clone();
            us.DateTimeFormat = monthFirst;

            Assert.NotEqual(dayFirst.DateTimeFormat.ShortDatePattern, us.DateTimeFormat.ShortDatePattern);

            CultureInfo.CurrentCulture = dayFirst;
            var underGb = GanttDateFormatting.Format(date, GanttDateDisplayFormat.DdMMyyyy);
            CultureInfo.CurrentCulture = us;
            var underUs = GanttDateFormatting.Format(date, GanttDateDisplayFormat.DdMMyyyy);

            Assert.Equal(underGb, underUs);
            Assert.Equal("05/09/2026", underUs);
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }

    [Fact]
    public void An_undefined_date_display_format_is_refused()
    {
        // Positive test for the formatter guard: a value outside the closed set
        // must throw rather than fall through to a default pattern.
        var undefined = (GanttDateDisplayFormat)99;

        Assert.Throws<ArgumentOutOfRangeException>(
            () => GanttDateFormatting.Format(new DateOnly(2026, 9, 5), undefined));
    }
}