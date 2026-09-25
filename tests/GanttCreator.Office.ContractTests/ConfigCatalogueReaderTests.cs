using Excel = Microsoft.Office.Interop.Excel;
using GanttCreator.Core;
using Moq;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for <see cref="ExcelConfigCatalogueReader"/>. They need no
/// live Office: the configuration worksheet is the
/// <see cref="ConfigSheetFake"/> graph, and every damaged-catalogue fixture
/// refuses with the exact typed reason (AGENTS.md validator rule). The
/// round-trip test runs the real writer then the real reader through the
/// seams and proves the outcome equals the code-owned Core catalogue.
/// </summary>
public class ConfigCatalogueReaderTests
{
    [Fact]
    public void Read_refuses_when_the_milestone_size_column_drifts()
    {
        // Regression: the tblGanttStyles numeric columns are not contiguous
        // (MilestoneSizePt is column 10, after TextColour at 7). A contiguous
        // 5..9 mapping compared TextColour against StandardOutlinePt and
        // never read MilestoneSizePt, so every style row reported
        // CatalogueMismatch and this column's drift went undetected.
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        fake.Tables[1].Body[2][10] = 99.0;

        var outcome = ConfigGraph.BuildReader(fake).Read();

        Assert.Equal(ConfigReadOutcome.Refused(ConfigReadRefusalReason.CatalogueMismatch), outcome);
    }

    [Fact]
    public void Read_refuses_when_the_standard_outline_column_drifts()
    {
        // The same mapping, for the StandardOutlinePt column (8).
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        fake.Tables[1].Body[2][8] = 99.0;

        var outcome = ConfigGraph.BuildReader(fake).Read();

        Assert.Equal(ConfigReadOutcome.Refused(ConfigReadRefusalReason.CatalogueMismatch), outcome);
    }

    [Fact]
    public void Read_refuses_when_a_style_colour_is_lowercase_hex()
    {
        // Case sensitivity of the #RRGGBB contract: "#ff0000" has the right
        // shape but not the required uppercase form.
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        fake.Tables[1].Body[0][2] = "#ff0000";

        var outcome = ConfigGraph.BuildReader(fake).Read();

        Assert.Equal(ConfigReadOutcome.Refused(ConfigReadRefusalReason.BadColourFormat), outcome);
    }

    [Fact]
    public void Read_reports_a_valid_but_different_colour_as_catalogue_drift()
    {
        // Taxonomy pin: a well-formed colour that is not the catalogue value
        // is drift (CatalogueMismatch), not a format problem. This is why the
        // format check must run before the exact-match check — with
        // exact-match first, BadColourFormat is unreachable for style rows.
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        fake.Tables[1].Body[0][2] = "#123456";

        var outcome = ConfigGraph.BuildReader(fake).Read();

        Assert.Equal(ConfigReadOutcome.Refused(ConfigReadRefusalReason.CatalogueMismatch), outcome);
    }

    [Fact]
    public void Read_round_trips_a_writer_payload_to_the_code_owned_catalogues()
    {
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();

        var outcome = ConfigGraph.BuildReader(fake).Read();

        Assert.True(outcome.Succeeded);
        Assert.True(Guid.TryParse(outcome.WorkbookId, out _));
        Assert.Equal(GanttCatalogues.Settings.Count, outcome.Settings.Count);
        foreach (GanttSettingDefinition setting in GanttCatalogues.Settings)
        {
            Assert.Equal(setting.DefaultValue, outcome.Settings[setting.Key]);
        }
    }

    [Fact]
    public void Read_reports_no_active_workbook_without_an_application_object() =>
        Assert.Equal(
            ConfigReadOutcome.Refused(ConfigReadRefusalReason.NoActiveWorkbook),
            new ExcelConfigCatalogueReader(null).Read());

    [Fact]
    public void Read_degrades_to_no_active_workbook_for_a_foreign_object() =>
        Assert.Equal(
            ConfigReadOutcome.Refused(ConfigReadRefusalReason.NoActiveWorkbook),
            new ExcelConfigCatalogueReader("not an Excel application").Read());

    [Fact]
    public void Read_reports_no_active_workbook_when_excel_has_no_active_workbook()
    {
        var application = new Mock<Excel.Application>();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns((Excel.Workbook)null!);

        Assert.Equal(
            ConfigReadOutcome.Refused(ConfigReadRefusalReason.NoActiveWorkbook),
            new ExcelConfigCatalogueReader(application.Object).Read());
    }

    [Fact]
    public void Read_refuses_when_the_configuration_sheet_is_missing()
    {
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        var sheets = new Mock<Excel.Sheets>();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.SetupGet(w => w.Sheets).Returns(sheets.Object);
        _ = sheets.SetupGet(s => s.Count).Returns(0);

        Assert.Equal(
            ConfigReadOutcome.Refused(ConfigReadRefusalReason.ConfigSheetMissing),
            new ExcelConfigCatalogueReader(application.Object).Read());
    }

    [Fact]
    public void Read_refuses_when_a_catalogue_table_is_missing()
    {
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        fake.Tables.RemoveAt(2);

        var outcome = ConfigGraph.BuildReader(fake).Read();

        Assert.Equal(ConfigReadOutcome.Refused(ConfigReadRefusalReason.TableMissing), outcome);
    }

    [Fact]
    public void Read_refuses_when_a_header_is_renamed()
    {
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        fake.Tables[0].Headers[1] = "TypeNameRenamed";

        var outcome = ConfigGraph.BuildReader(fake).Read();

        Assert.Equal(ConfigReadOutcome.Refused(ConfigReadRefusalReason.HeaderMismatch), outcome);
    }

    [Fact]
    public void Read_refuses_when_a_catalogue_table_has_an_extra_row()
    {
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        fake.Tables[2].Body.Insert(0, new object?[] { "SmuggledMetric", 1.0, 0.0, 2.0 });

        var outcome = ConfigGraph.BuildReader(fake).Read();

        Assert.Equal(ConfigReadOutcome.Refused(ConfigReadRefusalReason.RowCountMismatch), outcome);
    }

    [Fact]
    public void Read_refuses_when_a_type_row_drifts_from_the_code_owned_catalogue()
    {
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        fake.Tables[0].Body[3][1] = "Renamed Display Name";

        var outcome = ConfigGraph.BuildReader(fake).Read();

        Assert.Equal(ConfigReadOutcome.Refused(ConfigReadRefusalReason.CatalogueMismatch), outcome);
    }

    [Fact]
    public void Read_refuses_when_a_metric_value_leaves_its_range()
    {
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        var index = GanttCatalogues.Metrics.ToList().FindIndex(metric => metric.Name == "LaneHeightPt");
        fake.Tables[2].Body[index][1] = 999.0;

        var outcome = ConfigGraph.BuildReader(fake).Read();

        Assert.Equal(ConfigReadOutcome.Refused(ConfigReadRefusalReason.ValueOutOfRange), outcome);
    }

    [Fact]
    public void Read_refuses_when_a_style_colour_is_not_uppercase_rrggbb()
    {
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        fake.Tables[1].Body[0][2] = "red";

        var outcome = ConfigGraph.BuildReader(fake).Read();

        Assert.Equal(ConfigReadOutcome.Refused(ConfigReadRefusalReason.BadColourFormat), outcome);
    }

    [Fact]
    public void Read_refuses_when_the_stored_hash_is_stale()
    {
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        fake.Tables[4].Body[1][1] = new string('0', 64);

        var outcome = ConfigGraph.BuildReader(fake).Read();

        Assert.Equal(ConfigReadOutcome.Refused(ConfigReadRefusalReason.CatalogueHashMismatch), outcome);
    }

    [Fact]
    public void Read_refuses_duplicate_config_metadata_keys()
    {
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        fake.Tables[4].Body[1] = [GanttCatalogues.ConfigSchemaVersionKey, fake.Tables[4].Body[1][1]];

        var outcome = ConfigGraph.BuildReader(fake).Read();

        Assert.Equal(ConfigReadOutcome.Refused(ConfigReadRefusalReason.CatalogueMismatch), outcome);
    }

    [Fact]
    public void Read_refuses_missing_config_metadata_keys()
    {
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        fake.Tables[4].Body.RemoveAt(3);

        var outcome = ConfigGraph.BuildReader(fake).Read();

        Assert.Equal(ConfigReadOutcome.Refused(ConfigReadRefusalReason.RowCountMismatch), outcome);
    }

    [Fact]
    public void Read_refuses_unexpected_config_metadata_keys()
    {
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        fake.Tables[4].Body[3] = ["Unexpected", "value"];

        var outcome = ConfigGraph.BuildReader(fake).Read();

        Assert.Equal(ConfigReadOutcome.Refused(ConfigReadRefusalReason.CatalogueMismatch), outcome);
    }

    [Fact]
    public void Read_refuses_a_boolean_setting_is_not_true_or_false()
    {
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        var index = GanttCatalogues.Settings.ToList().FindIndex(setting => setting.Key == "ShowTitle");
        fake.Tables[3].Body[index][1] = "yes";

        var outcome = ConfigGraph.BuildReader(fake).Read();

        Assert.Equal(ConfigReadOutcome.Refused(ConfigReadRefusalReason.ValueOutOfRange), outcome);
    }

    [Theory]
    [InlineData("Month", "MM")]
    [InlineData("Month", "MMM")]
    [InlineData("Quarter", "Quarter")]
    [InlineData("Year", "Year")]
    public void Read_accepts_each_supported_scale_and_period_format_pair(string scale, string format)
    {
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        var scaleIndex = GanttCatalogues.Settings.ToList().FindIndex(setting => setting.Key == "TimeScale");
        var formatIndex = GanttCatalogues.Settings.ToList().FindIndex(setting => setting.Key == "PeriodLabelFormat");
        fake.Tables[3].Body[scaleIndex][1] = scale;
        fake.Tables[3].Body[formatIndex][1] = format;

        var outcome = ConfigGraph.BuildReader(fake).Read();

        Assert.True(outcome.Succeeded);
        Assert.Equal(scale, outcome.Settings["TimeScale"]);
        Assert.Equal(format, outcome.Settings["PeriodLabelFormat"]);
    }

    [Fact]
    public void Read_refuses_a_blank_chart_title()
    {
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        var index = GanttCatalogues.Settings.ToList().FindIndex(setting => setting.Key == "ChartTitle");
        fake.Tables[3].Body[index][1] = "   ";

        var outcome = ConfigGraph.BuildReader(fake).Read();

        Assert.Equal(ConfigReadOutcome.Refused(ConfigReadRefusalReason.ValueOutOfRange), outcome);
    }

    [Theory]
    [InlineData("TimeScale", "Day")]
    [InlineData("PeriodLabelFormat", "MMM ")]
    public void Read_refuses_unknown_chart_settings(string key, string value)
    {
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        var index = GanttCatalogues.Settings.ToList().FindIndex(setting => setting.Key == key);
        fake.Tables[3].Body[index][1] = value;

        var outcome = ConfigGraph.BuildReader(fake).Read();

        Assert.Equal(ConfigReadOutcome.Refused(ConfigReadRefusalReason.ValueOutOfRange), outcome);
    }

    [Fact]
    public void Read_refuses_an_incompatible_scale_and_period_format()
    {
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        var index = GanttCatalogues.Settings.ToList().FindIndex(setting => setting.Key == "PeriodLabelFormat");
        fake.Tables[3].Body[index][1] = "Quarter";

        var outcome = ConfigGraph.BuildReader(fake).Read();

        Assert.Equal(ConfigReadOutcome.Refused(ConfigReadRefusalReason.ValueOutOfRange), outcome);
    }

    [Fact]
    public void Read_returns_the_present_setting_values_not_just_the_defaults()
    {
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        var index = GanttCatalogues.Settings.ToList().FindIndex(setting => setting.Key == "LegendPosition");
        fake.Tables[3].Body[index][1] = "Bottom";

        var outcome = ConfigGraph.BuildReader(fake).Read();

        Assert.True(outcome.Succeeded);
        Assert.Equal("Bottom", outcome.Settings["LegendPosition"]);
    }

    [Fact]
    public void Read_iterates_one_based_body_matrices_like_live_excel()
    {
        // Live DataBodyRange.Value2 is a one-based SAFEARRAY: the R2.7 live
        // office gate crashed on matrix[0, ...] inside a 0-based-only read
        // loop, while ExcelGanttTableReader's contract fake already pins
        // Array.CreateInstance(..., [1, 1]) for the same shape. This pins
        // both halves — the config fake serves one-based matrices and the
        // reader converts them — so the bug stays a contract-test failure.
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        var served = Assert.IsType<object[,]>(fake.Tables[0].BodyMock.Object.Value2);
        Assert.Equal(1, served.GetLowerBound(0));
        Assert.Equal(1, served.GetLowerBound(1));

        Assert.True(ConfigGraph.BuildReader(fake).Read().Succeeded);
    }

    [Fact]
    public void Read_accepts_live_coerced_boolean_setting_cells()
    {
        // Live probe (2026-09-22, Microsoft 365 x64): Value2 coerces the
        // string "TRUE" on write and reads back a Boolean, whose invariant
        // text is "True" — not "TRUE". The R2.7a office gate refused with
        // ValueOutOfRange until ToText normalised bool cells to the ADR-0007
        // D3 canonical "TRUE"/"FALSE" (the shape CellMatches already
        // handles). Contract fakes keep raw strings, so only an injected
        // Boolean fixture exercises the live cell shape.
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        var settings = GanttCatalogues.Settings.ToList();
        fake.Tables[3].Body[settings.FindIndex(s => s.Key == "ShowTitle")][1] = true;
        fake.Tables[3].Body[settings.FindIndex(s => s.Key == "ExportIncludeLegend")][1] = false;

        var outcome = ConfigGraph.BuildReader(fake).Read();

        Assert.True(outcome.Succeeded);
        Assert.Equal("TRUE", outcome.Settings["ShowTitle"]);
        Assert.Equal("FALSE", outcome.Settings["ExportIncludeLegend"]);
    }
    [Fact]
    public void Read_accepts_a_user_style_with_explicit_valid_capabilities()
    {
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        fake.Tables[1].Body.Add(UserStyleRow());

        Assert.True(ConfigGraph.BuildReader(fake).Read().Succeeded);
    }

    [Theory]
    [InlineData("Sideways", "Auto", "Fill")]
    [InlineData("Auto", "", "Fill")]
    [InlineData("Auto", "Sideways", "Fill")]
    [InlineData("Auto", "Auto Auto", "Fill")]
    [InlineData("Inside", "Auto", "Fill")]
    [InlineData("Auto", "Auto", "16")]
    public void Read_refuses_when_a_user_style_capability_is_invalid(
        string defaultPosition,
        string allowedPositions,
        string colourCapability)
    {
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        object?[] userStyle = UserStyleRow();
        userStyle[11] = defaultPosition;
        userStyle[12] = allowedPositions;
        userStyle[13] = colourCapability;
        fake.Tables[1].Body.Add(userStyle);

        Assert.Equal(
            ConfigReadOutcome.Refused(ConfigReadRefusalReason.ValueOutOfRange),
            ConfigGraph.BuildReader(fake).Read());
    }

    private static object?[] UserStyleRow() =>
    [
        "UserStyle",
        "User Style",
        "#112233",
        "#445566",
        "None",
        4.0,
        0.5,
        "#FFFFFF",
        0.75,
        8.0,
        0.0,
        "Inside",
        "Auto Inside",
        "Fill, Stroke",
    ];

    [Fact]
    public void Read_projects_a_user_style_as_renderable_formatting()
    {
        // R2.7c: a user row that validates must reach the registry with resolved
        // formatting, not capabilities only, or the resolver refuses it later.
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        fake.Tables[1].Body.Add(UserStyleRow());

        ConfigReadOutcome outcome = ConfigGraph.BuildReader(fake).Read();

        Assert.True(outcome.Succeeded);
        Assert.True(outcome.Styles.TryGet("UserStyle", out GanttStyleDefinition? style));
        Assert.True(style!.HasFormatting);
        Assert.Equal("#112233", style.FillColour);
        Assert.Equal("#445566", style.StrokeColour);
        Assert.Equal("#FFFFFF", style.TextColour);
        Assert.Equal(GanttLabelPosition.Inside, style.DefaultLabelPosition);
        Assert.Equal(8, style.ActivityHeightPt);
    }

    [Theory]
    [InlineData(2, "#ff0000")]
    [InlineData(3, "112233")]
    [InlineData(7, "#GGGGGG")]
    [InlineData(2, "#11223344")]
    [InlineData(3, " #112233")]
    [InlineData(7, "#112233 ")]
    public void Read_refuses_a_user_style_with_a_malformed_resolved_colour(int column, string value)
    {
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        object?[] userStyle = UserStyleRow();
        userStyle[column] = value;
        fake.Tables[1].Body.Add(userStyle);

        Assert.Equal(
            ConfigReadOutcome.Refused(ConfigReadRefusalReason.ValueOutOfRange),
            ConfigGraph.BuildReader(fake).Read());
    }

    [Theory]
    [InlineData(5)]
    [InlineData(6)]
    [InlineData(8)]
    [InlineData(9)]
    [InlineData(10)]
    public void Read_refuses_a_user_style_with_a_negative_metric(int column)
    {
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        object?[] userStyle = UserStyleRow();
        userStyle[column] = -1.0;
        fake.Tables[1].Body.Add(userStyle);

        Assert.Equal(
            ConfigReadOutcome.Refused(ConfigReadRefusalReason.ValueOutOfRange),
            ConfigGraph.BuildReader(fake).Read());
    }

    [Fact]
    public void Read_refuses_a_user_style_with_a_non_numeric_metric()
    {
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        object?[] userStyle = UserStyleRow();
        userStyle[9] = "tall";
        fake.Tables[1].Body.Add(userStyle);

        Assert.Equal(
            ConfigReadOutcome.Refused(ConfigReadRefusalReason.ValueOutOfRange),
            ConfigGraph.BuildReader(fake).Read());
    }

    [Fact]
    public void Read_refuses_a_user_style_with_an_undefined_hatch_pattern()
    {
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        object?[] userStyle = UserStyleRow();
        userStyle[4] = "Diagonalish";
        fake.Tables[1].Body.Add(userStyle);

        Assert.Equal(
            ConfigReadOutcome.Refused(ConfigReadRefusalReason.ValueOutOfRange),
            ConfigGraph.BuildReader(fake).Read());
    }
}
