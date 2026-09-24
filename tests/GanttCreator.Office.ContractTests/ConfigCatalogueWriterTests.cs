using Excel = Microsoft.Office.Interop.Excel;
using GanttCreator.Core;
using Moq;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for <see cref="ExcelConfigCatalogueWriter"/>. They need no
/// live Office: the application object is absent, foreign, or a Moq proxy of
/// the Excel PIA interfaces, and the configuration worksheet is the
/// <see cref="ConfigSheetFake"/> graph. Every refusal validator has a
/// positive test that constructs the bad input and asserts the typed refusal
/// fires with no mutation (AGENTS.md validator rule); the D4 preservation
/// tests prove user-authored style rows and present setting values survive
/// regeneration.
/// </summary>
public class ConfigCatalogueWriterTests
{
    [Fact]
    public void Write_creates_the_five_tables_with_the_contract_names_headers_and_counts()
    {
        var fake = new ConfigSheetFake();
        var writer = ConfigGraph.BuildWriter(fake);

        var outcome = writer.Write();

        Assert.Equal(ConfigWriteOutcome.Ok(), outcome);
        Assert.Equal(5, fake.Tables.Count);

        AssertContractTable(
            fake.Tables[0],
            GanttCatalogues.TypesHeaders,
            GanttCatalogues.TypeRows.Count);
        AssertContractTable(
            fake.Tables[1],
            GanttCatalogues.StylesHeaders,
            GanttCatalogues.StylePresets.Count);
        AssertContractTable(
            fake.Tables[2],
            GanttCatalogues.MetricsHeaders,
            GanttCatalogues.Metrics.Count);
        AssertContractTable(
            fake.Tables[3],
            GanttCatalogues.SettingsHeaders,
            GanttCatalogues.Settings.Count);
        AssertContractTable(
            fake.Tables[4],
            GanttCatalogues.ConfigHeaders,
            4);
    }

    [Fact]
    public void Write_materialises_the_types_from_the_code_owned_catalogue_only()
    {
        var fake = new ConfigSheetFake();
        var writer = ConfigGraph.BuildWriter(fake);

        _ = writer.Write();

        ConfigSheetFake.TableFake types = fake.Tables[0];
        for (var index = 0; index < GanttCatalogues.TypeRows.Count; index++)
        {
            GanttTypeCatalogueRow expected = GanttCatalogues.TypeRows[index];
            Assert.Equal(expected.TypeName, CellText(types.Body[index][0]));
            Assert.Equal(expected.DisplayName, CellText(types.Body[index][1]));
            Assert.Equal(expected.DefaultStyleKey, CellText(types.Body[index][4]));
        }
    }

    [Fact]
    public void Write_writes_a_stable_workbook_id_and_the_current_catalogue_hash()
    {
        var fake = new ConfigSheetFake();
        var writer = ConfigGraph.BuildWriter(fake);

        _ = writer.Write();

        ConfigSheetFake.TableFake config = fake.Tables[4];
        Assert.Equal(
            GanttSchemaVersion.CurrentSchemaVersion.ToString(System.Globalization.CultureInfo.InvariantCulture),
            CellText(config.Body[0][1]));
        Assert.Equal(GanttCatalogues.ComputeCatalogueHash(), CellText(config.Body[1][1]));
        Assert.True(Guid.TryParse(CellText(config.Body[2][1]), out _));
        Assert.Equal(VersionInfo.SemanticVersion, CellText(config.Body[3][1]));
    }

    [Fact]
    public void Write_keeps_the_workbook_id_stable_across_regeneration()
    {
        var first = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(first).Write();
        var second = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(second).Write();

        var firstId = CellText(first.Tables[4].Body[2][1]);
        var secondId = CellText(second.Tables[4].Body[2][1]);
        Assert.False(string.Equals(firstId, secondId, StringComparison.Ordinal));
    }

    [Fact]
    public void Write_reports_no_active_workbook_without_an_application_object() =>
        Assert.Equal(
            ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.NoActiveWorkbook),
            new ExcelConfigCatalogueWriter(null).Write());

    [Fact]
    public void Write_degrades_to_no_active_workbook_for_a_foreign_object() =>
        Assert.Equal(
            ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.NoActiveWorkbook),
            new ExcelConfigCatalogueWriter("not an Excel application").Write());

    [Fact]
    public void Write_reports_no_active_workbook_when_excel_has_no_active_workbook()
    {
        var application = new Mock<Excel.Application>();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns((Excel.Workbook)null!);

        Assert.Equal(
            ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.NoActiveWorkbook),
            new ExcelConfigCatalogueWriter(application.Object).Write());
    }

    [Fact]
    public void Write_refuses_when_the_configuration_sheet_is_missing()
    {
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        var sheets = new Mock<Excel.Sheets>();
        var activeSheet = new Mock<Excel._Worksheet>();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        _ = workbook.SetupGet(w => w.Sheets).Returns(sheets.Object);
        _ = workbook.SetupGet(w => w.ActiveSheet).Returns(activeSheet.Object);
        _ = activeSheet.SetupGet(w => w.ProtectContents).Returns(false);
        _ = sheets.SetupGet(s => s.Count).Returns(0);

        Assert.Equal(
            ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.ConfigSheetMissing),
            new ExcelConfigCatalogueWriter(application.Object).Write());
    }

    [Fact]
    public void Write_refuses_with_no_mutation_when_the_configuration_sheet_is_protected()
    {
        var fake = new ConfigSheetFake(protectContents: true);
        var writer = ConfigGraph.BuildWriter(fake);

        var outcome = writer.Write();

        Assert.Equal(
            ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.TargetProtected),
            outcome);
        Assert.Empty(fake.Tables);
    }

    [Fact]
    public void Write_refuses_with_no_mutation_when_the_workbook_structure_is_protected()
    {
        var fake = new ConfigSheetFake();
        var writer = ConfigGraph.BuildWriter(fake, structureProtected: true);

        var outcome = writer.Write();

        Assert.Equal(
            ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.TargetProtected),
            outcome);
        Assert.Empty(fake.Tables);
    }

    [Theory]
    [InlineData(ProtectionGuardOutcome.SheetProtected)]
    [InlineData(ProtectionGuardOutcome.WorkbookStructureProtected)]
    public void Write_shared_guard_refusal_does_not_mutate(ProtectionGuardOutcome protection)
    {
        var fake = new ConfigSheetFake();
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(protection);
        var writer = ConfigGraph.BuildWriter(fake, protectionGuard: guard.Object);

        var outcome = writer.Write();

        Assert.Equal(ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.TargetProtected), outcome);
        guard.Verify(g => g.Query(), Times.Once);
        Assert.Empty(fake.Tables);
    }

    [Fact]
    public void Write_preserves_user_style_rows_and_present_settings_on_regeneration()
    {
        // D4: the same writer regenerates built-ins without destroying user
        // content — a user style row and a present setting value must
        // survive a second Write.
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        var reader = ConfigGraph.BuildReader(fake);
        ConfigReadOutcome before = reader.Read();
        Assert.True(before.Succeeded);
        object?[] userStyle =
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
        fake.Tables[1].Body.Add(userStyle);

        Assert.True(ConfigGraph.BuildReader(fake).Read().Succeeded);
        Assert.Equal(16, fake.Tables[1].Body.Count);

        var outcome = ConfigGraph.BuildWriter(fake).Write();

        Assert.Equal(ConfigWriteOutcome.Ok(), outcome);
        Assert.Equal(5, fake.Tables.Count);
        Assert.Equal(16, fake.Tables[1].Body.Count);
        Assert.Equal(userStyle, fake.Tables[1].Body[^1]);
        Assert.True(ConfigGraph.BuildReader(fake).Read().Succeeded);
    }

    [Fact]
    public void Write_regenerates_exactly_the_approved_setting_keys()
    {
        // ADR-0007 D4: the settings table always regenerates to exactly the
        // approved key set, so an unapproved (dialog-authored) key is not
        // carried forward — preserving it would leave the workbook
        // unreadable after every regeneration, because the reader validates
        // the exact key set.
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        fake.Tables[3].Body.Add(["UnapprovedKey", "present"]);

        var outcome = ConfigGraph.BuildWriter(fake).Write();

        Assert.Equal(ConfigWriteOutcome.Ok(), outcome);
        Assert.Equal(GanttCatalogues.Settings.Count, fake.Tables[3].Body.Count);
        Assert.DoesNotContain(
            fake.Tables[3].Body,
            row => CellText(row.ElementAtOrDefault(0)) == "UnapprovedKey");
        Assert.True(ConfigGraph.BuildReader(fake).Read().Succeeded);
    }

    [Fact]
    public void Write_emits_the_style_columns_in_the_styles_header_order()
    {
        // Layout pin. The numeric columns are non-contiguous: TextColour (7)
        // sits between HatchLinePt (6) and StandardOutlinePt (8), and
        // MilestoneSizePt (10) closes the row. A reader that indexes the
        // numbers contiguously from 5 silently compares TextColour against
        // StandardOutlinePt, so this order is asserted column by column.
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();

        for (var index = 0; index < GanttCatalogues.StylePresets.Count; index++)
        {
            GanttStylePreset preset = GanttCatalogues.StylePresets[index];
            object?[] row = fake.Tables[1].Body[index];
            Assert.Equal(preset.StyleKey, CellText(row[0]));
            Assert.Equal(preset.DisplayName, CellText(row[1]));
            Assert.Equal(preset.FillColour, CellText(row[2]));
            Assert.Equal(preset.StrokeColour, CellText(row[3]));
            Assert.Equal(preset.HatchPattern.ToString(), CellText(row[4]));
            Assert.Equal(preset.HatchPitchPt, Assert.IsType<double>(row[5]));
            Assert.Equal(preset.HatchLinePt, Assert.IsType<double>(row[6]));
            Assert.Equal(preset.TextColour, CellText(row[7]));
            Assert.Equal(preset.StandardOutlinePt, Assert.IsType<double>(row[8]));
            Assert.Equal(preset.ActivityHeightPt, Assert.IsType<double>(row[9]));
            Assert.Equal(preset.MilestoneSizePt, Assert.IsType<double>(row[10]));
            Assert.Equal(preset.DefaultLabelPosition.ToString(), CellText(row[11]));
            Assert.Equal(
                string.Join(
                    " ",
                    preset.AllowedLabelPositions
                        .Select(position => position.ToString())
                        .OrderBy(name => name, StringComparer.Ordinal)),
                CellText(row[12]));
            Assert.Equal(preset.ColourCapability.ToString(), CellText(row[13]));
        }
    }

    [Fact]
    public void Write_preserves_a_present_approved_setting_value_on_regeneration()
    {
        // The complement of the previous test: an approved key's present
        // value survives, so regeneration is not a reset to defaults.
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        var index = GanttCatalogues.Settings.ToList().FindIndex(setting => setting.Key == "LegendPosition");
        fake.Tables[3].Body[index][1] = "Bottom";

        var outcome = ConfigGraph.BuildWriter(fake).Write();

        Assert.Equal(ConfigWriteOutcome.Ok(), outcome);
        Assert.Equal("Bottom", CellText(fake.Tables[3].Body[index][1]));
        ConfigReadOutcome read = ConfigGraph.BuildReader(fake).Read();
        Assert.True(read.Succeeded);
        Assert.Equal("Bottom", read.Settings["LegendPosition"]);
    }

    [Fact]
    public void Write_regenerates_from_one_based_body_matrices()
    {
        // The regeneration path re-reads existing table bodies (D4
        // preservation, workbook-id stability) through the same Value2 shape
        // the reader sees: one-based lower bounds, live-proven. The first
        // Write seeds the tables; the second must read them back through
        // ConvertMatrix without IndexOutOfRangeException and leave the
        // catalogues readable.
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();

        var outcome = ConfigGraph.BuildWriter(fake).Write();

        Assert.Equal(ConfigWriteOutcome.Ok(), outcome);
        Assert.True(ConfigGraph.BuildReader(fake).Read().Succeeded);
    }

    [Fact]
    public void Write_preserves_live_coerced_boolean_settings_as_canonical_true_false_text()
    {
        // Live probe (2026-09-22): Value2 reads boolean setting cells back
        // as Boolean. D4 preservation runs those cells through the writer's
        // ToText, which must emit the canonical "TRUE"/"FALSE" (ADR-0007
        // D3) — otherwise the regenerated table stores "True" and the reader
        // refuses it, breaking read-after-regenerate.
        var fake = new ConfigSheetFake();
        _ = ConfigGraph.BuildWriter(fake).Write();
        var index = GanttCatalogues.Settings.ToList().FindIndex(s => s.Key == "ShowTitle");
        fake.Tables[3].Body[index][1] = true;

        var outcome = ConfigGraph.BuildWriter(fake).Write();

        Assert.Equal(ConfigWriteOutcome.Ok(), outcome);
        Assert.Equal("TRUE", CellText(fake.Tables[3].Body[index][1]));
        Assert.True(ConfigGraph.BuildReader(fake).Read().Succeeded);
    }

    private static void AssertContractTable(
        ConfigSheetFake.TableFake table,
        string[] expectedHeaders,
        int expectedRowCount)
    {
        Assert.Equal(expectedRowCount, table.Body.Count);
        Assert.Equal(expectedHeaders, table.Headers);
    }

    private static string CellText(object? cell) =>
        cell switch
        {
            null => string.Empty,
            string text => text,
            bool flag => flag ? "TRUE" : "FALSE",
            _ => Convert.ToString(cell, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty,
        };
}
