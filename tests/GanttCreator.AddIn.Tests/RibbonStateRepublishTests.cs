using GanttCreator.Core;
using GanttCreator.Office;
using Moq;
using static GanttCreator.AddIn.Tests.RibbonStateTestSupport;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// The ribbon is a projection of the table, never of the render (owner ruling 2026-11-27):
/// every completed pass and every suspended burst republishes the snapshot, so the
/// cell-aware and data-gated controls describe the sheet whether or not a chart was built.
/// </summary>
/// <remarks>
/// The class shares the <c>diagnostics-service</c> collection because
/// <see cref="RibbonStateService"/> is a session singleton; xUnit runs a collection's
/// classes sequentially, so the host lifecycle tests cannot race these.
/// </remarks>
[Collection("diagnostics-service")]
public sealed class RibbonStateRepublishTests : IDisposable
{
    public RibbonStateRepublishTests()
    {
        RibbonStateService.Reset();
    }

    public void Dispose()
    {
        RibbonStateService.Reset();
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// A change that adds a child row re-asks the data-gated controls. This is the case the
    /// seven-control invalidation list could never cover, and the facts block is what makes
    /// it possible.
    /// </summary>
    [Fact]
    public void A_table_structure_change_republishes_the_data_gated_controls()
    {
        GanttRowDto parent = Row(1, "Activity (Fixed Duration)");
        GanttRowDto child = Row(2, "Activity (Fixed Duration)", parent.Id);
        Mock<IGanttTableReader> reader = new();
        _ = reader.SetupSequence(item => item.Read())
            .Returns(GanttTableReadOutcome.Ok([parent]))
            .Returns(GanttTableReadOutcome.Ok([parent, child]));

        RibbonStateService service = ArmReader(reader.Object);
        service.Refresh();
        Assert.False(service.GetEnabled(RibbonControlIds.CollapseAllGroups));

        // The child arrives through an edit while the render is suspended (switch off or an
        // unlocked chart); the projection must still see it.
        service.RefreshAfterTableChange();

        Assert.True(service.GetEnabled(RibbonControlIds.ExpandAllGroups));
        Assert.True(service.GetEnabled(RibbonControlIds.CollapseAllGroups));
    }

    /// <summary>
    /// A Type edit moves the custom-style dropdown between enabled and disabled, because the
    /// new Type's catalogue definition decides whether it requires a style key.
    /// </summary>
    [Fact]
    public void A_type_change_republishes_the_cell_aware_controls()
    {
        string builtInKey = GanttCatalogues.StylePresets[0].StyleKey;
        List<GanttStyleDefinition> definitions =
        [
            new GanttStyleDefinition(
                builtInKey,
                new HashSet<GanttLabelPosition>([GanttLabelPosition.Auto]),
                EntityColourCapability.Fill),
            new GanttStyleDefinition(
                "MyStyle",
                new HashSet<GanttLabelPosition>([GanttLabelPosition.Auto]),
                EntityColourCapability.StyleDefined),
        ];
        GanttStyleRegistry registry = new(definitions);

        string type = "Activity (Fixed Duration)";
        Mock<IGanttTableReader> reader = new();
        _ = reader.Setup(item => item.Read())
            .Returns(() => GanttTableReadOutcome.Ok([Row(1, type)]));

        RibbonStateService service = RibbonStateService.Instance;
        service.SetApplicationAdapter(new TestExcelApplicationAdapter { ActiveWorkbook = true });
        service.SetSelectionIdentifier(new TestSelectionIdentifier
        {
            Facts = new SelectionFacts(true, true, 1, false, 1),
        });
        service.SetTableReader(reader.Object);

        Mock<IConfigCatalogueReader> catalogues = new();
        _ = catalogues.Setup(item => item.Read()).Returns(
            ConfigReadOutcome.Ok(
                "W-1",
                new Dictionary<string, string>(),
                registry,
                new Dictionary<string, double>()));
        service.SetCatalogueReader(catalogues.Object);

        service.Refresh();
        Assert.False(service.GetEnabled(RibbonControlIds.CustomStyle));

        // The user changes the row's Type cell. The render may be suspended; the projection
        // is not, and it must resolve the NEW type's catalogue definition.
        type = "Custom Activity";
        service.RefreshAfterTableChange();

        Assert.True(service.GetEnabled(RibbonControlIds.CustomStyle));
    }

    /// <summary>
    /// The republish takes the range the pass rendered and never re-derives it from its own
    /// read, which is the second derivation the outcome's plot fields exist to prevent.
    /// </summary>
    [Fact]
    public void The_republish_applies_the_outcome_dates_without_re_deriving_them()
    {
        DateOnly start = new(2030, 1, 1);
        DateOnly finish = new(2030, 1, 31);

        RibbonStateService service = RibbonStateService.Instance;
        service.SetApplicationAdapter(new TestExcelApplicationAdapter { ActiveWorkbook = true });

        service.RefreshAfterTableChange(start, finish);

        Assert.Equal(GanttDateFormatting.FormatDdMMMyy(start), service.GetPlotStartDate());
        Assert.Equal(GanttDateFormatting.FormatDdMMMyy(finish), service.GetPlotFinishDate());
    }

    /// <summary>
    /// A stored label token the new type does not offer reads as Default instead of an
    /// out-of-list selection (R5.23).
    /// </summary>
    [Fact]
    public void A_label_token_outside_the_offered_set_reads_as_default()
    {
        RibbonStateService service = ArmWithLabels(
            "As-Planned Activity", labelPosition: "TopLeft", dateLabelPosition: null);

        Assert.True(service.GetDescriptionLabelCount() > 0);
        Assert.Equal(0, service.GetDescriptionLabelSelectedIndex());
    }

    /// <summary>
    /// The same rule for the date-label vocabulary: a span's <c>Both</c> on a milestone
    /// reads as Default.
    /// </summary>
    [Fact]
    public void A_date_label_choice_outside_the_offered_set_reads_as_default()
    {
        RibbonStateService service = ArmWithLabels(
            "As-Planned Milestone", labelPosition: null, dateLabelPosition: "Both");

        Assert.True(service.GetDateLabelCount() > 0);
        Assert.Equal(0, service.GetDateLabelSelectedIndex());
    }

    private static RibbonStateService ArmWithLabels(
        string type,
        string? labelPosition,
        string? dateLabelPosition)
    {
        Mock<IGanttTableReader> reader = new();
        _ = reader.Setup(item => item.Read()).Returns(
            GanttTableReadOutcome.Ok([RowWithLabels(type, labelPosition, dateLabelPosition)]));

        RibbonStateService service = RibbonStateService.Instance;
        service.SetApplicationAdapter(new TestExcelApplicationAdapter { ActiveWorkbook = true });
        service.SetSelectionIdentifier(new TestSelectionIdentifier
        {
            Facts = new SelectionFacts(true, true, 1, false, 1),
        });
        service.SetTableReader(reader.Object);
        service.Refresh();
        return service;
    }

    private static GanttRowDto RowWithLabels(
        string type,
        string? labelPosition,
        string? dateLabelPosition) =>
        new(
            1,
            NewRowId('a'),
            null,
            0,
            type,
            null,
            null,
            null,
            null,
            null,
            labelPosition,
            null,
            null,
            true,
            null,
            dateLabelPosition);

    /// <summary>Arms the service with one table reader over the supplied substitute.</summary>
    /// <param name="reader">The table reader.</param>
    /// <returns>The armed session service.</returns>
    private static RibbonStateService ArmReader(IGanttTableReader reader)
    {
        RibbonStateService service = RibbonStateService.Instance;
        service.SetApplicationAdapter(new TestExcelApplicationAdapter { ActiveWorkbook = true });
        service.SetTableReader(reader);
        return service;
    }
}
